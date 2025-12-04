#!/usr/bin/env python3
from __future__ import annotations

import json
import platform
import re
import subprocess
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from openpyxl import Workbook  # type: ignore

APP_TITLE = "RAM Barcode Scanner"
MAPPINGS_JSON = Path(__file__).with_name("user_mappings.json")
RESULTS_JSON = Path(__file__).with_name("scan_results.json")

SUCCESS_WAV = Path(__file__).with_name("success.wav")
UNKNOWN_WAV = Path(__file__).with_name("unknown.wav")

# only last N results shown in UI (all still saved in JSON)
MAX_VISIBLE_RESULTS = 10

COLOR_SIZE = "#1f77b4"
COLOR_SPEED = "#2ca02c"
COLOR_TYPE = "#d62728"
COLOR_CLASS = "#ff7f0e"
COLOR_ECC = "#17becf"
COLOR_MFR = "#9467bd"
COLOR_VERSION = "#8c564b"
COLOR_BARCODE = "#555555"


# -------------------- Data models --------------------
@dataclass
class Mapping:
    pattern: str
    size_gb: Optional[int]
    speed_mts: Optional[int]
    mem_type: Optional[str]
    manufacturer: Optional[str]
    module_class: Optional[str] = None
    ecc: Optional[bool] = None
    regex: bool = False

    def matches(self, code: str) -> bool:
        if self.regex:
            try:
                return re.search(self.pattern, code) is not None
            except re.error:
                return False
        return self.pattern == code


@dataclass
class ScanResult:
    id: int
    timestamp: str
    barcode: str           # canonical barcode key
    size_gb: int
    speed_mts: int
    mem_type: str
    module_class: Optional[str]
    ecc: Optional[bool]
    manufacturer: str
    version: int           # permanently bound to that barcode


# -------------------- Mapping store --------------------
class MappingStore:
    """
    Combines mappings + permanent version numbers in a single JSON file.

    File layout:
    {
      "mappings": [
         {
           "pattern": "M391B2873FH0",
           "size_gb": 8,
           "speed_mts": 1600,
           "mem_type": "DDR3",
           "manufacturer": "Samsung",
           "module_class": "PC3-12800",
           "ecc": true,
           "regex": false
         },
         ...
      ],
      "versions": {
        "M391B2873FH0": 1,
        "HMT84GL7AMR4C": 2,
        ...
      }
    }
    """

    def __init__(self, path: Path):
        self.path = path
        self.mappings: List[Mapping] = []
        self.versions: Dict[str, int] = {}
        self.load()

    def load(self):
        if not self.path.exists():
            self.mappings = []
            self.versions = {}
            return

        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            self.mappings = []
            self.versions = {}
            return

        if isinstance(raw, list):
            # legacy format: just a list of mappings
            self.mappings = []
            for m in raw:
                m = dict(m)
                m.setdefault("module_class", None)
                m.setdefault("ecc", None)
                m.setdefault("regex", False)
                self.mappings.append(Mapping(**m))
            self.versions = {}
        elif isinstance(raw, dict):
            data = raw
            mappings_src = data.get("mappings", [])
            self.mappings = []
            for m in mappings_src:
                m = dict(m)
                m.setdefault("module_class", None)
                m.setdefault("ecc", None)
                m.setdefault("regex", False)
                self.mappings.append(Mapping(**m))
            versions_src = data.get("versions", {})
            self.versions = {str(k): int(v) for k, v in versions_src.items()}
        else:
            self.mappings = []
            self.versions = {}

    def save(self):
        data = {
            "mappings": [asdict(m) for m in self.mappings],
            "versions": self.versions,
        }
        self.path.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def add(self, m: Mapping):
        self.mappings.append(m)
        self.save()

    def remove_index(self, idx: int):
        try:
            del self.mappings[idx]
            self.save()
        except Exception:
            pass

    def find(self, code: str) -> Optional[Mapping]:
        # exact first
        for m in self.mappings:
            if not m.regex and m.matches(code):
                return m
        # regex after
        for m in self.mappings:
            if m.regex and m.matches(code):
                return m
        return None

    def all_descriptions(self) -> List[str]:
        out: List[str] = []
        for m in self.mappings:
            pat = f"/{m.pattern}/" if m.regex else m.pattern
            parts = [
                f"Pattern: {pat}",
                f"Size: {m.size_gb or '?'} GB",
                f"Speed: {m.speed_mts or '?'} MT/s",
                f"Type: {m.mem_type or '?'}",
                f"Class: {m.module_class or '?'}",
                f"ECC: {'Yes' if m.ecc else ('No' if m.ecc is not None else '?')}",
                f"Mfr: {m.manufacturer or '?'}",
            ]
            out.append(" | ".join(parts))
        return out


# -------------------- Heuristic parsing --------------------
MANUFACTURER_HINTS = {
    "samsung": "Samsung",
    "hynix": "SK Hynix",
    "skhynix": "SK Hynix",
    "micron": "Micron",
    "crucial": "Crucial",
    "kingston": "Kingston",
    "gskill": "G.SKILL",
    "g.skill": "G.SKILL",
    "adata": "ADATA",
    "corsair": "Corsair",
    "patriot": "Patriot",
    "teamgroup": "TeamGroup",
    "lexar": "Lexar",
    "ramaxel": "Ramaxel",
    "nanya": "Nanya",
}

TYPE_PATTERNS = {
    r"ddr\s*5": "DDR5",
    r"ddr\s*4": "DDR4",
    r"ddr\s*3": "DDR3",
    r"ddr\s*2": "DDR2",
    r"pc5": "DDR5",
    r"pc4": "DDR4",
    r"pc3": "DDR3",
    r"pc2": "DDR2",
}

SIZE_PATTERNS = [
    r"\b(\d+)\s*gb\b",
    r"\b(\d+)\s*g\b",
    r"\b(\d+)\s*x?\s*\d+gb\b",
]

SPEED_PATTERNS = [
    r"\b(\d{3,5})\s*mt\/?s\b",
    r"\b(\d{3,5})\s*mhz\b",
]

# PCn-XXXXX(L/E/R/U/...) pattern
MODULE_CLASS_RE = re.compile(
    r"pc(?P<gen>\d+)(?P<lv>l)?[- ]?(?P<rating>\d{3,6})(?P<suffix>[a-zA-Z]*)",
    re.IGNORECASE,
)

# Map common "PCx-14900" style ratings (MB/s) to MT/s speeds
COMMON_MB_TO_MTS = {
    6400: 800,
    8500: 1066,
    8533: 1066,
    10600: 1333,
    10660: 1333,
    10700: 1333,
    12800: 1600,
    14900: 1866,
    15000: 1866,
    16000: 2000,
    17000: 2133,
    17900: 2250,
    19200: 2400,
    21300: 2666,
    22400: 2800,
    23000: 2888,
    24000: 3000,
    25600: 3200,
    28800: 3600,
    32000: 4000,
}

# Inverse: MT/s -> canonical PC rating (MB/s)
SPEED_TO_MB: Dict[int, int] = {}
for mb, mts in COMMON_MB_TO_MTS.items():
    SPEED_TO_MB.setdefault(mts, mb)


def synth_module_class(
    mem_type: Optional[str],
    speed_mts: Optional[int],
    manual_kind: Optional[str] = None,
) -> Optional[str]:
    """
    Build a PCx-XXXXX style module class from manual DDR type + speed.
    Example: DDR3 + 1333 -> PC3-10600
    """
    if not mem_type or not speed_mts:
        return None

    m = mem_type.upper().replace(" ", "")
    m = re.sub(r"[^A-Z0-9]", "", m)
    if not m.startswith("DDR") or len(m) < 4 or not m[3].isdigit():
        return None

    gen = m[3]          # '3' from DDR3
    pc_prefix = f"PC{gen}"

    mb = SPEED_TO_MB.get(speed_mts)
    if mb is None:
        # Fallback: bandwidth = MT/s * 8 bytes -> MB/s, round to nearest 100
        mb = int(round(speed_mts * 8 / 100.0)) * 100

    return f"{pc_prefix}-{mb}"


def extract_module_class(text: str) -> Optional[str]:
    m = MODULE_CLASS_RE.search(text)
    if not m:
        return None
    gen = m.group("gen")
    lv = m.group("lv") or ""
    rating = m.group("rating")
    suffix = (m.group("suffix") or "").upper()
    prefix = f"PC{gen}{lv.upper()}"
    return f"{prefix}-{rating}{suffix}"


def parse_module_class(module_class: str) -> Tuple[Optional[str], Optional[int], Optional[bool]]:
    s = module_class.strip().upper()
    m = MODULE_CLASS_RE.search(s)
    if not m:
        return None, None, None

    gen = m.group("gen")
    rating = m.group("rating")
    suffix = (m.group("suffix") or "").upper()

    mem_type = {"2": "DDR2", "3": "DDR3", "4": "DDR4", "5": "DDR5"}.get(gen)

    speed: Optional[int] = None
    try:
        rating_int = int(rating)
        if rating_int in COMMON_MB_TO_MTS:
            speed = COMMON_MB_TO_MTS[rating_int]
        elif rating_int < 4000:
            speed = rating_int
        else:
            speed = max(1, round(rating_int / 8.0))
    except Exception:
        speed = None

    ecc: Optional[bool] = None
    if "E" in suffix or "Q" in suffix:
        ecc = True

    return mem_type, speed, ecc


def parse_barcode(
    code: str,
    store: MappingStore,
) -> Tuple[Optional[int], Optional[int], Optional[str], Optional[str], Optional[str], Optional[bool], bool]:
    """
    Return:
      size_gb, speed_mts, mem_type, manufacturer, module_class, ecc, parsed_ok
    """
    code_clean = code.strip()
    m = store.find(code_clean)
    size_gb: Optional[int] = None
    speed_mts: Optional[int] = None
    mem_type: Optional[str] = None
    manufacturer: Optional[str] = None
    module_class: Optional[str] = None
    ecc: Optional[bool] = None

    if m:
        size_gb = m.size_gb
        speed_mts = m.speed_mts
        mem_type = m.mem_type
        manufacturer = m.manufacturer
        module_class = m.module_class
        ecc = m.ecc
        if module_class:
            mc_type, mc_speed, mc_ecc = parse_module_class(module_class)
            if mc_type and not mem_type:
                mem_type = mc_type
            if mc_speed and not speed_mts:
                speed_mts = mc_speed
            if mc_ecc is not None and ecc is None:
                ecc = mc_ecc
        parsed_ok = all([size_gb, speed_mts, mem_type, manufacturer])
        return size_gb, speed_mts, mem_type, manufacturer, module_class, ecc, parsed_ok

    lower = code_clean.lower()

    for pat in SIZE_PATTERNS:
        mm = re.search(pat, lower)
        if mm:
            try:
                size_gb = int(mm.group(1))
                break
            except Exception:
                pass

    for pat, tval in TYPE_PATTERNS.items():
        if re.search(pat, lower):
            mem_type = tval
            break

    for pat in SPEED_PATTERNS:
        mm = re.search(pat, lower)
        if not mm:
            continue
        try:
            val = int(mm.group(1))
            speed_mts = val
            break
        except Exception:
            continue

    module_class = extract_module_class(lower)
    if module_class:
        mc_type, mc_speed, mc_ecc = parse_module_class(module_class)
        if mc_type and not mem_type:
            mem_type = mc_type
        if mc_speed and not speed_mts:
            speed_mts = mc_speed
        if mc_ecc is not None:
            ecc = mc_ecc

    for key, name in MANUFACTURER_HINTS.items():
        if key in lower:
            manufacturer = name
            break

    parsed_ok = all([size_gb, speed_mts, mem_type, manufacturer])
    return size_gb, speed_mts, mem_type, manufacturer, module_class, ecc, parsed_ok


# -------------------- Sound helpers --------------------
def play_wav_if_available(path: Path) -> bool:
    if not path.exists():
        return False
    try:
        system = platform.system()
        if system == "Windows":
            import winsound  # type: ignore

            winsound.PlaySound(str(path), winsound.SND_FILENAME | winsound.SND_ASYNC)
            return True
        elif system == "Darwin":
            subprocess.Popen(["afplay", str(path)])
            return True
        else:
            for cmd in ("paplay", "aplay", "ffplay"):
                try:
                    subprocess.Popen(
                        [cmd, str(path)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
                    return True
                except FileNotFoundError:
                    continue
    except Exception:
        pass
    return False


def play_success():
    if not play_wav_if_available(SUCCESS_WAV):
        root.bell()


def play_unknown():
    if play_wav_if_available(UNKNOWN_WAV):
        return
    system = platform.system()
    try:
        if system == "Windows":
            import winsound  # type: ignore

            try:
                winsound.MessageBeep(winsound.MB_ICONHAND)
                return
            except Exception:
                pass
            for f, d in [(900, 200), (500, 200), (900, 200)]:
                winsound.Beep(f, d)
            return
        elif system == "Darwin":
            for snd in ("Basso", "Sosumi", "Glass"):
                p = Path(f"/System/Library/Sounds/{snd}.aiff")
                if p.exists():
                    subprocess.Popen(["afplay", str(p)])
                    return
            root.bell()
            root.after(120, root.bell)
            root.after(240, root.bell)
            return
        else:
            candidates = [
                "/usr/share/sounds/freedesktop/stereo/dialog-error.oga",
                "/usr/share/sounds/freedesktop/stereo/alarm-clock-elapsed.oga",
                "/usr/share/sounds/freedesktop/stereo/complete.oga",
            ]
            for cmd in ("paplay", "aplay", "ffplay"):
                for snd in candidates:
                    if Path(snd).exists():
                        try:
                            subprocess.Popen(
                                [cmd, snd],
                                stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL,
                            )
                            return
                        except Exception:
                            continue
            root.bell()
            root.after(120, root.bell)
            root.after(240, root.bell)
            return
    except Exception:
        pass
    root.bell()
    root.after(120, root.bell)
    root.after(240, root.bell)


# -------------------- GUI App --------------------
class App(ttk.Frame):
    def __init__(self, master: tk.Tk, store: MappingStore):
        super().__init__(master)
        self.master = master
        self.store = store
        self.pack(fill=tk.BOTH, expand=True)

        # permanent barcode -> version mapping (from user_mappings.json)
        self.barcode_versions: Dict[str, int] = dict(self.store.versions)

        # meta: (size, speed, mem_type, module_class, ecc, manufacturer)
        self.variant_meta: Dict[
            str, Tuple[Optional[int], Optional[int], Optional[str], Optional[str], Optional[bool], Optional[str]]
        ] = {}
        self.variant_counts: Dict[str, int] = {}

        self.results_data: List[ScanResult] = []
        self.result_seq: int = 0

        self.latest_item_id: Optional[int] = None
        self.latest_key: Optional[str] = None

        self.counts_sort_column: str = "version"
        self.counts_sort_reverse: bool = False

        self._build_ui()
        self._load_results_from_file_and_rebuild()

    # --------- UI construction ----------
    def _build_ui(self):
        self.master.title(APP_TITLE)
        self.master.geometry("1300x800")
        self.master.minsize(1100, 680)
        self.master.resizable(True, True)

        paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(paned, padding=10)
        middle = ttk.Frame(paned, padding=10)
        right = ttk.Frame(paned, padding=10)
        paned.add(left, weight=1)
        paned.add(middle, weight=3)
        paned.add(right, weight=2)

        # ----- Left: Scanner -----
        scan_box = ttk.LabelFrame(left, text="Scanner Input")
        scan_box.pack(fill=tk.X)

        self.scan_var = tk.StringVar()
        self.scan_entry = ttk.Entry(scan_box, textvariable=self.scan_var, font=("Segoe UI", 14))
        self.scan_entry.pack(fill=tk.X, pady=(6, 6))
        self.scan_entry.focus_set()
        self.scan_entry.bind("<Return>", self.on_scan_submit)

        ttk.Label(scan_box, text="Scan a barcode and press Enter.").pack(anchor="w", pady=(0, 8))

        # ----- Left: Unknown Mapping -----
        map_box = ttk.LabelFrame(left, text="Unknown Mapping")
        map_box.pack(fill=tk.X, pady=(10, 6))

        ttk.Label(map_box, text="Barcode:").grid(row=0, column=0, sticky="w")
        self.map_code_var = tk.StringVar()
        self.map_code_entry = ttk.Entry(map_box, textvariable=self.map_code_var)
        self.map_code_entry.grid(row=0, column=1, sticky="ew")

        ttk.Label(map_box, text="Size (GB):").grid(row=1, column=0, sticky="w")
        self.map_size_var = tk.StringVar()
        self.map_size_entry = ttk.Entry(map_box, textvariable=self.map_size_var, width=8)
        self.map_size_entry.grid(row=1, column=1, sticky="w")

        ttk.Label(map_box, text="Module Class (PCx-XXXXX):").grid(row=2, column=0, sticky="w")
        self.map_class_var = tk.StringVar()
        self.map_class_entry = ttk.Entry(map_box, textvariable=self.map_class_var, width=16)
        self.map_class_entry.grid(row=2, column=1, sticky="w")

        ttk.Label(map_box, text="Manufacturer:").grid(row=3, column=0, sticky="w")
        self.map_mfr_var = tk.StringVar()
        self.map_mfr_entry = ttk.Entry(map_box, textvariable=self.map_mfr_var, width=16)
        self.map_mfr_entry.grid(row=3, column=1, sticky="w")
        self.map_mfr_entry.bind("<Return>", lambda e: self.on_save_mapping())

        ttk.Label(map_box, text="ECC:").grid(row=4, column=0, sticky="w")
        self.map_ecc_var = tk.StringVar()
        self.map_ecc_combo = ttk.Combobox(
            map_box,
            textvariable=self.map_ecc_var,
            values=["", "Yes", "No"],
            width=8,
            state="readonly",
        )
        self.map_ecc_combo.grid(row=4, column=1, sticky="w")
        self.map_ecc_combo.set("")

        self.map_regex_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(map_box, text="Treat pattern as REGEX", variable=self.map_regex_var).grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )

        btn_row = ttk.Frame(map_box)
        btn_row.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(btn_row, text="Save Mapping", command=self.on_save_mapping).pack(side=tk.LEFT)
        ttk.Button(btn_row, text="Clear", command=self.clear_mapping_form).pack(side=tk.LEFT, padx=(8, 0))

        # Right-side manual DDR/kind/speed panel
        manual_box = ttk.Frame(map_box)
        manual_box.grid(row=0, column=2, rowspan=7, padx=(10, 0), sticky="nsew")

        ttk.Label(manual_box, text="DDR Type:").grid(row=0, column=0, sticky="w")
        self.map_ddr_var = tk.StringVar()
        self.map_ddr_combo = ttk.Combobox(
            manual_box,
            textvariable=self.map_ddr_var,
            values=["", "DDR2", "DDR3", "DDR4", "DDR5"],
            width=8,
            state="readonly",
        )
        self.map_ddr_combo.grid(row=0, column=1, sticky="w")
        self.map_ddr_combo.set("")

        ttk.Label(manual_box, text="Reg/Unbuffered:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.map_kind_var = tk.StringVar()
        self.map_kind_combo = ttk.Combobox(
            manual_box,
            textvariable=self.map_kind_var,
            values=["", "Unbuffered", "Registered", "Load-Reduced"],
            width=14,
            state="readonly",
        )
        self.map_kind_combo.grid(row=1, column=1, sticky="w", pady=(4, 0))
        self.map_kind_combo.set("")

        ttk.Label(manual_box, text="Speed (MT/s):").grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.map_speed_var = tk.StringVar()
        self.map_speed_entry = ttk.Entry(manual_box, textvariable=self.map_speed_var, width=10)
        self.map_speed_entry.grid(row=2, column=1, sticky="w", pady=(4, 0))

        map_box.columnconfigure(1, weight=1)
        map_box.columnconfigure(2, weight=1)

        # Enter-to-advance across mapping fields
        self.map_code_entry.bind("<Return>", lambda e: self._enter_advance(self.map_code_var, self.map_size_entry))
        self.map_size_entry.bind("<Return>", lambda e: self._enter_advance(self.map_size_var, self.map_class_entry))
        self.map_class_entry.bind("<Return>", lambda e: self._enter_advance(self.map_class_var, self.map_mfr_entry))

        # ----- Left: Saved mappings -----
        saved_box = ttk.LabelFrame(left, text="Saved Mappings")
        saved_box.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        self.saved_list = tk.Listbox(saved_box, selectmode=tk.EXTENDED)
        self.saved_list.pack(fill=tk.BOTH, expand=True)

        del_row = ttk.Frame(saved_box)
        del_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(del_row, text="Delete Selected Mapping(s)", command=self.on_delete_mapping).pack(side=tk.LEFT)
        self.saved_list.bind("<Delete>", lambda e: self.on_delete_mapping())

        self.refresh_saved_mappings_list()

        # ----- Middle: Latest Scan -----
        latest_box = ttk.LabelFrame(middle, text="Latest Scan")
        latest_box.pack(fill=tk.X, pady=(0, 8))
        latest_grid = ttk.Frame(latest_box)
        latest_grid.pack(fill=tk.X, padx=8, pady=8)

        self.latest_ver_lbl = ttk.Label(latest_grid, text="v–", foreground=COLOR_VERSION, font=("Segoe UI", 22, "bold"))
        self.latest_ver_lbl.grid(row=0, column=0, sticky="w", padx=(0, 16))

        self.latest_size_val = ttk.Label(latest_grid, text="– GB", foreground=COLOR_SIZE, font=("Segoe UI", 14))
        self.latest_speed_val = ttk.Label(latest_grid, text="– MT/s", foreground=COLOR_SPEED, font=("Segoe UI", 14))
        self.latest_type_val = ttk.Label(latest_grid, text="–", foreground=COLOR_TYPE, font=("Segoe UI", 14))
        self.latest_class_val = ttk.Label(latest_grid, text="–", foreground=COLOR_CLASS, font=("Segoe UI", 14))
        self.latest_ecc_val = ttk.Label(latest_grid, text="–", foreground=COLOR_ECC, font=("Segoe UI", 14))
        self.latest_mfr_val = ttk.Label(latest_grid, text="–", foreground=COLOR_MFR, font=("Segoe UI", 14))
        self.latest_code_val = ttk.Label(latest_grid, text="–", foreground=COLOR_BARCODE, font=("Segoe UI", 12))

        ttk.Label(latest_grid, text="Size:").grid(row=0, column=1, sticky="e")
        self.latest_size_val.grid(row=0, column=2, sticky="w", padx=(6, 16))
        ttk.Label(latest_grid, text="Speed:").grid(row=0, column=3, sticky="e")
        self.latest_speed_val.grid(row=0, column=4, sticky="w", padx=(6, 16))

        ttk.Label(latest_grid, text="Type:").grid(row=1, column=1, sticky="e")
        self.latest_type_val.grid(row=1, column=2, sticky="w", padx=(6, 16))
        ttk.Label(latest_grid, text="Class:").grid(row=1, column=3, sticky="e")
        self.latest_class_val.grid(row=1, column=4, sticky="w", padx=(6, 16))

        ttk.Label(latest_grid, text="ECC:").grid(row=2, column=1, sticky="e")
        self.latest_ecc_val.grid(row=2, column=2, sticky="w", padx=(6, 16))
        ttk.Label(latest_grid, text="Mfr:").grid(row=2, column=3, sticky="e")
        self.latest_mfr_val.grid(row=2, column=4, sticky="w", padx=(6, 16))

        ttk.Label(latest_grid, text="Barcode:").grid(row=3, column=1, sticky="e", pady=(6, 0))
        self.latest_code_val.grid(row=3, column=2, columnspan=5, sticky="w", pady=(6, 0))

        self.latest_remove_btn = ttk.Button(latest_grid, text="Remove Latest", command=self.remove_latest_scan)
        self.latest_remove_btn.grid(row=0, column=6, sticky="e")

        # Right-click copy on latest panel
        self._attach_copy(self.latest_ver_lbl, lambda: self.latest_ver_lbl.cget("text"))
        for w in (
            self.latest_size_val,
            self.latest_speed_val,
            self.latest_type_val,
            self.latest_class_val,
            self.latest_ecc_val,
            self.latest_mfr_val,
            self.latest_code_val,
        ):
            self._attach_copy(w, lambda w=w: w.cget("text"))

        # ----- Middle: Results cards -----
        results_box = ttk.LabelFrame(middle, text="Scan Results")
        results_box.pack(fill=tk.BOTH, expand=True)

        self.results_canvas = tk.Canvas(results_box, highlightthickness=0)
        self.results_scroll = ttk.Scrollbar(results_box, orient=tk.VERTICAL, command=self.results_canvas.yview)
        self.results_canvas.configure(yscrollcommand=self.results_scroll.set)
        self.results_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.results_inner = ttk.Frame(self.results_canvas)
        self.results_canvas.create_window((0, 0), window=self.results_inner, anchor="nw")
        self.results_inner.bind(
            "<Configure>", lambda e: self.results_canvas.configure(scrollregion=self.results_canvas.bbox("all"))
        )

        self.results_items: List[Dict] = []

        action_row = ttk.Frame(middle)
        action_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(action_row, text="Clear Results & Counts", command=self.clear_results_and_counts).pack(side=tk.LEFT)

        # ----- Right: Variant counts -----
        counts_box = ttk.LabelFrame(right, text="Variant Counts")
        counts_box.pack(fill=tk.BOTH, expand=True)

        self.counts_tree = ttk.Treeview(
            counts_box,
            columns=("variant", "size", "speed", "class", "ecc", "manufacturer", "version", "count"),
            show="headings",
            selectmode="browse",
        )
        self.counts_tree.heading("variant", text="Variant (Barcode)", command=lambda c="variant": self.on_counts_heading_click(c))
        self.counts_tree.heading("size", text="Size (GB)", command=lambda c="size": self.on_counts_heading_click(c))
        self.counts_tree.heading("speed", text="Speed (MT/s)", command=lambda c="speed": self.on_counts_heading_click(c))
        self.counts_tree.heading("class", text="Class", command=lambda c="class": self.on_counts_heading_click(c))
        self.counts_tree.heading("ecc", text="ECC", command=lambda c="ecc": self.on_counts_heading_click(c))
        self.counts_tree.heading("manufacturer", text="Manufacturer", command=lambda c="manufacturer": self.on_counts_heading_click(c))
        self.counts_tree.heading("version", text="Version", command=lambda c="version": self.on_counts_heading_click(c))
        self.counts_tree.heading("count", text="Count", command=lambda c="count": self.on_counts_heading_click(c))

        self.counts_tree.column("variant", width=260, anchor="w")
        self.counts_tree.column("size", width=70, anchor="center")
        self.counts_tree.column("speed", width=90, anchor="center")
        self.counts_tree.column("class", width=160, anchor="center")
        self.counts_tree.column("ecc", width=60, anchor="center")
        self.counts_tree.column("manufacturer", width=120, anchor="center")
        self.counts_tree.column("version", width=70, anchor="center")
        self.counts_tree.column("count", width=70, anchor="center")
        self.counts_tree.pack(fill=tk.BOTH, expand=True)

        counts_btn_row = ttk.Frame(counts_box)
        counts_btn_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(counts_btn_row, text="Export Versions…", command=self.export_versions).pack(side=tk.LEFT)

        # Right-click copy in counts table
        self.counts_tree.bind("<Button-3>", self.on_counts_right_click)
        self.counts_tree.bind("<Button-2>", self.on_counts_right_click)

    # --------- Results persistence ----------
    def _save_results_file(self):
        """
        Persist scan history with minimal info:
        each entry only stores id, timestamp, barcode, version.
        Specs are reconstructed from mappings/heuristics on load.
        """
        data = [
            {
                "id": r.id,
                "timestamp": r.timestamp,
                "barcode": r.barcode,
                "version": r.version,
            }
            for r in self.results_data
        ]
        RESULTS_JSON.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def _load_results_file(self) -> List[ScanResult]:
        """
        Load scan history from scan_results.json which now only has
        id, timestamp, barcode, version. For each entry we re-parse
        the barcode using current mappings to get size/speed/type/etc.
        """
        if not RESULTS_JSON.exists():
            return []
        try:
            raw = json.loads(RESULTS_JSON.read_text(encoding="utf-8"))
        except Exception:
            return []

        out: List[ScanResult] = []
        if not isinstance(raw, list):
            return out

        for obj in raw:
            try:
                barcode = str(obj.get("barcode", "")).strip()
                if not barcode:
                    continue

                rid = int(obj.get("id", len(out) + 1))
                ts = str(obj.get("timestamp") or datetime.now(timezone.utc).isoformat(timespec="seconds"))
                version = int(obj.get("version", 0))

                size_gb, speed_mts, mem_type, manufacturer, module_class, ecc, parsed_ok = parse_barcode(
                    barcode, self.store
                )

                # Use safe defaults if parsing fails (so GUI still works)
                size_safe = size_gb or 0
                speed_safe = speed_mts or 0
                type_safe = mem_type or ""
                mfr_safe = manufacturer or ""

                out.append(
                    ScanResult(
                        id=rid,
                        timestamp=ts,
                        barcode=barcode,
                        size_gb=size_safe,
                        speed_mts=speed_safe,
                        mem_type=type_safe,
                        module_class=module_class,
                        ecc=ecc,
                        manufacturer=mfr_safe,
                        version=version,
                    )
                )
            except Exception:
                continue

        return out

    def _load_results_from_file_and_rebuild(self):
        self.results_data = self._load_results_file()
        self.result_seq = max((r.id for r in self.results_data), default=0)
        self._rebuild_everything_from_results()

    def _rebuild_everything_from_results(self):
        """
        Rebuild counts + meta from scan results, but KEEP permanent
        barcode_versions mapping. New barcodes get new versions.
        """
        self.variant_counts.clear()
        self.variant_meta.clear()

        next_ver = (max(self.barcode_versions.values()) if self.barcode_versions else 0) + 1
        for r in self.results_data:
            b = r.barcode
            if b not in self.barcode_versions:
                self.barcode_versions[b] = next_ver
                next_ver += 1
            r.version = self.barcode_versions[b]

            self.variant_counts[b] = self.variant_counts.get(b, 0) + 1
            self.variant_meta[b] = (
                r.size_gb,
                r.speed_mts,
                r.mem_type,
                r.module_class,
                r.ecc,
                r.manufacturer,
            )

        # persist versions alongside mappings
        self.store.versions = dict(self.barcode_versions)
        self.store.save()

        # Rebuild results cards
        for it in getattr(self, "results_items", []):
            try:
                it["frame"].destroy()
            except Exception:
                pass
        self.results_items = []

        if self.results_data:
            visible_results = (
                self.results_data[-MAX_VISIBLE_RESULTS:]
                if MAX_VISIBLE_RESULTS and MAX_VISIBLE_RESULTS > 0
                else self.results_data
            )
            for r in visible_results:
                self._render_card_from_result(r, add_to_top=True)

            newest = self.results_data[-1]
            self.latest_item_id = newest.id
            self.latest_key = newest.barcode
            self.update_latest_panel(
                newest.version,
                newest.size_gb,
                newest.speed_mts,
                newest.mem_type,
                newest.module_class,
                newest.ecc,
                newest.manufacturer,
                newest.barcode,
            )
        else:
            self.latest_item_id = None
            self.latest_key = None
            self.update_latest_panel(0, 0, 0, "", None, None, "", "")

        self.refresh_counts_view()

    # --------- Small helpers ----------
    def _enter_advance(self, var: tk.StringVar, next_widget):
        if var.get().strip():
            try:
                next_widget.focus_set()
                try:
                    next_widget.icursor(tk.END)
                except Exception:
                    pass
            except Exception:
                pass

    def _toast(self, x: int, y: int, text: str = "Copied", ms: int = 700):
        try:
            top = tk.Toplevel(self.master)
            top.overrideredirect(True)
            top.attributes("-topmost", True)
            top.geometry(f"+{x+10}+{y+10}")
            lbl = ttk.Label(top, text=text, relief=tk.SOLID, padding=(6, 3))
            lbl.pack()
            top.after(ms, top.destroy)
        except Exception:
            pass

    def _copy_text(self, text: str, event: Optional[tk.Event] = None):
        try:
            self.master.clipboard_clear()
            self.master.clipboard_append(text)
            if event is not None:
                self._toast(event.x_root, event.y_root, "Copied")
        except Exception:
            pass

    def _attach_copy(self, widget: tk.Widget, text_getter):
        def handler(event, tg=text_getter):
            txt = tg()
            if txt is None:
                return
            self._copy_text(str(txt), event)

        widget.bind("<Button-3>", handler)
        widget.bind("<Button-2>", handler)

    # --------- Counts UI ----------
    def on_counts_right_click(self, event):
        row_id = self.counts_tree.identify_row(event.y)
        col_id = self.counts_tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        self.counts_tree.selection_set(row_id)
        values = self.counts_tree.item(row_id, "values")
        try:
            idx = int(col_id[1:]) - 1
            if 0 <= idx < len(values):
                self._copy_text(str(values[idx]), event)
        except Exception:
            pass

    def on_counts_heading_click(self, column: str):
        if self.counts_sort_column == column:
            self.counts_sort_reverse = not self.counts_sort_reverse
        else:
            self.counts_sort_column = column
            self.counts_sort_reverse = False
        self.refresh_counts_view()

    def _counts_sort_key(self, row: Dict) -> object:
        col = self.counts_sort_column
        val = row.get(col)
        if col in ("size", "speed", "version", "count"):
            try:
                return int(val) if val is not None else -1
            except Exception:
                return -1
        return str(val).lower() if val is not None else ""

    # --------- Export ----------
    def export_versions(self):
        if not self.barcode_versions:
            messagebox.showinfo(APP_TITLE, "No variants to export yet.")
            return

        filename = filedialog.asksaveasfilename(
            title="Export Versions",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")],
        )
        if not filename:
            return

        rows = []
        for barcode, version in self.barcode_versions.items():
            meta = self.variant_meta.get(barcode)
            if meta is None:
                found = next((r for r in self.results_data if r.barcode == barcode), None)
                if found:
                    meta = (
                        found.size_gb,
                        found.speed_mts,
                        found.mem_type,
                        found.module_class,
                        found.ecc,
                        found.manufacturer,
                    )
                else:
                    meta = (None, None, None, None, None, None)

            size_gb, speed, mem_type, module_class, ecc, mfr = meta
            count = self.variant_counts.get(barcode, 0)

            spec_parts = []
            if size_gb is not None:
                spec_parts.append(f"{size_gb} GB")
            if speed is not None:
                spec_parts.append(f"{speed} MT/s")
            if mem_type:
                spec_parts.append(str(mem_type))
            spec = " ".join(spec_parts)

            ecc_str = "Yes" if ecc else ("No" if ecc is not None else "?")

            rows.append(
                {
                    "VersionNum": version,
                    "Version": f"v{version}",
                    "Spec": spec,
                    "Class": module_class or "",
                    "ECC": ecc_str,
                    "Manufacturer": mfr or "",
                    "Count": count,
                    "Barcode": barcode,
                }
            )

        rows.sort(key=lambda r: r["VersionNum"])

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Variants"

            headers = ["Version", "Spec", "Class", "ECC", "Manufacturer", "Count", "Barcode"]
            ws.append(headers)

            for r in rows:
                ws.append(
                    [
                        r["Version"],
                        r["Spec"],
                        r["Class"],
                        r["ECC"],
                        r["Manufacturer"],
                        r["Count"],
                        r["Barcode"],
                    ]
                )

            wb.save(filename)
            messagebox.showinfo(APP_TITLE, f"Versions exported to:\n{filename}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Failed to export versions:\n{e}")

    # --------- Scan handling ----------
    def _clean_scanned_code(self, raw: str) -> str:
        """
        Canonical barcode key used for mappings + version binding.

        Rules:
        - Strip outer whitespace.
        - If first token is a 2-letter code (CN, KR, etc.), drop it.
        - If there are at least two tokens and the last two are all digits
          (e.g. '10680063 123201927'), keep the LAST token.
        - Otherwise keep the first remaining token.
        - Then strip anything after '+' in that token.
        """
        s = raw.strip()
        if not s:
            return s

        tokens = s.split()

        # Drop country code prefix, e.g. "CN", "KR"
        if len(tokens) >= 2 and re.fullmatch(r"[A-Z]{2}", tokens[0]):
            tokens = tokens[1:]

        if len(tokens) >= 2 and tokens[-1].isdigit() and tokens[-2].isdigit():
            candidate = tokens[-1]
        else:
            candidate = tokens[0]

        if "+" in candidate:
            candidate = candidate.split("+", 1)[0]

        return candidate

    def on_scan_submit(self, event=None):
        raw_code = self.scan_var.get().strip()
        if not raw_code:
            self.scan_entry.focus_set()
            return

        barcode_key = self._clean_scanned_code(raw_code)

        size_gb, speed_mts, mem_type, manufacturer, module_class, ecc, parsed_ok = parse_barcode(
            barcode_key, self.store
        )

        if parsed_ok:
            self.add_result(barcode_key, size_gb, speed_mts, mem_type, module_class, ecc, manufacturer)
            play_success()
            self.scan_var.set("")
            self.scan_entry.focus_set()
        else:
            # Fill unknown mapping form with the barcode key
            self.map_code_var.set(barcode_key)
            self.map_size_var.set(str(size_gb or ""))
            self.map_class_var.set(str(module_class or ""))
            self.map_mfr_var.set(str(manufacturer or ""))
            if ecc is True:
                self.map_ecc_var.set("Yes")
            elif ecc is False:
                self.map_ecc_var.set("No")
            else:
                self.map_ecc_var.set("")

            self.map_ddr_var.set("")
            self.map_kind_var.set("")
            self.map_speed_var.set("")

            play_unknown()

            target = None
            if not size_gb:
                target = self.map_size_var
            elif not module_class:
                target = self.map_class_var
            elif not manufacturer:
                target = self.map_mfr_var
            self.focus_mapping_field(target)

    def focus_mapping_field(self, var: Optional[tk.StringVar]):
        self.master.lift()
        for w in self.master.winfo_children():
            if self._focus_entry_recursive(w, var):
                return

    def _focus_entry_recursive(self, widget, target_var):
        try:
            if isinstance(widget, (ttk.Entry, ttk.Combobox)):
                tv = widget.cget("textvariable")
                if tv and target_var and str(target_var) == str(tv):
                    widget.focus_set()
                    try:
                        widget.icursor(tk.END)
                    except Exception:
                        pass
                    return True
        except Exception:
            pass
        for child in widget.winfo_children():
            if self._focus_entry_recursive(child, target_var):
                return True
        return False

    # --------- Save mapping ----------
    def on_save_mapping(self):
        raw_code = self.map_code_var.get().strip()
        if not raw_code:
            return

        barcode_key = self._clean_scanned_code(raw_code)

        def as_int(s):
            try:
                return int(s)
            except Exception:
                return None

        size_gb = as_int(self.map_size_var.get())
        module_class = self.map_class_var.get().strip().upper() or None
        manufacturer = self.map_mfr_var.get().strip().title() or None

        ecc_sel = self.map_ecc_var.get()
        if ecc_sel == "Yes":
            ecc: Optional[bool] = True
        elif ecc_sel == "No":
            ecc = False
        else:
            ecc = None

        manual_ddr = self.map_ddr_var.get().strip().upper()
        manual_speed = as_int(self.map_speed_var.get())
        manual_kind = self.map_kind_var.get().strip()

        mem_type: Optional[str] = None
        speed_mts: Optional[int] = None

        if manual_ddr:
            mem_type = manual_ddr
        if manual_speed:
            speed_mts = manual_speed

        if (mem_type is None or speed_mts is None) and module_class:
            mc_type, mc_speed, mc_ecc = parse_module_class(module_class)
            if mem_type is None and mc_type:
                mem_type = mc_type
            if speed_mts is None and mc_speed:
                speed_mts = mc_speed
            if ecc is None and mc_ecc is not None:
                ecc = mc_ecc

        # If no explicit Module Class but we have manual DDR + speed,
        # synthesize a PCx-XXXXX class so it behaves like a normal scan.
        if not module_class and (mem_type or speed_mts):
            auto_class = synth_module_class(mem_type, speed_mts, manual_kind)
            if auto_class:
                module_class = auto_class
            else:
                # Fallback: human-readable "DDR3 1333" etc.
                parts = []
                if mem_type:
                    parts.append(mem_type)
                if speed_mts:
                    parts.append(str(speed_mts))
                if manual_kind:
                    parts.append(manual_kind)
                if parts:
                    module_class = " ".join(parts)

        m = Mapping(
            pattern=barcode_key,
            size_gb=size_gb,
            speed_mts=speed_mts,
            mem_type=mem_type,
            manufacturer=manufacturer,
            module_class=module_class,
            ecc=ecc,
            regex=self.map_regex_var.get(),
        )
        self.store.add(m)
        self.refresh_saved_mappings_list()

        # Ensure this exact barcode has a permanent version
        if barcode_key not in self.barcode_versions:
            next_ver = (max(self.barcode_versions.values()) if self.barcode_versions else 0) + 1
            self.barcode_versions[barcode_key] = next_ver
            self.store.versions = dict(self.barcode_versions)
            self.store.save()

        # Immediately try scanning with the new mapping
        size_gb2, speed_mts2, mem_type2, manufacturer2, module_class2, ecc2, parsed_ok = parse_barcode(
            barcode_key, self.store
        )
        if parsed_ok:
            self.add_result(barcode_key, size_gb2, speed_mts2, mem_type2, module_class2, ecc2, manufacturer2)
            play_success()

        self.clear_mapping_form()
        self.scan_var.set("")
        self.scan_entry.focus_set()

    # --------- Results handling ----------
    def _next_result_id(self) -> int:
        self.result_seq += 1
        return self.result_seq

    def add_result(
        self,
        barcode_key: str,
        size_gb: int,
        speed_mts: int,
        mem_type: str,
        module_class: Optional[str],
        ecc: Optional[bool],
        manufacturer: str,
    ):
        rid = self._next_result_id()
        ts = datetime.now(timezone.utc).isoformat(timespec="seconds")

        if barcode_key in self.barcode_versions:
            version = self.barcode_versions[barcode_key]
        else:
            version = (max(self.barcode_versions.values()) if self.barcode_versions else 0) + 1
            self.barcode_versions[barcode_key] = version
            self.store.versions = dict(self.barcode_versions)
            self.store.save()

        self.variant_counts[barcode_key] = self.variant_counts.get(barcode_key, 0) + 1
        self.variant_meta[barcode_key] = (size_gb, speed_mts, mem_type, module_class, ecc, manufacturer)

        sr = ScanResult(
            id=rid,
            timestamp=ts,
            barcode=barcode_key,
            size_gb=size_gb,
            speed_mts=speed_mts,
            mem_type=mem_type,
            module_class=module_class,
            ecc=ecc,
            manufacturer=manufacturer,
            version=version,
        )
        self.results_data.append(sr)

        self._render_card_from_result(sr, add_to_top=True)

        if MAX_VISIBLE_RESULTS and len(self.results_items) > MAX_VISIBLE_RESULTS:
            oldest = self.results_items.pop()
            try:
                oldest["frame"].destroy()
            except Exception:
                pass

        self.latest_item_id = sr.id
        self.latest_key = sr.barcode
        self.update_latest_panel(
            sr.version,
            sr.size_gb,
            sr.speed_mts,
            sr.mem_type,
            sr.module_class,
            sr.ecc,
            sr.manufacturer,
            sr.barcode,
        )

        self.refresh_counts_view()
        self._save_results_file()

    def _render_card_from_result(self, r: ScanResult, add_to_top: bool):
        card = ttk.Frame(self.results_inner, padding=8, relief=tk.RIDGE)
        card.grid_columnconfigure(1, weight=1)

        size_lbl = ttk.Label(card, text=f"{r.size_gb} GB", foreground=COLOR_SIZE)
        speed_lbl = ttk.Label(card, text=f"{r.speed_mts} MT/s", foreground=COLOR_SPEED)
        type_lbl = ttk.Label(card, text=f"{r.mem_type.upper()}", foreground=COLOR_TYPE)
        class_lbl = ttk.Label(card, text=f"{r.module_class or ''}", foreground=COLOR_CLASS)
        ecc_str = "Yes" if r.ecc else ("No" if r.ecc is not None else "?")
        ecc_lbl = ttk.Label(card, text=ecc_str, foreground=COLOR_ECC)
        mfr_lbl = ttk.Label(card, text=f"{r.manufacturer}", foreground=COLOR_MFR)
        ver_lbl = ttk.Label(card, text=f"v{r.version}", foreground=COLOR_VERSION)
        code_lbl = ttk.Label(card, text=r.barcode, foreground=COLOR_BARCODE)

        ttk.Label(card, text="Size:").grid(row=0, column=0, sticky="w")
        size_lbl.grid(row=0, column=1, sticky="w")
        ttk.Label(card, text="Speed:").grid(row=0, column=2, sticky="w", padx=(12, 0))
        speed_lbl.grid(row=0, column=3, sticky="w")
        ttk.Label(card, text="Version:").grid(row=0, column=4, sticky="w", padx=(12, 0))
        ver_lbl.grid(row=0, column=5, sticky="w")

        ttk.Label(card, text="Type:").grid(row=1, column=0, sticky="w")
        type_lbl.grid(row=1, column=1, sticky="w")
        ttk.Label(card, text="Class:").grid(row=1, column=2, sticky="w", padx=(12, 0))
        class_lbl.grid(row=1, column=3, sticky="w")
        ttk.Label(card, text="ECC:").grid(row=1, column=4, sticky="w", padx=(12, 0))
        ecc_lbl.grid(row=1, column=5, sticky="w")
        ttk.Label(card, text="Mfr:").grid(row=1, column=6, sticky="w", padx=(12, 0))
        mfr_lbl.grid(row=1, column=7, sticky="w")

        ttk.Label(card, text="Barcode:").grid(row=2, column=0, sticky="w")
        code_lbl.grid(row=2, column=1, columnspan=7, sticky="w")

        for w in (size_lbl, speed_lbl, type_lbl, class_lbl, ecc_lbl, mfr_lbl, ver_lbl, code_lbl):
            self._attach_copy(w, lambda w=w: w.cget("text"))

        rm_btn = ttk.Button(
            card,
            text="Remove",
            command=lambda i=r.id, k=r.barcode, f=card: self.remove_result(i, k, f),
        )
        rm_btn.grid(row=0, column=6, rowspan=3, sticky="e")

        if add_to_top and self.results_items:
            card.pack(fill=tk.X, expand=True, pady=4, before=self.results_items[0]["frame"])
            self.results_items.insert(0, {"id": r.id, "frame": card, "key": r.barcode})
        else:
            card.pack(fill=tk.X, expand=True, pady=4)
            self.results_items.append({"id": r.id, "frame": card, "key": r.barcode})

        try:
            self.results_canvas.update_idletasks()
            self.results_canvas.yview_moveto(0)
        except Exception:
            pass

    def update_latest_panel(
        self,
        version: int,
        size_gb: int,
        speed_mts: int,
        mem_type: str,
        module_class: Optional[str],
        ecc: Optional[bool],
        manufacturer: str,
        barcode: str,
    ):
        self.latest_ver_lbl.configure(text=f"v{version}" if version else "v–")
        self.latest_size_val.configure(text=f"{size_gb} GB" if size_gb else "– GB")
        self.latest_speed_val.configure(text=f"{speed_mts} MT/s" if speed_mts else "– MT/s")
        self.latest_type_val.configure(text=f"{mem_type.upper()}" if mem_type else "–")
        self.latest_class_val.configure(text=module_class or "–")
        ecc_str = "Yes" if ecc else ("No" if ecc is not None else "–")
        self.latest_ecc_val.configure(text=ecc_str)
        self.latest_mfr_val.configure(text=manufacturer if manufacturer else "–")
        self.latest_code_val.configure(text=barcode if barcode else "–")

    def remove_latest_scan(self):
        if self.latest_item_id is None:
            return
        self.results_data = [r for r in self.results_data if r.id != self.latest_item_id]
        self._rebuild_everything_from_results()
        self._save_results_file()

    def remove_result(self, item_id: int, key: str, frame: ttk.Frame):
        self.results_data = [r for r in self.results_data if r.id != item_id]
        self._rebuild_everything_from_results()
        self._save_results_file()

    def refresh_counts_view(self):
        self.counts_tree.delete(*self.counts_tree.get_children())

        rows: List[Dict] = []
        for key, count in self.variant_counts.items():
            ver = self.barcode_versions.get(key, 0)
            meta = self.variant_meta.get(key, (None, None, None, None, None, None))
            size_val = meta[0]
            speed_val = meta[1]
            module_class = meta[3]
            ecc = meta[4]
            mfr = meta[5] if meta[5] is not None else "?"
            ecc_str = "Yes" if ecc else ("No" if ecc is not None else "?")
            rows.append(
                {
                    "variant": key,
                    "size": size_val,
                    "speed": speed_val,
                    "class": module_class or "",
                    "ecc": ecc_str,
                    "manufacturer": mfr,
                    "version": ver,
                    "count": count,
                }
            )

        rows.sort(key=self._counts_sort_key, reverse=self.counts_sort_reverse)

        for row in rows:
            size_disp = row["size"] if row["size"] is not None else "?"
            speed_disp = row["speed"] if row["speed"] is not None else "?"
            ver_disp = f"v{row['version']}" if row["version"] else "?"
            count_disp = row["count"]
            self.counts_tree.insert(
                "",
                tk.END,
                values=(
                    row["variant"],
                    size_disp,
                    speed_disp,
                    row["class"],
                    row["ecc"],
                    row["manufacturer"],
                    ver_disp,
                    count_disp,
                ),
            )

    def clear_results_and_counts(self):
        if messagebox.askyesno(
            APP_TITLE,
            "Clear all scan results and reset counts? (Saved mappings and versions are kept)",
        ):
            # Clear data
            self.results_data.clear()
            self.variant_counts.clear()
            self.variant_meta.clear()
            self._save_results_file()

            # Clear UI cards
            for it in self.results_items:
                try:
                    it["frame"].destroy()
                except Exception:
                    pass
            self.results_items = []

            # Clear counts + latest panel
            self.refresh_counts_view()
            self.latest_item_id = None
            self.latest_key = None
            self.update_latest_panel(0, 0, 0, "", None, None, "", "")

    # --------- Mappings UI ----------
    def on_delete_mapping(self):
        selection = list(self.saved_list.curselection())
        if not selection:
            return
        for idx in sorted(selection, reverse=True):
            self.store.remove_index(idx)
        self.refresh_saved_mappings_list()

    def clear_mapping_form(self):
        self.map_code_var.set("")
        self.map_size_var.set("")
        self.map_class_var.set("")
        self.map_mfr_var.set("")
        self.map_ecc_var.set("")
        self.map_regex_var.set(False)
        self.map_ddr_var.set("")
        self.map_kind_var.set("")
        self.map_speed_var.set("")

    def refresh_saved_mappings_list(self):
        self.saved_list.delete(0, tk.END)
        for desc in self.store.all_descriptions():
            self.saved_list.insert(tk.END, desc)


# -------------------- Main --------------------
def main():
    global root
    root = tk.Tk()
    root.resizable(True, True)
    store = MappingStore(MAPPINGS_JSON)
    App(root, store)
    root.mainloop()


if __name__ == "__main__":
    main()
