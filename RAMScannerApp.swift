import SwiftUI
import AppKit

// ===== App Entrypoint =====
@main
struct RamScannerApp: App {
    var body: some Scene {
        WindowGroup {
            ContentView()
                .frame(width: 1600, height: 900)
        }
        .defaultSize(width: 1600, height: 900)
        .windowResizability(.contentSize)
        .windowStyle(HiddenTitleBarWindowStyle())
    }
}

// ===== Persistence Keys =====
fileprivate enum UDKey {
    static let results                = "results_v9"
    static let customSpeedMap         = "customSpeedMap_v9"
    static let customSizeMap          = "customSizeMap_v9"
    static let customMfgMap           = "customMfgMap_v9"
    static let variantIndex           = "variantIndex_v9"
    static let speedIndex             = "speedIndex_v9"
    static let sizeIndex              = "sizeIndex_v9"
    static let nextVariantId          = "nextVariantId_v9"
    static let nextSpeedId            = "nextSpeedId_v9"
    static let nextSizeId             = "nextSizeId_v9"
}

// ===== Data Model =====
struct ScanResult: Identifiable, Hashable, Codable {
    let id: UUID
    let barcode: String
    let normalizedVariant: String?
    var speed: String
    var size: String
    let manufacturer: String
    let date: Date
    var variantId: Int
    var speedId: Int
    var sizeId: Int
}

// ===== Sound (debounced) =====
final class Sounder {
    static let shared = Sounder()
    private var lastPlay: Date = .distantPast
    private let minGap: TimeInterval = 0.20
    private var current: NSSound?
    private func canPlay() -> Bool { Date().timeIntervalSince(lastPlay) >= minGap }
    private func playNamed(_ name: String) {
        guard canPlay() else { return }
        lastPlay = Date()
        current?.stop()
        if let s = NSSound(named: NSSound.Name(name)) { current = s; s.play() } else { NSSound.beep() }
    }
    func success() { playNamed("Pop") }
    func error()   { playNamed("Basso") }
}

// ===== Main View =====
struct ContentView: View {
    enum FocusField: Hashable { case scan, addCombinedCode, addCombinedSpeed, addCombinedSize, addCombinedMfg, batch }
    @FocusState private var focusedField: FocusField?

    enum SortKey: String, CaseIterable, Identifiable {
        case dateDesc = "Date", variant = "Variant#", speed = "Speed#", size = "Size#"
        var id: String { rawValue }
    }

    @State private var sortKey: SortKey = .dateDesc
    @State private var scanText: String = ""
    @State private var results: [ScanResult] = []
    @State private var lastScanned: ScanResult? = nil
    @State private var batchInput: String = ""

    // Add Known Mapping fields
    @State private var mapCode: String = ""
    @State private var mapSpeed: String = ""
    @State private var mapSize: String = ""
    @State private var mapManufacturer: String = ""   // NEW

    // Delete dialogs
    @State private var variantPendingDelete: String? = nil
    @State private var showDeleteVariantAlert: Bool = false

    @State private var mappingPendingDeleteKey: String? = nil
    @State private var showDeleteMappingAlert: Bool = false

    // ID tracking
    @State private var variantIndex: [String: Int] = [:]
    @State private var speedIndex: [String: Int] = [:]
    @State private var sizeIndex: [String: Int] = [:]
    @State private var nextVariantId: Int = 1
    @State private var nextSpeedId: Int = 1
    @State private var nextSizeId: Int = 1

    // Unknown queue
    @State private var pendingUnknowns: [String] = []

    // Base + custom maps
    private let baseSpeedMap: [String: String] = [
        "M393A4K40BB2": "2666 MHz", "M393A4K40DB3": "3200 MHz",
        "+VK": "2666 MHz", "+XN": "3200 MHz", "BB2": "2666 MHz", "DB3": "3200 MHz",
        "3200": "3200 MHz", "2666": "2666 MHz", "2933": "2933 MHz", "2400": "2400 MHz", "2133": "2133 MHz"
    ]
    @State private var customSpeedMap: [String: String] = [:]
    @State private var customSizeMap:  [String: String] = [:]
    @State private var customManufacturerMap: [String: String] = [:] // NEW

    private var fullSpeedMap: [String: String] { baseSpeedMap.merging(customSpeedMap) { _, n in n } }

    var body: some View {
        HStack(spacing: 16) {
            leftPanel
            Divider()
            centerPanel
            Divider()
            rightPanel
        }
        .onAppear { loadAll(); focusedField = .scan }
        .onChange(of: results) { _, _ in saveResults() }
        .onChange(of: customSpeedMap) { _, _ in saveCustomSpeedMap() }
        .onChange(of: customSizeMap) { _, _ in saveCustomSizeMap() }
        .onChange(of: customManufacturerMap) { _, _ in saveCustomMfgMap() }
        .alert("Delete variant?", isPresented: $showDeleteVariantAlert, actions: {
            Button("Cancel", role: .cancel) { variantPendingDelete = nil }
            Button("Delete", role: .destructive) { performDeleteVariant() }
        }, message: { Text("This will remove all scans for variant “\(variantPendingDelete ?? "")”.") })
        .alert("Remove mapping?", isPresented: $showDeleteMappingAlert, actions: {
            Button("Cancel", role: .cancel) { mappingPendingDeleteKey = nil }
            Button("Remove", role: .destructive) { performDeleteMapping() }
        }, message: { Text("This removes the custom mapping for “\(mappingPendingDeleteKey ?? "")”.") })
    }

    // ===== Panels =====
    private var leftPanel: some View {
        VStack(alignment: .leading, spacing: 16) {
            Text("RAM Barcode Scanner").font(.system(size: 30, weight: .bold)).padding(.top, 12)
            Text("Scan a barcode with your handheld scanner.").font(.system(size: 15)).foregroundColor(.secondary)

            VStack(alignment: .leading, spacing: 8) {
                Text("Scan input").font(.headline)
                TextField("Waiting for scan...", text: $scanText)
                    .textFieldStyle(RoundedBorderTextFieldStyle())
                    .font(.system(size: 18, design: .monospaced))
                    .focused($focusedField, equals: .scan)
                    .onSubmit { processScan() }
            }

            if let last = lastScanned {
                LastScanBox(last: last, speedColor: speedColor(_:), sizeColors: sizeTagColors(_:))
            }

            HStack(spacing: 12) {
                Button(action: copyResultsAsCSV) { Label("Copy CSV", systemImage: "doc.on.doc") }
                Button(action: {
                    results.removeAll()
                    lastScanned = nil
                    resetIds()
                    saveAll()
                }) { Label("Clear All", systemImage: "trash") }
                Spacer()
                Text("Results: \(results.count)").foregroundColor(.secondary)
            }

            Divider().padding(.vertical, 10)

            // Add Known Mapping
            VStack(alignment: .leading, spacing: 10) {
                Text("Add Known Mapping").font(.headline)
                HStack(spacing: 8) {
                    TextField("Barcode substring or variant (e.g. HMA84…+VK)", text: $mapCode)
                        .textFieldStyle(RoundedBorderTextFieldStyle())
                        .focused($focusedField, equals: .addCombinedCode)
                        .onSubmit { focusedField = .addCombinedSpeed }

                    TextField("Speed (e.g. 3200 MHz)", text: $mapSpeed)
                        .textFieldStyle(RoundedBorderTextFieldStyle())
                        .focused($focusedField, equals: .addCombinedSpeed)
                        .onSubmit { focusedField = .addCombinedSize }

                    TextField("Size (e.g. 16 GB or 16384 MB)", text: $mapSize)
                        .textFieldStyle(RoundedBorderTextFieldStyle())
                        .focused($focusedField, equals: .addCombinedSize)
                        .onSubmit { focusedField = .addCombinedMfg }
                }
                HStack(spacing: 8) {
                    TextField("Manufacturer (optional)", text: $mapManufacturer) // NEW
                        .textFieldStyle(RoundedBorderTextFieldStyle())
                        .focused($focusedField, equals: .addCombinedMfg)
                        .onSubmit { addCombinedMapping() }

                    Button("Add", action: addCombinedMapping)
                        .disabled(mapCode.trimmingCharacters(in: .whitespaces).isEmpty ||
                                  mapSpeed.trimmingCharacters(in: .whitespaces).isEmpty ||
                                  mapSize.trimmingCharacters(in: .whitespaces).isEmpty)
                }
            }

            // Known Mappings
            VStack(alignment: .leading, spacing: 8) {
                Divider().padding(.vertical, 8)
                Text("Known Mappings").font(.headline)
                if mappingKeys.isEmpty {
                    Text("No custom mappings yet.").foregroundColor(.secondary)
                } else {
                    ScrollView {
                        VStack(spacing: 6) {
                            ForEach(mappingKeys, id: \.self) { key in
                                HStack(spacing: 8) {
                                    Text(key).font(.system(size: 12, design: .monospaced))
                                    Spacer()
                                    if let mfg = customManufacturerMap[key], !mfg.isEmpty {
                                        Tag(text: mfg)
                                    }
                                    if let spd = customSpeedMap[key] {
                                        Tag(text: spd)
                                    }
                                    if let sz = customSizeMap[key] {
                                        let sc = sizeTagColors(sz); Tag(text: sz, fg: sc.fg, bg: sc.bg)
                                    }
                                    Button {
                                        mappingPendingDeleteKey = key
                                        showDeleteMappingAlert = true
                                    } label: { Image(systemName: "trash") }
                                    .buttonStyle(.borderless)
                                }
                                .padding(8)
                                .background(RoundedRectangle(cornerRadius: 8).fill(Color(NSColor.windowBackgroundColor)))
                                .overlay(RoundedRectangle(cornerRadius: 8).stroke(Color.secondary.opacity(0.15), lineWidth: 1))
                            }
                        }
                    }
                    .frame(minHeight: 120, maxHeight: 220)
                }
            }

            Divider().padding(.vertical, 10)

            // Batch
            VStack(alignment: .leading, spacing: 8) {
                Text("Batch paste (one barcode per line)").font(.headline)
                TextEditor(text: $batchInput)
                    .font(.system(size: 14, design: .monospaced))
                    .frame(minHeight: 100)
                    .overlay(RoundedRectangle(cornerRadius: 8).stroke(Color.secondary.opacity(0.3), lineWidth: 1))
                    .focused($focusedField, equals: .batch)
                HStack {
                    Button("Process Lines") { processBatch() }
                        .disabled(batchInput.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty)
                    Button("Clear") { batchInput = "" }.foregroundColor(.secondary)
                    Spacer()
                }
            }

            Spacer()
        }
        .padding(20)
        .frame(minWidth: 580)
    }

    private var centerPanel: some View {
        VStack(alignment: .leading, spacing: 16) {
            HStack {
                Text("Scanned Modules").font(.system(size: 24, weight: .bold))
                Spacer()
                Picker("Sort", selection: $sortKey) {
                    Text("Date").tag(SortKey.dateDesc)
                    Text("Variant#").tag(SortKey.variant)
                    Text("Speed#").tag(SortKey.speed)
                    Text("Size#").tag(SortKey.size)
                }.pickerStyle(.segmented).frame(maxWidth: 400)
            }

            if sortedResults.isEmpty {
                Spacer()
                Text("No scans yet — scan a barcode or paste a batch").foregroundColor(.secondary).multilineTextAlignment(.center)
                Spacer()
            } else {
                ScrollView {
                    VStack(spacing: 12) {
                        ForEach(sortedResults.reversed()) { r in
                            ScanRowView(result: r,
                                        speedColor: speedColor(_:),
                                        sizeColors: sizeTagColors(_:),
                                        onDelete: { removeResult(id: r.id) })
                        }
                    }
                }
            }
            Spacer()
        }
        .padding(20)
        .frame(minWidth: 650)
    }

    private var rightPanel: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("Variants by Manufacturer").font(.system(size: 20, weight: .bold))
            ScrollView {
                VStack(alignment: .leading, spacing: 10) {
                    ForEach(variantsByManufacturer.keys.sorted(), id: \.self) { m in
                        VStack(alignment: .leading, spacing: 6) {
                            Text(m).font(.headline)
                            if let variants = variantsByManufacturer[m] {
                                ForEach(variants.keys.sorted(), id: \.self) { v in
                                    let vid = variantIndex[v] ?? 0
                                    VariantRowView(variantText: v,
                                                   count: variants[v] ?? 0,
                                                   variantIdText: "V\(vid)",
                                                   onDelete: {
                                                       variantPendingDelete = v
                                                       showDeleteVariantAlert = true
                                                   })
                                }
                            }
                        }
                        .padding(8)
                        .background(RoundedRectangle(cornerRadius: 10).fill(Color(NSColor.windowBackgroundColor)))
                        .overlay(RoundedRectangle(cornerRadius: 10).stroke(Color.secondary.opacity(0.15), lineWidth: 1))
                    }
                }
            }
            Spacer()
        }
        .padding(20)
        .frame(minWidth: 300)
    }

    // ===== Scan Flow =====
    private func processScan() {
        let raw = scanText.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !raw.isEmpty else { scanText = ""; return }
        processRawBarcode(raw)
        scanText = ""
    }

    private func processRawBarcode(_ raw: String) {
        let variant = normalizedVariant(from: raw)
        let speed = detectSpeed(from: raw)
        let size  = detectSize(from: raw)
        let mfg   = detectManufacturer(from: raw)   // now considers customManufacturerMap

        if speed == "Unknown" || size == "Unknown" {
            lastScanned = ScanResult(
                id: UUID(), barcode: raw, normalizedVariant: variant,
                speed: speed, size: size, manufacturer: mfg, date: Date(),
                variantId: assignVariantId(variant ?? "Unknown"),
                speedId: assignSpeedId(speed), sizeId: assignSizeId(size)
            )
            if !pendingUnknowns.contains(raw) { pendingUnknowns.append(raw) }
            mapCode  = variant ?? ""
            mapSpeed = (speed == "Unknown") ? "" : speed
            mapSize  = (size  == "Unknown") ? "" : size
            mapManufacturer = (mfg == "Unknown") ? "" : mfg    // prefill if known
            focusedField = .addCombinedCode
            Sounder.shared.error()
            return
        }

        appendResolvedScan(raw: raw, variant: variant, speed: speed, size: size, manufacturer: mfg, playSound: true)
        focusedField = .scan
    }

    private func appendResolvedScan(raw: String, variant: String?, speed: String, size: String, manufacturer: String, playSound: Bool) {
        let vid = assignVariantId(variant ?? "Unknown")
        let sid = assignSpeedId(speed)
        let zid = assignSizeId(size)
        let result = ScanResult(id: UUID(), barcode: raw, normalizedVariant: variant,
                                speed: speed, size: size, manufacturer: manufacturer,
                                date: Date(), variantId: vid, speedId: sid, sizeId: zid)
        results.append(result)
        lastScanned = result
        if playSound { Sounder.shared.success() }
        saveAll()
    }

    private func processBatch() {
        let lines = batchInput
            .split(whereSeparator: \.isNewline)
            .map { String($0).trimmingCharacters(in: .whitespacesAndNewlines) }
            .filter { !$0.isEmpty }
        for line in lines {
            let variant = normalizedVariant(from: line)
            let speed = detectSpeed(from: line)
            let size  = detectSize(from: line)
            let mfg   = detectManufacturer(from: line)
            if speed != "Unknown", size != "Unknown" {
                appendResolvedScan(raw: line, variant: variant, speed: speed, size: size, manufacturer: mfg, playSound: false)
            } else {
                let r = ScanResult(
                    id: UUID(), barcode: line, normalizedVariant: variant,
                    speed: speed, size: size, manufacturer: mfg, date: Date(),
                    variantId: assignVariantId(variant ?? "Unknown"),
                    speedId: assignSpeedId(speed), sizeId: assignSizeId(size)
                )
                lastScanned = r
                if !pendingUnknowns.contains(line) { pendingUnknowns.append(line) }
            }
        }
        focusedField = .scan
        saveAll()
    }

    // ===== Add Mapping (Speed + Size + optional Manufacturer) =====
    private func addCombinedMapping() {
        let codeRaw  = mapCode.trimmingCharacters(in: .whitespacesAndNewlines)
        let speedRaw = mapSpeed.trimmingCharacters(in: .whitespacesAndNewlines)
        let sizeNorm = normalizeInputSize(mapSize.trimmingCharacters(in: .whitespacesAndNewlines))
        let mfgRaw   = mapManufacturer.trimmingCharacters(in: .whitespacesAndNewlines)

        guard !codeRaw.isEmpty, !speedRaw.isEmpty, sizeNorm != "Unknown" else { return }

        let codeUC = codeRaw.uppercased()
        let canon  = (normalizedVariant(from: codeUC) ?? codeUC)

        func writeKey(_ key: String) {
            customSpeedMap[key] = speedRaw
            customSizeMap[key]  = sizeNorm
            if !mfgRaw.isEmpty { customManufacturerMap[key] = mfgRaw } // only save if provided
        }
        writeKey(codeUC)
        if canon != codeUC { writeKey(canon) }

        saveCustomSpeedMap()
        saveCustomSizeMap()
        saveCustomMfgMap()

        func matches(_ raw: String) -> Bool {
            let u = raw.uppercased()
            if u.contains(codeUC) || u.contains(canon) { return true }
            if let nv = normalizedVariant(from: raw), nv == codeUC || nv == canon { return true }
            return false
        }

        // If last scan was unknown and matches, append/update NOW (silent)
        if let last = lastScanned, (last.speed == "Unknown" || last.size == "Unknown" || last.manufacturer == "Unknown"), matches(last.barcode) {
            let variant = normalizedVariant(from: last.barcode)
            let resolvedSpeed = (last.speed == "Unknown") ? speedRaw : last.speed
            let resolvedSize  = (last.size  == "Unknown") ? sizeNorm : last.size
            let resolvedMfg   = (!mfgRaw.isEmpty ? mfgRaw : last.manufacturer)
            appendResolvedScan(raw: last.barcode, variant: variant, speed: resolvedSpeed, size: resolvedSize, manufacturer: resolvedMfg, playSound: false)
            pendingUnknowns.removeAll { $0 == last.barcode }
        }

        // Backfill Unknowns in results that match (silent)
        for i in results.indices {
            if matches(results[i].barcode) {
                if results[i].speed == "Unknown"  { results[i].speed = speedRaw;  results[i].speedId = assignSpeedId(speedRaw) }
                if results[i].size  == "Unknown"  { results[i].size  = sizeNorm; results[i].sizeId  = assignSizeId(sizeNorm) }
                if results[i].manufacturer == "Unknown", !mfgRaw.isEmpty {
                    results[i] = ScanResult(id: results[i].id,
                                            barcode: results[i].barcode,
                                            normalizedVariant: results[i].normalizedVariant,
                                            speed: results[i].speed,
                                            size: results[i].size,
                                            manufacturer: mfgRaw,
                                            date: results[i].date,
                                            variantId: results[i].variantId,
                                            speedId: results[i].speedId,
                                            sizeId: results[i].sizeId)
                }
            }
        }

        // Append pending unknowns that match (silent)
        var stillPending: [String] = []
        for raw in pendingUnknowns {
            if matches(raw) {
                let manufacturer = !mfgRaw.isEmpty ? mfgRaw : detectManufacturer(from: raw)
                appendResolvedScan(raw: raw, variant: normalizedVariant(from: raw), speed: speedRaw, size: sizeNorm, manufacturer: manufacturer, playSound: false)
            } else {
                stillPending.append(raw)
            }
        }
        pendingUnknowns = stillPending

        // Clear fields
        mapCode = ""; mapSpeed = ""; mapSize = ""; mapManufacturer = ""
        focusedField = .scan
        saveAll()
    }

    // ===== Mapping Removal =====
    private var mappingKeys: [String] {
        Array(Set(customSpeedMap.keys).union(customSizeMap.keys).union(customManufacturerMap.keys)).sorted()
    }

    private func performDeleteMapping() {
        guard let key = mappingPendingDeleteKey else { return }
        customSpeedMap.removeValue(forKey: key)
        customSizeMap.removeValue(forKey: key)
        customManufacturerMap.removeValue(forKey: key) // NEW
        mappingPendingDeleteKey = nil
        saveCustomSpeedMap()
        saveCustomSizeMap()
        saveCustomMfgMap()
    }

    // ===== Sorting/Grouping =====
    private var sortedResults: [ScanResult] {
        var arr = results
        switch sortKey {
        case .dateDesc: arr.sort { $0.date < $1.date }
        case .variant:  arr.sort { $0.variantId < $1.variantId }
        case .speed:    arr.sort { $0.speedId < $1.speedId }
        case .size:     arr.sort { $0.sizeId < $1.sizeId }
        }
        return arr
    }

    private var variantsByManufacturer: [String: [String: Int]] {
        var dict: [String: [String: Int]] = [:]
        for r in results {
            guard let variant = r.normalizedVariant else { continue }
            dict[r.manufacturer, default: [:]][variant, default: 0] += 1
        }
        return dict
    }

    // ===== Deletion Helpers =====
    private func performDeleteVariant() {
        guard let v = variantPendingDelete else { return }
        results.removeAll { $0.normalizedVariant == v }
        if results.isEmpty { lastScanned = nil; resetIds() } else { lastScanned = results.last }
        variantPendingDelete = nil
        saveResults()
    }

    private func removeResult(id: UUID) {
        results.removeAll { $0.id == id }
        if results.isEmpty { lastScanned = nil; resetIds() }
        else if lastScanned?.id == id { lastScanned = results.last }
        saveResults()
    }

    // ===== ID Helpers =====
    private func assignVariantId(_ variant: String) -> Int {
        if let id = variantIndex[variant] { return id }
        let id = nextVariantId; variantIndex[variant] = id; nextVariantId += 1; saveIndexes(); return id
    }
    private func assignSpeedId(_ speed: String) -> Int {
        if let id = speedIndex[speed] { return id }
        let id = nextSpeedId; speedIndex[speed] = id; nextSpeedId += 1; saveIndexes(); return id
    }
    private func assignSizeId(_ size: String) -> Int {
        if let id = sizeIndex[size] { return id }
        let id = nextSizeId; sizeIndex[size] = id; nextSizeId += 1; saveIndexes(); return id
    }
    private func resetIds() {
        variantIndex.removeAll(); speedIndex.removeAll(); sizeIndex.removeAll()
        nextVariantId = 1; nextSpeedId = 1; nextSizeId = 1
        saveIndexes()
    }

    // ===== Detection =====
    private func detectManufacturer(from barcode: String) -> String {
        let u = barcode.uppercased()
        // Prefer custom mapping (exact canonical)
        if let nv = normalizedVariant(from: u), let m = customManufacturerMap[nv], !m.isEmpty { return m }
        // Custom substring
        for (k, m) in customManufacturerMap where u.contains(k.uppercased()) && !m.isEmpty { return m }

        // Heuristic fallbacks
        if u.hasPrefix("HMA") || u.contains("HYNIX") || u.contains("SKHYNIX") { return "SK hynix" }
        if u.hasPrefix("MTA") || u.contains("MICRON") || u.hasPrefix("MT") { return "Micron" }
        if u.hasPrefix("M393") || u.contains("SAMSUNG") { return "Samsung" }
        if u.contains("CRUCIAL") { return "Crucial" }
        return "Unknown"
    }

    // Canonical variant (up to +XX if present)
    private func normalizedVariant(from s: String) -> String? {
        let u = s.uppercased()
        if let plus = u.firstIndex(of: "+"), plus < u.endIndex {
            let next = u.index(after: plus)
            if next < u.endIndex {
                let end = u.index(next, offsetBy: 2, limitedBy: u.endIndex) ?? u.endIndex
                return String(u[..<end])
            }
        }
        let cleaned = u.replacingOccurrences(of: #"[^A-Z0-9\+]"#, with: "", options: .regularExpression)
        return cleaned.isEmpty ? nil : cleaned
    }

    private func detectSpeed(from barcode: String) -> String {
        let u = barcode.uppercased()
        if let nv = normalizedVariant(from: u), let v = customSpeedMap[nv] { return v }
        for (k, v) in fullSpeedMap where u.contains(k.uppercased()) { return v }
        if u.contains("3200") { return "3200 MHz" }
        if u.contains("2933") { return "2933 MHz" }
        if u.contains("2666") { return "2666 MHz" }
        if u.contains("2400") { return "2400 MHz" }
        if u.contains("2133") { return "2133 MHz" }
        return "Unknown"
    }

    private func detectSize(from barcode: String) -> String {
        let u = barcode.uppercased()
        if let nv = normalizedVariant(from: u), let v = customSizeMap[nv] { return normalizeInputSize(v) }
        for (k, v) in customSizeMap where u.contains(k.uppercased()) { return normalizeInputSize(v) }
        if let tb = firstMatch(in: u, pattern: #"(\d{1,3})(?:\.\d+)?\s*TB\b"#), let t = Int(tb) { return "\(t * 1024) GB" }
        if let gb = firstMatch(in: u, pattern: #"(\d{1,4})\s*GB\b"#) { return normalizeGB(Int(gb)) }
        if let g  = firstMatch(in: u, pattern: #"(\d{1,4})\s*G\b"#)   { return normalizeGB(Int(g)) }
        if let mb = firstMatch(in: u, pattern: #"(\d{3,6})\s*MB\b"#), let m = Int(mb) { return m % 1024 == 0 ? "\(m/1024) GB" : "\(m) MB" }
        if let gb2 = firstMatch(in: u, pattern: #"(?<!\d)(\d{1,4})(?=G|GB)(?:G|GB)"#), let n = Int(gb2) { return normalizeGB(Int(n)) }
        for n in [8,16,24,32,48,64,96,128,192,256,384,512,768,1024] where u.contains("\(n)G") || u.contains("\(n)GB") { return "\(n) GB" }
        return "Unknown"
    }

    private func firstMatch(in text: String, pattern: String) -> String? {
        do {
            let r = try NSRegularExpression(pattern: pattern, options: [])
            let range = NSRange(text.startIndex..<text.endIndex, in: text)
            if let m = r.firstMatch(in: text, options: [], range: range),
               m.numberOfRanges >= 2, let rr = Range(m.range(at: 1), in: text) {
                return String(text[rr])
            }
        } catch {}
        return nil
    }

    private func normalizeInputSize(_ s: String) -> String {
        let u = s.uppercased().trimmingCharacters(in: .whitespacesAndNewlines)
        if let mbStr = firstMatch(in: u, pattern: #"^(\d{3,6})\s*MB$"#), let mb = Int(mbStr) { return mb % 1024 == 0 ? "\(mb/1024) GB" : "\(mb) MB" }
        if let gStr = firstMatch(in: u, pattern: #"^(\d{1,4})(?:\s*G|GB|\s*GB)?$"#) { return normalizeGB(Int(gStr)) }
        if let tb = firstMatch(in: u, pattern: #"^(\d{1,3})(?:\.\d+)?\s*TB$"#), let t = Int(tb) { return "\(t * 1024) GB" }
        return u.isEmpty ? "Unknown" : u
    }
    private func normalizeGB(_ n: Int?) -> String { guard let n = n else { return "Unknown" }; return "\(n) GB" }

    // ===== CSV / Persistence =====
    private func copyResultsAsCSV() {
        var lines = ["variant_id,speed_id,size_id,variant,speed,size,manufacturer,barcode,timestamp"]
        for r in results {
            let ts = ISO8601DateFormatter().string(from: r.date)
            lines.append("\(r.variantId),\(r.speedId),\(r.sizeId),\(r.normalizedVariant ?? ""),\(r.speed),\(r.size),\(r.manufacturer),\(r.barcode),\(ts)")
        }
        NSPasteboard.general.clearContents()
        NSPasteboard.general.setString(lines.joined(separator: "\n"), forType: .string)
    }

    private func saveResults() {
        if let data = try? JSONEncoder().encode(results) { UserDefaults.standard.set(data, forKey: UDKey.results) }
    }
    private func saveIndexes() {
        if let v = try? JSONEncoder().encode(variantIndex) { UserDefaults.standard.set(v, forKey: UDKey.variantIndex) }
        if let s = try? JSONEncoder().encode(speedIndex)   { UserDefaults.standard.set(s, forKey: UDKey.speedIndex) }
        if let z = try? JSONEncoder().encode(sizeIndex)    { UserDefaults.standard.set(z, forKey: UDKey.sizeIndex) }
        UserDefaults.standard.set(nextVariantId, forKey: UDKey.nextVariantId)
        UserDefaults.standard.set(nextSpeedId,   forKey: UDKey.nextSpeedId)
        UserDefaults.standard.set(nextSizeId,    forKey: UDKey.nextSizeId)
    }
    private func saveCustomSpeedMap() { if let data = try? JSONEncoder().encode(customSpeedMap) { UserDefaults.standard.set(data, forKey: UDKey.customSpeedMap) } }
    private func saveCustomSizeMap()  { if let data = try? JSONEncoder().encode(customSizeMap)  { UserDefaults.standard.set(data, forKey: UDKey.customSizeMap) } }
    private func saveCustomMfgMap()   { if let data = try? JSONEncoder().encode(customManufacturerMap) { UserDefaults.standard.set(data, forKey: UDKey.customMfgMap) } }
    private func saveAll() { saveResults(); saveIndexes(); saveCustomSpeedMap(); saveCustomSizeMap(); saveCustomMfgMap() }

    private func loadAll() {
        if let d = UserDefaults.standard.data(forKey: UDKey.results),
           let dec = try? JSONDecoder().decode([ScanResult].self, from: d) { results = dec; lastScanned = results.last }
        if let d = UserDefaults.standard.data(forKey: UDKey.customSpeedMap),
           let dec = try? JSONDecoder().decode([String: String].self, from: d) { customSpeedMap = dec }
        if let d = UserDefaults.standard.data(forKey: UDKey.customSizeMap),
           let dec = try? JSONDecoder().decode([String: String].self, from: d) { customSizeMap = dec }
        if let d = UserDefaults.standard.data(forKey: UDKey.customMfgMap),
           let dec = try? JSONDecoder().decode([String: String].self, from: d) { customManufacturerMap = dec }
        if let d = UserDefaults.standard.data(forKey: UDKey.variantIndex),
           let dec = try? JSONDecoder().decode([String: Int].self, from: d) { variantIndex = dec }
        if let d = UserDefaults.standard.data(forKey: UDKey.speedIndex),
           let dec = try? JSONDecoder().decode([String: Int].self, from: d) { speedIndex = dec }
        if let d = UserDefaults.standard.data(forKey: UDKey.sizeIndex),
           let dec = try? JSONDecoder().decode([String: Int].self, from: d) { sizeIndex = dec }
        let nv = UserDefaults.standard.integer(forKey: UDKey.nextVariantId)
        let ns = UserDefaults.standard.integer(forKey: UDKey.nextSpeedId)
        let nz = UserDefaults.standard.integer(forKey: UDKey.nextSizeId)
        nextVariantId = nv == 0 ? max((variantIndex.values.max() ?? 0) + 1, 1) : nv
        nextSpeedId   = ns == 0 ? max((speedIndex.values.max()   ?? 0) + 1, 1) : ns
        nextSizeId    = nz == 0 ? max((sizeIndex.values.max()    ?? 0) + 1, 1) : nz
    }

    // ===== Colors =====
    private func speedColor(_ s: String) -> Color { if s.contains("2666") { return .blue }; if s.contains("3200") { return .yellow }; if s == "Unknown" { return .red }; return .primary }
    private func sizeTagColors(_ s: String) -> (fg: Color, bg: Color) {
        switch s {
        case "32 GB": return (.white, .blue)
        case "16 GB": return (.black, .yellow)
        case "8 GB":  return (.white, .orange)
        default: return (.primary, Color.secondary.opacity(0.12))
        }
    }
}

// ===== View Pieces =====
struct LastScanBox: View {
    let last: ScanResult
    let speedColor: (String) -> Color
    let sizeColors: (String) -> (fg: Color, bg: Color)
    var body: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("Last Scan:").font(.headline)
            VStack(alignment: .leading, spacing: 10) {
                HStack(spacing: 16) {
                    Text(last.speed).font(.system(size: 36, weight: .bold, design: .rounded)).foregroundColor(speedColor(last.speed))
                    BigIdTag("V\(last.variantId)")
                    Tag(text: last.manufacturer)
                    let sz = sizeColors(last.size); Tag(text: last.size, fg: sz.fg, bg: sz.bg)
                    TinyIdTag("S\(last.speedId)"); TinyIdTag("Z\(last.sizeId)")
                }
                Text(last.barcode).font(.system(size: 20, design: .monospaced))
                if let v = last.normalizedVariant { Text("Variant: \(v)").font(.caption).foregroundColor(.secondary) }
                Text(last.date, style: .time).font(.system(size: 13)).foregroundColor(.secondary)
            }
            .padding(18)
            .background(RoundedRectangle(cornerRadius: 12).fill(Color(NSColor.windowBackgroundColor)))
            .overlay(RoundedRectangle(cornerRadius: 12).stroke(Color.secondary.opacity(0.3), lineWidth: 1))
        }
    }
}

struct ScanRowView: View {
    let result: ScanResult
    let speedColor: (String) -> Color
    let sizeColors: (String) -> (fg: Color, bg: Color)
    let onDelete: () -> Void
    var body: some View {
        HStack(alignment: .center, spacing: 24) {
            VStack(alignment: .leading, spacing: 8) {
                HStack(spacing: 14) {
                    Text(result.speed).font(.system(size: 24, weight: .bold)).foregroundColor(speedColor(result.speed))
                    Tag(text: result.manufacturer)
                    let sz = sizeColors(result.size); Tag(text: result.size, fg: sz.fg, bg: sz.bg)
                    TinyIdTag("V\(result.variantId)"); TinyIdTag("S\(result.speedId)"); TinyIdTag("Z\(result.sizeId)")
                }
                Text(result.barcode).font(.system(size: 18, design: .monospaced))
                if let v = result.normalizedVariant { Text("Variant: \(v)").font(.caption).foregroundColor(.secondary) }
                Text(result.date, style: .time).font(.system(size: 13)).foregroundColor(.secondary)
            }
            Spacer()
            Button(action: onDelete) { Image(systemName: "xmark.circle").font(.system(size: 22)) }
                .buttonStyle(BorderlessButtonStyle())
        }
        .padding(16)
        .background(RoundedRectangle(cornerRadius: 10).fill(Color(NSColor.windowBackgroundColor)))
        .overlay(RoundedRectangle(cornerRadius: 10).stroke(Color.secondary.opacity(0.15), lineWidth: 1))
    }
}

struct VariantRowView: View {
    let variantText: String
    let count: Int
    let variantIdText: String
    let onDelete: () -> Void
    var body: some View {
        HStack(spacing: 8) {
            BigIdTag(variantIdText)
            Text(variantText).font(.system(size: 12, design: .monospaced))
            Spacer()
            Text("×\(count)").font(.caption).foregroundColor(.secondary)
            Button(action: onDelete) { Image(systemName: "trash") }.buttonStyle(.borderless)
        }
        .padding(.vertical, 4)
    }
}

// ===== UI Helpers =====
struct Tag: View {
    let text: String; var fg: Color? = nil; var bg: Color? = nil
    var body: some View {
        Text(text)
            .font(.caption)
            .foregroundColor(fg ?? .primary)
            .padding(.vertical, 3)
            .padding(.horizontal, 8)
            .background(Capsule().fill(bg ?? Color.secondary.opacity(0.12)))
    }
}
struct TinyIdTag: View { let text: String; init(_ t:String){text=t}
    var body: some View {
        Text(text)
            .font(.system(size:10,weight:.semibold,design:.monospaced))
            .padding(.vertical,2)
            .padding(.horizontal,6)
            .background(Capsule().fill(Color.secondary.opacity(0.15)))
    }
}
struct BigIdTag: View { let text: String; init(_ t:String){text=t}
    var body: some View {
        Text(text)
            .font(.system(size:18,weight:.bold,design:.rounded))
            .padding(.vertical,4)
            .padding(.horizontal,10)
            .background(Capsule().fill(Color.accentColor.opacity(0.18)))
            .overlay(Capsule().stroke(Color.accentColor.opacity(0.9), lineWidth: 1))
    }
}

