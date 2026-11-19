/**
 * macOS Vision OCR helper for pdf2docx (see src/ocr.ts).
 *
 * Reads a PNG from stdin, writes recognized lines as JSON to stdout:
 *   {"paper":0.93,"lines":[{"t":"text","c":0.83,"b":[x0,y0,x1,y1],"chars":[{"t":"字","b":[...]}]}]}
 * "paper" is the near-white pixel share (photo-vs-document signal).
 * Boxes are normalized 0–1, origin bottom-left, y up (Vision's native space,
 * which matches PDF page space directly).
 *
 * argv[1] (optional): comma-separated recognition language hints, e.g.
 * "zh-Hans,en-US". Without it Vision auto-detects the language.
 *
 * Build: swiftc -O vision-ocr.swift -o vision-ocr
 *
 * Caveat learned the hard way: Vision silently returns ZERO results for
 * transparent-background bitmaps (text lives in the alpha channel, RGB is
 * black-on-black). PDFium page renders are white-backed, so the pipeline is
 * safe; standalone callers must flatten alpha first.
 */
import AppKit
import Foundation
import Vision

let stdinData = FileHandle.standardInput.readDataToEndOfFile()
guard let img = NSImage(data: stdinData),
      let cg = img.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
    FileHandle.standardError.write("cannot decode input image\n".data(using: .utf8)!)
    exit(2)
}

let request = VNRecognizeTextRequest()
request.recognitionLevel = .accurate
request.usesLanguageCorrection = true
if #available(macOS 13.0, *) {
    request.automaticallyDetectsLanguage = true
}
if CommandLine.arguments.count > 1, !CommandLine.arguments[1].isEmpty {
    request.recognitionLanguages = CommandLine.arguments[1].split(separator: ",").map(String.init)
}

let handler = VNImageRequestHandler(cgImage: cg, options: [:])
do {
    try handler.perform([request])
} catch {
    FileHandle.standardError.write("recognition failed: \(error)\n".data(using: .utf8)!)
    exit(3)
}

/// paper-tone share: fraction of sampled pixels whose min channel reads as
/// near-white paper. Document scans are paper-dominated; photos are not —
/// the TS policy layer uses this to keep photos as images.
func paperShare(_ cg: CGImage) -> Double {
    let w = cg.width, h = cg.height
    guard w > 0, h > 0,
          let ctx = CGContext(data: nil, width: w, height: h, bitsPerComponent: 8,
                              bytesPerRow: w * 4, space: CGColorSpaceCreateDeviceRGB(),
                              bitmapInfo: CGImageAlphaInfo.noneSkipLast.rawValue) else { return 1 }
    // white-fill first: transparent sources must read as paper, not black
    ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
    ctx.fill(CGRect(x: 0, y: 0, width: w, height: h))
    ctx.draw(cg, in: CGRect(x: 0, y: 0, width: w, height: h))
    guard let data = ctx.data else { return 1 }
    let px = data.bindMemory(to: UInt8.self, capacity: w * h * 4)
    let stride = max(1, (w * h) / 200_000)
    var paper = 0, total = 0
    var i = 0
    while i < w * h {
        let o = i * 4
        if min(px[o], min(px[o + 1], px[o + 2])) >= 200 { paper += 1 }
        total += 1
        i += stride
    }
    return total > 0 ? Double(paper) / Double(total) : 1
}

func rect(_ b: CGRect) -> [Double] {
    [Double(b.minX), Double(b.minY), Double(b.maxX), Double(b.maxY)]
}

var lines: [[String: Any]] = []
for obs in request.results ?? [] {
    guard let cand = obs.topCandidates(1).first, !cand.string.isEmpty else { continue }
    var chars: [[String: Any]] = []
    let s = cand.string
    var idx = s.startIndex
    while idx < s.endIndex {
        let next = s.index(after: idx)
        if let charBox = try? cand.boundingBox(for: idx ..< next)?.boundingBox {
            chars.append(["t": String(s[idx ..< next]), "b": rect(charBox)])
        } else {
            chars.append(["t": String(s[idx ..< next]), "b": [0.0, 0.0, 0.0, 0.0]])
        }
        idx = next
    }
    lines.append([
        "t": s,
        "c": Double(cand.confidence),
        "b": rect(obs.boundingBox),
        "chars": chars,
    ])
}

let out = try JSONSerialization.data(withJSONObject: [
    "lines": lines,
    "paper": paperShare(cg),
] as [String: Any])
FileHandle.standardOutput.write(out)
