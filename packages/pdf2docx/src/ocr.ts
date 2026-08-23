/**
 * Local OCR recovery for scanned pages (no text layer + a page-covering
 * image). The recognition engine is caller-supplied and platform-specific
 * (macOS Vision, Windows.Media.Ocr, ONNX models…) — this module owns only the
 * conversion policy: confidence gates, coordinate mapping, and PdfChar
 * synthesis, so the recovered lines flow through the exact same layout
 * analysis (columns / tables / paragraphs) as a digital text layer.
 *
 * OCR is strictly additive: pages with a real text layer never enter this
 * path (digital extraction beats recognition), and any doubt — engine
 * failure, too little text, low confidence, empty analysis — falls back to
 * today's behavior, the full-page bitmap export.
 */
import { analyzePage } from './analyze'
import type { ExtractedPage } from './extract'
import type { Rect } from './geometry'
import type { IrPage, PdfChar } from './ir'
import { isEastAsianScript, scriptOf } from './script'

/** one recognized character; box normalized to 0–1, origin bottom-left, y up */
export interface OcrChar {
  text: string
  box: Rect
}

/** one recognized text line */
export interface OcrLine {
  text: string
  /** engine confidence, 0–1 */
  confidence: number
  /** normalized box (0–1 of page size, origin bottom-left, y up) */
  box: Rect
  /** per-char boxes when the engine provides them (better geometry) */
  chars?: OcrChar[]
}

/** one page's recognition output */
export interface OcrRecognition {
  lines: OcrLine[]
  /**
   * near-white pixel share of the page render (0–1), when the engine can
   * measure it. Document scans are paper-dominated; photos are not — below
   * the gate the page stays an image even if it contains incidental text
   * (signs, UI labels, captions).
   */
  paperShare?: number
}

/**
 * Synchronous recognition engine: page render PNG in, recognition out.
 * `null` means the engine is unavailable or failed — the page falls back to
 * its bitmap. Must be sync: the pipeline runs inside a sync PDFium document
 * scope (a helper-binary engine should use spawnSync).
 */
export type OcrEngine = (
  png: Uint8Array,
  page: { widthPt: number; heightPt: number },
) => OcrRecognition | null

/** below this many recognized non-space chars the "scan" is likely a photo */
const OCR_MIN_CHARS = 8
/** mean line confidence below this ships the bitmap instead (garbled text is
 * worse than a faithful image) */
const OCR_MIN_MEAN_CONFIDENCE = 0.55
/** individual lines below this confidence are dropped (stray photo texture) */
const OCR_LINE_MIN_CONFIDENCE = 0.3
/** a page render less paper-toned than this is a photo, not a document scan —
 * it keeps its bitmap even when it contains real text (UI labels, signs) */
const OCR_MIN_PAPER_SHARE = 0.45
/** recognized text boxes must cover this share of the page area: a photo on
 * a white backdrop with one caption is not a document */
const OCR_MIN_TEXT_COVERAGE = 0.004

/** glyph-box share of the OCR line box height, by script majority: CJK glyphs
 * fill most of the line, Latin ascender/descender bands leave slack */
const FONT_SIZE_SHARE_CJK = 0.9
const FONT_SIZE_SHARE_LATIN = 0.72

/** default output fonts (the source font is unknowable from pixels) */
const OCR_FONT_CJK = 'SimSun'
const OCR_FONT_LATIN = 'Times New Roman'

/** relative advance weights for distributing chars along a line when the
 * engine gives no per-char boxes */
function advanceWeight(code: number): number {
  const script = scriptOf(code)
  if (isEastAsianScript(script)) return 1
  const ch = String.fromCodePoint(code)
  if (ch === ' ') return 0.4
  if (/[0-9A-Z]/.test(ch)) return 0.6
  if (/[.,:;!'"|()[\]{}]/.test(ch)) return 0.3
  return 0.5
}

function eastAsianShare(text: string): number {
  let east = 0
  let total = 0
  for (const ch of text) {
    if (ch.trim() === '') continue
    total++
    if (isEastAsianScript(scriptOf(ch.codePointAt(0) ?? 0))) east++
  }
  return total > 0 ? east / total : 0
}

/** scale a normalized (0–1) rect into page points */
function toPageRect(b: Rect, widthPt: number, heightPt: number): Rect {
  return { x0: b.x0 * widthPt, y0: b.y0 * heightPt, x1: b.x1 * widthPt, y1: b.y1 * heightPt }
}

function synthChar(
  text: string,
  x0: number,
  x1: number,
  baselineY: number,
  fontSize: number,
  cjkLine: boolean,
): PdfChar {
  const code = text.codePointAt(0) ?? 0
  return {
    code,
    text,
    box: { x0, x1, y0: baselineY - fontSize * 0.21, y1: baselineY + fontSize * 0.72 },
    looseBox: { x0, x1, y0: baselineY - fontSize * 0.25, y1: baselineY + fontSize * 0.95 },
    originX: x0,
    originY: baselineY,
    angle: 0,
    fontSize,
    fontWeight: 400,
    fontFamily: cjkLine ? OCR_FONT_CJK : OCR_FONT_LATIN,
    italic: false,
    color: '000000',
    isGenerated: false,
    isHyphen: false,
    script: scriptOf(code),
  }
}

/** synthesize analysis-layer chars from recognized lines (exported for tests) */
export function ocrLinesToChars(lines: OcrLine[], widthPt: number, heightPt: number): PdfChar[] {
  const chars: PdfChar[] = []
  for (const line of lines) {
    const box = toPageRect(line.box, widthPt, heightPt)
    const lineH = box.y1 - box.y0
    const lineW = box.x1 - box.x0
    if (lineH <= 0 || lineW <= 0 || line.text.trim() === '') continue
    const cjkLine = eastAsianShare(line.text) >= 0.5
    const fontSize =
      Math.round(lineH * (cjkLine ? FONT_SIZE_SHARE_CJK : FONT_SIZE_SHARE_LATIN) * 2) / 2
    if (fontSize <= 0) continue
    // baseline: glyph boxes sit 0.21 em below / 0.72 em above it (see synthChar)
    const baselineY = box.y0 + (lineH - fontSize * 0.93) / 2 + fontSize * 0.21

    const glyphs = [...line.text]
    // Raw engine char boxes are NOT usable as glyph boxes: Vision's boxes
    // overlap heavily on proportional text (a narrow 'i' can sit entirely
    // inside its neighbor's box), and the analysis layer's overlap dedup then
    // drops real characters. Engine boxes are trusted only for WORD-SEGMENT
    // anchors (real word gaps — form labels vs values — must survive); glyphs
    // inside a segment are distributed by advance weight, which is
    // non-overlapping by construction.
    const engineBoxes =
      line.chars &&
      line.chars.length === glyphs.length &&
      line.chars.every((c, i) => glyphs[i]!.trim() === '' || c.box.x1 > c.box.x0)
        ? line.chars
        : null

    /** emit glyphs [from, to) distributed across [x0, x1] by advance weight */
    const emitSegment = (from: number, to: number, x0: number, x1: number): void => {
      const seg = glyphs.slice(from, to)
      const weights = seg.map((g) => advanceWeight(g.codePointAt(0) ?? 0))
      const total = weights.reduce((a, b) => a + b, 0)
      if (total <= 0 || x1 <= x0) return
      let x = x0
      for (let i = 0; i < seg.length; i++) {
        const w = (weights[i]! / total) * (x1 - x0)
        const g = seg[i]!
        chars.push(synthChar(g.trim() === '' ? ' ' : g, x, x + w, baselineY, fontSize, cjkLine))
        x += w
      }
    }

    if (!engineBoxes) {
      emitSegment(0, glyphs.length, box.x0, box.x1)
      continue
    }
    // word segments: maximal runs of non-space glyphs, anchored to the engine
    // extents of their first/last glyph (clamped monotone); spaces span the
    // gaps between segments
    let cursor = box.x0
    let i = 0
    while (i < glyphs.length) {
      if (glyphs[i]!.trim() === '') {
        i++
        continue
      }
      let j = i
      while (j < glyphs.length && glyphs[j]!.trim() !== '') j++
      const segX0 = Math.max(cursor, toPageRect(engineBoxes[i]!.box, widthPt, heightPt).x0)
      const segX1 = Math.max(
        segX0 + fontSize * 0.2,
        toPageRect(engineBoxes[j - 1]!.box, widthPt, heightPt).x1,
      )
      if (i > 0 && segX0 > cursor) {
        chars.push(synthChar(' ', cursor, segX0, baselineY, fontSize, cjkLine))
      }
      emitSegment(i, j, segX0, segX1)
      cursor = segX1
      i = j
    }
  }
  return chars
}

export interface OcrPageResult {
  page: IrPage
  /** mean line confidence of the kept lines */
  confidence: number
}

/**
 * Try to recover a scanned page's text through the engine. Returns the
 * re-analyzed page, or null when any gate fails (caller keeps the bitmap
 * fallback exactly as before).
 *
 * @param hiRender optional higher-resolution render fed to the engine only
 * (recognition quality rises with dpi; the fallback render stays as-is)
 */
export function tryOcrScannedPage(
  extracted: ExtractedPage,
  engine: OcrEngine,
  hiRender?: { data: Uint8Array },
): OcrPageResult | null {
  const render = hiRender ?? extracted.render
  if (!render) return null
  let recognition: OcrRecognition | null
  try {
    recognition = engine(render.data, { widthPt: extracted.widthPt, heightPt: extracted.heightPt })
  } catch {
    return null
  }
  if (!recognition || recognition.lines.length === 0) return null
  if (recognition.paperShare !== undefined && recognition.paperShare < OCR_MIN_PAPER_SHARE) {
    return null
  }
  const kept = recognition.lines.filter(
    (l) => l.confidence >= OCR_LINE_MIN_CONFIDENCE && l.text.trim() !== '',
  )
  if (kept.length === 0) return null
  const coverage = kept.reduce(
    (a, l) => a + Math.max(0, l.box.x1 - l.box.x0) * Math.max(0, l.box.y1 - l.box.y0),
    0,
  )
  if (coverage < OCR_MIN_TEXT_COVERAGE) return null
  const meanConfidence = kept.reduce((a, l) => a + l.confidence, 0) / kept.length
  if (meanConfidence < OCR_MIN_MEAN_CONFIDENCE) return null
  const chars = ocrLinesToChars(kept, extracted.widthPt, extracted.heightPt)
  if (chars.filter((c) => c.text.trim() !== '').length < OCR_MIN_CHARS) return null

  const synthetic: ExtractedPage = {
    ...extracted,
    chars,
    // the page-covering scan bitmap is superseded by the recovered text; other
    // artwork on a scanned page lives inside that bitmap, so nothing to keep
    images: [],
    paths: [],
    scanned: false,
    degraded: false,
    degradedReason: undefined,
    render: undefined,
    bgRender: undefined,
    textLost: false,
    badUnicodeRatio: 0,
  }
  const page = analyzePage(synthetic)
  if (page.scanned || page.degraded) return null
  page.ocrRecovered = true
  return { page, confidence: meanConfidence }
}
