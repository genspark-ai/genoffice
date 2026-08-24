/**
 * OCR text layer for scanned pages (issue #119): the platform system engine
 * returns line/word boxes; this module converts them into (a) a synthetic,
 * selectable transparent text overlay and (b) a PageEntry that patches the
 * search index — everything downstream (search, selection quads, markups,
 * AI tools) then works on scanned pages unchanged.
 *
 * Word boxes are stored in PDF user space (like markups), so unsaved page
 * rotations re-project them instead of invalidating the recognition.
 */
import type { ReactElement } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfOcrLine } from '../shared/ipc'
import { geomDispSize, pdfRectToCss, viewToPdf } from './annotations'
import type { PageGeom } from './annotations'
import type { PageEntry } from './search'
import { measurePt } from './text-wrap'
import { isNoSpaceScript, scriptOf } from '../../../../packages/pdf2docx/src/script'

export interface OcrWord {
  text: string
  /** PDF user space [x1,y1,x2,y2] */
  rect: [number, number, number, number]
  /** Render a space after this word so DOM copy keeps real word gaps */
  spaceAfter: boolean
}

export interface OcrPageData {
  entry: PageEntry
  words: OcrWord[]
}

/** Vision emits degenerate boxes for separators (spaces): full-image or zero-width */
const isSeparatorBox = (b: [number, number, number, number]): boolean =>
  b[2] - b[0] <= 0 || (b[0] === 0 && b[1] === 1 && b[2] === 0 && b[3] === 1)

/** Ignore noise the engine is unsure about (stray marks, bleed-through) */
const MIN_LINE_CONFIDENCE = 0.3

/** normalized bottom-left box (relative to the rendered display image) → display-space box at scale 1 */
function boxToDisp(
  b: [number, number, number, number],
  disp: { width: number; height: number },
): { left: number; top: number; right: number; bottom: number } {
  return {
    left: b[0] * disp.width,
    top: (1 - b[3]) * disp.height,
    right: b[2] * disp.width,
    bottom: (1 - b[1]) * disp.height,
  }
}

function dispToPdfRect(
  geom: PageGeom,
  d: { left: number; top: number; right: number; bottom: number },
): [number, number, number, number] {
  const [ax, ay] = viewToPdf(geom, d.left, d.top)
  const [bx, by] = viewToPdf(geom, d.right, d.bottom)
  return [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)]
}

interface LineWord {
  text: string
  box: [number, number, number, number]
  /** A real separator (space) preceded this word in the OCR stream. CJK scripts
      emit one word per character with no separators — joining them with spaces
      would corrupt the searchable text (`使 用 价 值` never matches `使用价值`). */
  gap: boolean
}

/** Split one OCR line into words: char boxes when present (chars of a word share
    its box; separators carry degenerate boxes), whole-line fallback otherwise */
function lineWords(line: PdfOcrLine): LineWord[] {
  if (!line.chars || line.chars.length === 0) {
    const text = line.text.trim()
    return text ? [{ text, box: line.box, gap: false }] : []
  }
  const words: LineWord[] = []
  let cur: LineWord | null = null
  let sep = false
  for (const c of line.chars) {
    if (isSeparatorBox(c.box) || c.text.trim() === '') {
      if (cur) words.push(cur)
      cur = null
      sep = true
      continue
    }
    if (cur && c.box[0] === cur.box[0] && c.box[2] === cur.box[2]) cur.text += c.text
    else {
      if (cur) words.push(cur)
      cur = { text: c.text, box: [...c.box] as [number, number, number, number], gap: sep }
      sep = false
    }
  }
  if (cur) words.push(cur)
  return words.length > 0 ? lineWordsFallback(words, line) : []
}

/** Guard against helpers whose char boxes don't cover the text: fall back to the line box */
function lineWordsFallback(words: LineWord[], line: PdfOcrLine): LineWord[] {
  const joined = words.map((w) => w.text).join(' ')
  // char boxes lost most of the text (some engines only box a prefix): keep the line whole
  if (joined.length < line.text.trim().length * 0.5) {
    const text = line.text.trim()
    return text ? [{ text, box: line.box, gap: false }] : []
  }
  return words
}

/** No-space scripts (cjk/kana/thai, NOT hangul — Korean spaces are real) via
    the shared pdf2docx classifier, so the viewer and the converter agree. */
const noSpaceEdge = (code: number | undefined): boolean =>
  code !== undefined && isNoSpaceScript(scriptOf(code))
const endsNoSpaceScript = (text: string): boolean => noSpaceEdge([...text].pop()?.codePointAt(0))
const startsNoSpaceScript = (text: string): boolean => noSpaceEdge(text.codePointAt(0))

/** A separator between two no-space-script neighbors is engine word
    segmentation, not a real space (Windows OcrLine.Text space-joins its CJK
    words) — writing it into the index would break search the same way
    unconditional joining did. */
const realSpace = (prev: LineWord, next: LineWord): boolean =>
  next.gap && !(endsNoSpaceScript(prev.text) && startsNoSpaceScript(next.text))

/** OCR lines (normalized, display orientation) → search-index entry + overlay words */
export function buildOcrPageData(lines: PdfOcrLine[], geom: PageGeom): OcrPageData | null {
  const disp = geomDispSize(geom)
  let text = ''
  const items: PageEntry['items'] = []
  const words: OcrWord[] = []
  for (const line of lines) {
    if (line.confidence < MIN_LINE_CONFIDENCE) continue
    const lw = lineWords(line)
    if (lw.length === 0) continue
    for (let i = 0; i < lw.length; i++) {
      const w = lw[i]!
      const rect = dispToPdfRect(geom, boxToDisp(w.box, disp))
      if (i > 0 && realSpace(lw[i - 1]!, w)) text += ' '
      const start = text.length
      text += w.text
      items.push({
        start,
        end: text.length,
        x: rect[0],
        y: rect[1],
        w: rect[2] - rect[0],
        h: rect[3] - rect[1],
      })
      const last = i === lw.length - 1
      words.push({
        text: w.text,
        rect,
        spaceAfter: last ? !endsNoSpaceScript(w.text) : realSpace(w, lw[i + 1]!),
      })
    }
    text += '\n'
  }
  if (items.length === 0) return null
  return { entry: { text, lower: text.toLowerCase(), items }, words }
}

/** True when a page has effectively no extractable text (scanned candidate).
    Shared with the AI read_pages fallback so both sides agree on what "scanned" means. */
export const isScannedText = (text: string): boolean => text.replace(/\s/g, '').length < 8

export const isScannedEntry = (entry: PageEntry): boolean => isScannedText(entry.text)

/** Render the page bitmap for recognition (display orientation, ~2k px long edge) */
export async function renderPageForOcr(
  doc: PDFDocumentProxy,
  origIdx: number,
  geom: PageGeom,
): Promise<string | null> {
  const disp = geomDispSize(geom)
  const scale = Math.min(4, Math.max(1.5, 2048 / Math.max(disp.width, disp.height, 1)))
  try {
    const page = await doc.getPage(origIdx + 1)
    const viewport = page.getViewport({ scale, rotation: ((geom.rot % 360) + 360) % 360 })
    const canvas = document.createElement('canvas')
    canvas.width = Math.floor(viewport.width)
    canvas.height = Math.floor(viewport.height)
    await page.render({ canvas, viewport }).promise
    const png = canvas.toDataURL('image/png').split(',')[1]
    return png && png.length > 0 ? png : null
  } catch {
    return null
  }
}

/** Transparent selectable overlay; spans are re-projected through the live geometry,
    so zoom and unsaved rotations need no re-recognition */
export function OcrTextLayer({
  data,
  geom,
  scale,
}: {
  data: OcrPageData
  geom: PageGeom
  scale: number
}): ReactElement {
  return (
    <div className="pdf-ocr-layer">
      {data.words.map((w, i) => {
        const box = pdfRectToCss(geom, w.rect, scale)
        const fontPx = Math.max(box.height, 1)
        const measured = measurePt(w.text, fontPx, 'sans-serif')
        return (
          <span
            key={i}
            style={{
              left: box.left,
              top: box.top,
              fontSize: fontPx,
              transform: `scaleX(${measured > 0 ? box.width / measured : 1})`,
            }}
          >
            {w.text}
            {w.spaceAfter ? ' ' : ''}
          </span>
        )
      })}
    </div>
  )
}
