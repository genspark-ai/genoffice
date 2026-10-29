import type { TextInsertInput } from '../shared/ipc'
import type { OcrWord } from './ocr-layer'
import { pdfToView, viewToPdf, type PageGeom } from './annotations'

/**
 * The persistence half of auto-OCR: the recognized words already carry their
 * PDF user-space boxes, so baking is grouping them into visual lines and
 * emitting one invisible (render mode 3) text insert per line — present,
 * selectable and searchable in every reader but never painted, and it survives
 * every save like any other content. After the save the reload finds a text
 * layer, so the auto-OCR pass never fires for these pages again: the bake is
 * its own dedupe.
 */

/** text-insert colours are 0-255 RGB (see the _FPDFPageObj_SetFillColor call in
 *  text-edit.ts) — [1, 1, 1] here was one-almost-black on every scan. The fill
 *  is belt-and-braces anyway: the insert rides with render mode 3. */
const BAKE_COLOR: [number, number, number] = [255, 255, 255]

/** a line groups words whose vertical centers sit within half a line height */
export function groupLines(words: OcrWord[]): OcrWord[][] {
  const sorted = [...words].sort((a, b) => {
    const ca = (a.rect[1] + a.rect[3]) / 2
    const cb = (b.rect[1] + b.rect[3]) / 2
    if (Math.abs(ca - cb) > Math.min(a.rect[3] - a.rect[1], b.rect[3] - b.rect[1]) / 2) {
      return cb - ca
    }
    return a.rect[0] - b.rect[0]
  })
  const lines: OcrWord[][] = []
  for (const word of sorted) {
    const line = lines[lines.length - 1]
    if (line && sameLine(line, word)) line.push(word)
    else lines.push([word])
  }
  return lines
}

function sameLine(line: OcrWord[], word: OcrWord): boolean {
  const last = line[line.length - 1]!
  const lh = Math.min(last.rect[3] - last.rect[1], word.rect[3] - word.rect[1])
  const ca = (last.rect[1] + last.rect[3]) / 2
  const cb = (word.rect[1] + word.rect[3]) / 2
  return Math.abs(ca - cb) <= lh / 2
}

/** a user-space rect's axis-aligned box in display coordinates (y down) */
function displayBox(
  geom: PageGeom,
  rect: readonly [number, number, number, number],
): { x0: number; y0: number; x1: number; y1: number } {
  const xs: number[] = []
  const ys: number[] = []
  for (const [x, y] of [
    [rect[0], rect[1]],
    [rect[2], rect[1]],
    [rect[2], rect[3]],
    [rect[0], rect[3]],
  ] as const) {
    const [vx, vy] = pdfToView(geom, x, y)
    xs.push(vx)
    ys.push(vy)
  }
  return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) }
}

interface DispWord {
  word: OcrWord
  box: { x0: number; y0: number; x1: number; y1: number }
}

function dispSameLine(line: DispWord[], word: DispWord): boolean {
  const last = line[line.length - 1]!
  const lh = Math.min(last.box.y1 - last.box.y0, word.box.y1 - word.box.y0)
  const ca = (last.box.y0 + last.box.y1) / 2
  const cb = (word.box.y0 + word.box.y1) / 2
  return Math.abs(ca - cb) <= lh / 2
}

/** the display-space twin of groupLines: same rule, y grows downward there */
function groupDisplayLines(words: DispWord[]): DispWord[][] {
  const sorted = [...words].sort((a, b) => {
    const ca = (a.box.y0 + a.box.y1) / 2
    const cb = (b.box.y0 + b.box.y1) / 2
    if (Math.abs(ca - cb) > Math.min(a.box.y1 - a.box.y0, b.box.y1 - b.box.y0) / 2) {
      return ca - cb
    }
    return a.box.x0 - b.box.x0
  })
  const lines: DispWord[][] = []
  for (const word of sorted) {
    const line = lines[lines.length - 1]
    if (line && dispSameLine(line, word)) line.push(word)
    else lines.push([word])
  }
  return lines
}

/** one invisible insert per line: origin at the line's left baseline, size from its box height */
export function buildOcrBakeInserts(
  origIdx: number,
  words: OcrWord[],
  geom: PageGeom,
): TextInsertInput[] {
  const rotate = ((geom.rot % 360) + 360) % 360
  // Everything — grouping, the line box, the baseline — happens in DISPLAY
  // space (the page as the reader sees it, after its /Rotate), because the
  // reading direction lives there. Working in user space agrees only at 0
  // degrees: for 90/270 the writer's counter-rotation sent the run down the
  // wrong axis from the wrong corner. The baseline maps back through
  // viewToPdf, the one transform every other layer uses.
  const disp = words.map((word) => ({ word, box: displayBox(geom, word.rect) }))
  return groupDisplayLines(disp).map((line) => {
    const left = Math.min(...line.map((w) => w.box.x0))
    const bottom = Math.max(...line.map((w) => w.box.y1))
    const height = Math.max(...line.map((w) => w.box.y1 - w.box.y0))
    let text = ''
    line.forEach(({ word }, i) => {
      // spaceAfter is the overlay's own convention: a real space follows this word
      if (i > 0 && line[i - 1]!.word.spaceAfter) text += ' '
      text += word.text
    })
    const [ux, uy] = viewToPdf(geom, left, bottom - height * 0.22)
    return {
      pageIndex: origIdx,
      origin: [ux, uy],
      text,
      fontSize: Math.max(4, Math.min(height * 0.82, 72)),
      color: BAKE_COLOR,
      lineLeading: height * 1.1,
      rotate,
      invisible: true,
    }
  })
}
