import type { TextInsertInput } from '../shared/ipc'
import type { OcrWord } from './ocr-layer'
import type { PageGeom } from './annotations'

/**
 * The persistence half of auto-OCR: the recognized words already carry their
 * PDF user-space boxes, so baking is grouping them into visual lines and
 * emitting one white text insert per line — invisible over the scan, but
 * selectable and searchable in every reader, and it survives every save like
 * any other content. After the save the reload finds a text layer, so the
 * auto-OCR pass never fires for these pages again: the bake is its own
 * dedupe.
 */

const BAKE_COLOR: [number, number, number] = [1, 1, 1]

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

/** one white insert per line: origin at the line's left baseline, size from its box height */
export function buildOcrBakeInserts(
  origIdx: number,
  words: OcrWord[],
  geom: PageGeom,
): TextInsertInput[] {
  const rotate = ((geom.rot % 360) + 360) % 360
  return groupLines(words).map((line) => {
    const left = Math.min(...line.map((w) => w.rect[0]))
    const bottom = Math.min(...line.map((w) => w.rect[1]))
    const height = Math.max(...line.map((w) => w.rect[3] - w.rect[1]))
    let text = ''
    line.forEach((word, i) => {
      // spaceAfter is the overlay's own convention: a real space follows this word
      if (i > 0 && line[i - 1]!.spaceAfter) text += ' '
      text += word.text
    })
    return {
      pageIndex: origIdx,
      origin: [left, bottom + height * 0.22],
      text,
      fontSize: Math.max(4, Math.min(height * 0.82, 72)),
      color: BAKE_COLOR,
      lineLeading: height * 1.1,
      rotate,
    }
  })
}
