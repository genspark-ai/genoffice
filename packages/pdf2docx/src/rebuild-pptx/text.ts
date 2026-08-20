/**
 * TextBlock → pptx paragraph mapping (P25). One IR text block becomes ONE
 * text box with ONE paragraph: lines re-join with the docx rebuild's exact
 * semantics (hyphen joins, hard breaks as <a:br/>, no space between CJK
 * lines), and the measured line pitch rides as exact line spacing so the
 * box's vertical rhythm matches the source page.
 */
import { rectHeight } from '../geometry'
import type { Span, TextBlock } from '../ir'
import { isNoSpaceScript } from '../script'
import type { Paragraph, TextRun } from '../../../pptx-engine/src/types'

/** end-of-line hyphenation hyphens (plain, soft, unicode hyphen) */
const TRAILING_HYPHEN = /[-­‐]$/

function runFromSpan(span: Span, scale: number): TextRun | null {
  // invisible source text (PDF Tr 3/7): pptx has no w:vanish equivalent, and
  // painting it would ink hidden watermarks onto the slide — dropped
  if (span.invisible) return null
  const run: TextRun = { text: span.text }
  if (span.bold) run.bold = true
  if (span.italic) run.italic = true
  if (span.underline) run.underline = true
  if (span.strike) run.strike = true
  if (span.fontSize > 0) run.fontSize = span.fontSize * scale
  if (span.fontFamily) run.fontFamily = span.fontFamily
  run.color = `#${span.color || '000000'}`
  if (span.highlight) run.highlight = `#${span.highlight}`
  if (span.charSpacingPt !== undefined && span.charSpacingPt !== 0) {
    run.letterSpacing = span.charSpacingPt * scale
  }
  return run
}

const sameRunStyle = (a: TextRun, b: TextRun): boolean =>
  a.bold === b.bold &&
  a.italic === b.italic &&
  a.underline === b.underline &&
  a.strike === b.strike &&
  a.fontSize === b.fontSize &&
  a.fontFamily === b.fontFamily &&
  a.color === b.color &&
  a.highlight === b.highlight &&
  a.letterSpacing === b.letterSpacing

function pushRun(runs: TextRun[], run: TextRun): void {
  const last = runs[runs.length - 1]
  if (last && sameRunStyle(last, run)) last.text += run.text
  else runs.push(run)
}

/** flatten a block's lines into merged runs ('\n' in text = intra-box break) */
export function blockRuns(block: TextBlock, scale = 1): TextRun[] {
  const runs: TextRun[] = []
  for (const [i, line] of block.lines.entries()) {
    if (i > 0) {
      const prevLine = block.lines[i - 1]!
      const last = runs[runs.length - 1]
      const lastSpan = prevLine.spans[prevLine.spans.length - 1]
      const nextSpan = line.spans[0]
      if (line.hardBreakBefore) {
        // intentional intra-block break: generateRunXml renders '\n' as <a:br/>
        if (last) last.text += '\n'
        else runs.push({ text: '\n' })
      } else if (prevLine.endsWithHyphen && last) {
        last.text = last.text.replace(TRAILING_HYPHEN, '')
      } else if (
        last &&
        lastSpan &&
        nextSpan &&
        !isNoSpaceScript(lastSpan.script) &&
        !isNoSpaceScript(nextSpan.script)
      ) {
        last.text += ' '
      }
    }
    for (const span of line.spans) {
      const run = runFromSpan(span, scale)
      if (run) pushRun(runs, run)
    }
  }
  // footnote anchors and fully-invisible spans leave empty shells behind
  return runs.filter((r) => r.text !== '')
}

/**
 * The block's single paragraph. `scale` maps a differently-sized page onto
 * the deck's slide size (font sizes and pitch scale with the geometry).
 */
export function textBlockParagraph(block: TextBlock, scale = 1): Paragraph {
  const runs = blockRuns(block, scale)
  // literal list markers (P20 canvas lesson): a pptx auto-bullet would indent
  // and renumber; the marker is already measured into the box, so it rides as
  // plain text
  if (block.list?.marker !== undefined && runs.length > 0) {
    runs.unshift({ ...runs[0]!, text: `${block.list.marker.trimEnd()} ` })
  }
  // TOC dot-leader lines carry the page number separately; a literal tab
  // approximates the leader gap (default tab stops, dots dropped)
  if (block.tocEntry && runs.length > 0) {
    runs.push({ ...runs[runs.length - 1]!, text: `\t${block.tocEntry.pageNumber}` })
  }
  const p: Paragraph = { runs }
  if (block.align !== 'left') p.align = block.align
  if (block.dir === 'rtl') {
    p.rtl = true
    if (!p.align) p.align = 'right'
  }
  if (block.lines.length > 0) {
    const pitch = (rectHeight(block.box) / block.lines.length) * scale
    if (pitch > 0) p.lineExact = pitch
  }
  return p
}
