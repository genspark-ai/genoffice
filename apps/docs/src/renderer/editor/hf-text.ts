/**
 * Plain-text view of a header/footer part, shared by the on-canvas editing
 * surface and the AI set_header_footer tool: paragraphs edit as lines, the
 * invisible PAGE / NUMPAGES field sentinels as visible {PAGE} / {NUMPAGES}
 * tokens. Layout-table rows (cells) and images stay out of the text flow and
 * keep their original content.
 */
import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfParagraph,
  type Run,
} from '@genoffice/docx-engine'

export const PAGE_TOKEN = '{PAGE}'
export const TOTAL_TOKEN = '{NUMPAGES}'

/** effective paragraphs: rich paras when present, else the legacy single line */
export function hfParasOf(value: HeaderFooter): HfParagraph[] {
  if (value.paras?.length) return value.paras
  const runs: Run[] = value.text ? [{ text: value.text }] : []
  if (value.pageNumber && !value.text.includes('#') && !value.text.includes(PAGE_MARK)) {
    runs.push({ text: runs.length > 0 ? ` ${PAGE_MARK}` : PAGE_MARK })
  }
  return [{ align: 'center', runs }]
}

/** editable text of the part: one line per text paragraph, field sentinels as tokens */
export function hfEditText(value: HeaderFooter): string {
  return hfParasOf(value)
    .filter((p) => !p.cells)
    .map((p) => p.runs.map((r) => r.text).join(''))
    .join('\n')
    .replaceAll(PAGE_MARK, PAGE_TOKEN)
    .replaceAll(TOTAL_PAGES_MARK, TOTAL_TOKEN)
}

/**
 * Map edited lines back onto the part: each line keeps its original
 * paragraph's format and first-run styling (extra lines reuse the last
 * template), cells rows are spliced back at their original positions.
 */
export function applyHfText(value: HeaderFooter | null, text: string): HeaderFooter {
  const base = value ?? { text: '' }
  const paras = hfParasOf(base)
  const lines = text
    .replace(/\n+$/, '')
    .replaceAll(PAGE_TOKEN, PAGE_MARK)
    .replaceAll(TOTAL_TOKEN, TOTAL_PAGES_MARK)
    .split('\n')
  const textParas = paras.filter((p) => !p.cells)
  const templates: HfParagraph[] =
    textParas.length > 0 ? textParas : [{ align: 'center', runs: [] }]
  const edited: HfParagraph[] = lines.map((line, i) => {
    const template = templates[Math.min(i, templates.length - 1)]
    const style = template.runs[0] ?? {}
    return { ...template, runs: line === '' ? [] : [{ ...style, text: line }] }
  })
  const nextParas: HfParagraph[] = []
  let ei = 0
  for (const p of paras) {
    if (p.cells) nextParas.push(p)
    else if (ei < edited.length) nextParas.push(edited[ei++])
  }
  nextParas.push(...edited.slice(ei))
  const nextText = edited.map((p) => p.runs.map((r) => r.text).join('')).join('')
  return { ...base, text: nextText, paras: nextParas }
}
