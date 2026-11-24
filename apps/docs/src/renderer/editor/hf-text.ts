/**
 * Plain-text view of a header/footer part, shared by the on-canvas editing
 * surface and the AI set_header_footer tool: paragraphs edit as lines, the
 * invisible PAGE / NUMPAGES / SECTIONPAGES field sentinels as visible {PAGE} /
 * {NUMPAGES} / {SECTIONPAGES} tokens. Layout-table rows (cells) and images stay out of the text flow and
 * keep their original content.
 */
import {
  BUILT_IN_HF_TAB_STOPS,
  PAGE_MARK,
  SECTION_PAGES_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfImage,
  type HfParagraph,
  type Run,
  type StyleInfo,
} from '@genoffice/docx-engine'

export const PAGE_TOKEN = '{PAGE}'
export const TOTAL_TOKEN = '{NUMPAGES}'
export const SECTION_TOKEN = '{SECTIONPAGES}'

/** effective paragraphs: rich paras when present, else the legacy single line */
export function hfParasOf(value: HeaderFooter, images?: HfImage[] | null): HfParagraph[] {
  if (value.paras?.length) return value.paras
  const runs: Run[] = value.text ? [{ text: value.text }] : []
  if (value.pageNumber && !value.text.includes('#') && !value.text.includes(PAGE_MARK)) {
    runs.push({ text: runs.length > 0 ? ` ${PAGE_MARK}` : PAGE_MARK })
  }
  // Word reserves exactly the picture height for a picture-only part: no text line
  if (runs.length === 0 && images?.some((im) => !im.floating)) return []
  return [{ align: 'center', runs }]
}

/**
 * The line a new header or footer starts with in Word: one paragraph in the
 * document's Header / Footer style, left-aligned with a center and a right tab
 * stop, or in Word's built-in definition when the document has none (the save
 * adds it). Its properties are resolved here like a parsed strip paragraph's.
 */
export function hfBlankPara(
  kind: 'header' | 'footer',
  styles?: ReadonlyMap<string, StyleInfo>,
): HfParagraph {
  const id = kind === 'header' ? 'Header' : 'Footer'
  const own =
    styles?.get(id) ??
    [...(styles?.values() ?? [])].find((s) => s.type === 'paragraph' && s.name === kind)
  if (!own)
    return {
      styleId: id,
      tabStops: [...BUILT_IN_HF_TAB_STOPS],
      spaceAfter: 0,
      lineSpacing: 1,
      lineRule: 'auto',
      lineRawTwips: 240,
      runs: [],
    }
  const d = own.display ?? {}
  return {
    styleId: own.styleId,
    ...(d.align ? { align: d.align } : {}),
    ...(d.tabStops ? { tabStops: d.tabStops } : {}),
    ...(d.spaceAfterTwips !== undefined ? { spaceAfter: d.spaceAfterTwips } : {}),
    ...(d.lineSpacing !== undefined ? { lineSpacing: d.lineSpacing } : {}),
    ...(d.lineRule ? { lineRule: d.lineRule } : {}),
    ...(d.lineRawTwips !== undefined ? { lineRawTwips: d.lineRawTwips } : {}),
    runs: [],
  }
}

/** a part with nothing in it yet: no paragraph, text or page number */
export function hfIsBlank(value: HeaderFooter): boolean {
  return !value.paras?.length && !value.text && !value.pageNumber
}

/** Pictures are not editable: they leave the text runs and return to the
 *  same side of the text they came from (a logo before the title stays first). */
export function hfPictureRuns(runs: Run[]): { leading: Run[]; text: Run[]; trailing: Run[] } {
  const leading: Run[] = []
  const text: Run[] = []
  const trailing: Run[] = []
  for (const run of runs) {
    if (run.image) {
      ;(text.length > 0 ? trailing : leading).push({ ...run, text: '' })
      if (run.text === '') continue
    }
    const { image: _image, ...rest } = run
    text.push(rest)
  }
  return { leading, text, trailing }
}

/** editable text of the part: one line per text paragraph, field sentinels as tokens */
export function hfEditText(value: HeaderFooter): string {
  return hfParasOf(value)
    .filter((p) => !p.cells)
    .map((p) =>
      p.runs
        .filter((r) => !r.del)
        .map((r) => r.text)
        .join(''),
    )
    .join('\n')
    .replaceAll(PAGE_MARK, PAGE_TOKEN)
    .replaceAll(TOTAL_PAGES_MARK, TOTAL_TOKEN)
    .replaceAll(SECTION_PAGES_MARK, SECTION_TOKEN)
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
    .replaceAll(SECTION_TOKEN, SECTION_PAGES_MARK)
    .split('\n')
  const textParas = paras.filter((p) => !p.cells)
  const templates: HfParagraph[] =
    textParas.length > 0 ? textParas : [{ align: 'center', runs: [] }]
  const edited: HfParagraph[] = lines.map((line, i) => {
    const template = templates[Math.min(i, templates.length - 1)]
    const { leading, text: textRuns, trailing } = hfPictureRuns(template.runs)
    // the rewritten line is the accepted text: it must not inherit a tracked wrapper
    const { ins: _ins, del: _del, ...style } = textRuns.find((r) => !r.del) ?? textRuns[0] ?? {}
    return {
      ...template,
      runs: [...leading, ...(line === '' ? [] : [{ ...style, text: line }]), ...trailing],
    }
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
