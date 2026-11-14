// Modified for ForaOffice: retain original header/footer text containers on replacement.
/**
 * Plain-text view of a header/footer part, shared by the on-canvas editing
 * surface and the AI set_header_footer tool: paragraphs edit as lines, the
 * invisible PAGE / NUMPAGES / SECTIONPAGES field sentinels as visible {PAGE} /
 * {NUMPAGES} / {SECTIONPAGES} tokens. Layout-table rows (cells) and images stay out of the text flow and
 * keep their original content.
 */
import {
  PAGE_MARK,
  SECTION_PAGES_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfImage,
  type HfParagraph,
  type Run,
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

/** Textbox coordinates describe the currently parsed part, not persistent IDs. */
export function hfTextboxes(value: HeaderFooter | null) {
  const groups = new Map<number, HfParagraph[]>()
  for (const para of value?.paras ?? []) {
    if (!para.textbox) continue
    const group = groups.get(para.textbox.index) ?? []
    group.push(para)
    groups.set(para.textbox.index, group)
  }
  return [...groups].map(([index, paras]) => ({
    index,
    text: hfEditText({ text: '', paras }),
    editable: !paras.some((para) => para.textbox?.readOnly),
    lifetime: 'current-part' as const,
  }))
}

/**
 * Map edited lines back onto the part: each line keeps its original
 * paragraph's format and first-run styling (extra lines reuse the last
 * template), cells rows are spliced back at their original positions.
 */
export function applyHfText(
  value: HeaderFooter | null,
  text: string,
  textboxIndex?: number,
): HeaderFooter {
  const base = value ?? { text: '' }
  const paras = hfParasOf(base)
  const lines = text
    .replace(/\n+$/, '')
    .replaceAll(PAGE_TOKEN, PAGE_MARK)
    .replaceAll(TOTAL_TOKEN, TOTAL_PAGES_MARK)
    .replaceAll(SECTION_TOKEN, SECTION_PAGES_MARK)
    .split('\n')
  const editable = (p: HfParagraph) =>
    !p.cells && !p.lineOnly && (textboxIndex === undefined || p.textbox?.index === textboxIndex)
  const textParas = paras.filter(editable)
  const templates: HfParagraph[] =
    textParas.length > 0 ? textParas : [{ align: 'center', runs: [] }]
  const edited: HfParagraph[] = lines.map((line, i) => {
    const template = templates[Math.min(i, templates.length - 1)]
    const { leading, text: textRuns, trailing } = hfPictureRuns(template.runs)
    // the rewritten line is the accepted text: it must not inherit a tracked wrapper
    const { ins: _ins, del: _del, ...style } = textRuns.find((r) => !r.del) ?? textRuns[0] ?? {}
    return {
      ...template,
      ...(template.textbox
        ? {
            textbox: {
              ...template.textbox,
              paragraph: template.textbox.paragraph + Math.max(0, i - templates.length + 1),
              rewrite: true,
            },
          }
        : {}),
      runs: [...leading, ...(line === '' ? [] : [{ ...style, text: line }]), ...trailing],
    }
  })
  const nextParas: HfParagraph[] = []
  let ei = 0
  for (const p of paras) {
    if (!editable(p)) nextParas.push(p)
    else if (ei < edited.length) nextParas.push(edited[ei++])
    // An omitted textbox paragraph must be cleared in its original container,
    // not dropped from the model while its old XML remains protected.
    else if (p.textbox) nextParas.push({ ...p, runs: [], textbox: { ...p.textbox, rewrite: true } })
  }
  nextParas.push(...edited.slice(ei))
  const nextText = nextParas
    .filter((p) => !p.cells)
    .map((p) => p.runs.map((r) => r.text).join(''))
    .join('')
  const pageNumber = nextParas.some((para) =>
    [...para.runs, ...(para.cells?.flatMap((cell) => cell.paras.flat()) ?? [])].some((run) =>
      run.text.includes(PAGE_MARK),
    ),
  )
  return { ...base, text: nextText, paras: nextParas, pageNumber }
}
