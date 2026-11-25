import { captionNumberRunsXml, type SeqNumberFormat } from './generate'
import { parseSeqInstruction, parseStyleRefInstruction, seqFieldsOfXml } from './seq-fields'
import { CHAPTER_SEPARATORS, type PageNumberChapter } from './types'

/** what Word's Caption Numbering dialog sets for a label */
export interface CaptionNumbering {
  format: SeqNumberFormat
  chapter: PageNumberChapter | null
}

const RPR = '(?:<w:rPr>(?:(?!</w:rPr>)[\\s\\S])*</w:rPr>|<w:rPr/>)?'
const RUN_OPEN = new RegExp(`<w:r(?:\\s[^>]*)?>(${RPR})$`)
const PROOF = '(?:<w:proofErr\\b[^>]*/>)*'
const SEPARATOR_RUN = new RegExp(
  `^${PROOF}<w:r(?:\\s[^>]*)?>${RPR}` +
    `(?:<w:(noBreakHyphen)/>|<w:t(?:\\s[^>]*)?>([^<]{1,2})</w:t>)</w:r>${PROOF}$`,
)

interface FieldRuns {
  from: number
  to: number
  rPr: string
}

/** the field with the runs its markers stand in; null when a marker shares its run with something else */
function fieldRuns(xml: string, range: [number, number]): FieldRuns | null {
  if (xml.startsWith('<w:fldSimple', range[0])) {
    const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(xml.slice(range[0], range[1]))?.[0] ?? ''
    return { from: range[0], to: range[1], rPr }
  }
  const open = RUN_OPEN.exec(xml.slice(0, range[0]))
  if (!open || !xml.startsWith('</w:r>', range[1])) return null
  return { from: open.index, to: range[1] + '</w:r>'.length, rPr: open[1] }
}

const separatorName = (text: string): string | null =>
  Object.keys(CHAPTER_SEPARATORS).find((k) => CHAPTER_SEPARATORS[k] === text) ?? null

interface CaptionNumber {
  label: string
  numbering: CaptionNumbering
  /** the SEQ field and, with a chapter number, the STYLEREF and separator before it */
  from: number
  to: number
  rPr: string
  result: string
  chapterResult: string | undefined
  dirty: boolean
}

function captionNumberOf(xml: string, label?: string): CaptionNumber | null {
  const fields = seqFieldsOfXml(xml)
  const at = fields.findIndex((f) => {
    const seq = f.kind === 'SEQ' ? parseSeqInstruction(f.instr) : null
    return !!seq && (label === undefined || seq.id === label)
  })
  const field = fields[at]
  const seq = field ? parseSeqInstruction(field.instr) : null
  // a number that restarts, repeats or hides is somebody's own arrangement
  if (!field || !seq || seq.unsupported || seq.restart !== null || seq.repeat || seq.hidden)
    return null
  if (seq.format === null) return null
  const runs = fieldRuns(xml, field.range)
  if (!runs) return null
  const out: CaptionNumber = {
    label: seq.id,
    numbering: { format: seq.format, chapter: null },
    from: runs.from,
    to: runs.to,
    rPr: runs.rPr,
    result: field.result,
    chapterResult: undefined,
    dirty: field.dirtyAttr !== null,
  }
  const lead = at > 0 ? fields[at - 1] : null
  const ref = lead?.kind === 'STYLEREF' ? parseStyleRefInstruction(lead.instr) : null
  if (!lead || !ref || ref.level === null || ref.show !== 'bare') return out
  const leadRuns = fieldRuns(xml, lead.range)
  const sep = leadRuns ? SEPARATOR_RUN.exec(xml.slice(leadRuns.to, runs.from)) : null
  const name = sep ? separatorName(sep[1] ? '-' : sep[2].replace('‑', '-')) : null
  if (!leadRuns || !name) return out
  out.from = leadRuns.from
  out.chapterResult = lead.result
  out.numbering.chapter = { style: ref.level, sep: name }
  return out
}

/** the numbering of the caption paragraph's first SEQ field (of `label`, when given) */
export function captionNumberingOfXml(
  xml: string,
  label?: string,
): { label: string; numbering: CaptionNumbering } | null {
  const found = captionNumberOf(xml, label)
  return found ? { label: found.label, numbering: found.numbering } : null
}

/**
 * The caption with its number written the way `numbering` asks, the results it shows
 * kept until the next renumbering. Null when it already is, or when its fields are not
 * laid out as Word's Caption dialog writes them.
 */
export function setCaptionNumberingXml(
  xml: string,
  label: string,
  numbering: CaptionNumbering,
): string | null {
  const found = captionNumberOf(xml, label)
  if (!found) return null
  const had = found.numbering
  if (
    had.format === numbering.format &&
    had.chapter?.style === numbering.chapter?.style &&
    had.chapter?.sep === numbering.chapter?.sep
  )
    return null
  const runs = captionNumberRunsXml({
    rPr: found.rPr,
    label,
    format: numbering.format,
    chapter: numbering.chapter,
    chapterResult: found.chapterResult,
    result: found.result,
    dirty: found.dirty,
  })
  return xml.slice(0, found.from) + runs + xml.slice(found.to)
}
