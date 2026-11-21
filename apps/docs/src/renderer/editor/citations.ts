import type { Node as PmNode } from '@tiptap/pm/model'
import type { FieldTextRun } from '@genoffice/docx-engine'

/** Word's table of authorities categories, by the number a TA field's \c switch carries */
export const CITATION_CATEGORIES = [1, 2, 3, 4, 5, 6, 7] as const

export interface Citation {
  /** the citation as the table lists it (\l) */
  long: string
  /** what later mentions of the same authority read (\s); they carry it alone */
  short?: string
  category: number
  /** the long citation in formatted pieces, when the cited text was not plain */
  runs?: FieldTextRun[]
}

/** the formatted pieces a citation marker carries (its `taRuns` attribute) */
export function citationRuns(attr: unknown): FieldTextRun[] | undefined {
  if (typeof attr !== 'string' || !attr) return undefined
  try {
    const runs = JSON.parse(attr) as FieldTextRun[]
    return Array.isArray(runs) && runs.every((r) => typeof r?.text === 'string') ? runs : undefined
  } catch {
    return undefined
  }
}

/**
 * The text between two positions in pieces of like bold / italic / underline,
 * trimmed as the Mark Citation dialog trims the selection. Undefined when it
 * is plain, or is not the long citation any more (the dialog's text was edited).
 */
export function citationRunsOf(
  doc: PmNode,
  from: number,
  to: number,
  long: string,
): FieldTextRun[] | undefined {
  const out: FieldTextRun[] = []
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return
    // the field's text cannot hold a quote: the pieces lose theirs as the long citation does
    const text = (node.text ?? '').slice(Math.max(0, from - pos), to - pos).replace(/"/g, '')
    if (!text) return
    const has = (name: string) => node.marks.some((m) => m.type.name === name)
    const piece: FieldTextRun = {
      text,
      ...(has('bold') ? { bold: true } : {}),
      ...(has('italic') ? { italic: true } : {}),
      ...(has('underline') ? { underline: true } : {}),
    }
    const last = out[out.length - 1]
    if (
      last &&
      !!last.bold === !!piece.bold &&
      !!last.italic === !!piece.italic &&
      !!last.underline === !!piece.underline
    )
      last.text += text
    else out.push(piece)
  })
  if (out.length > 0) {
    out[0].text = out[0].text.trimStart()
    out[out.length - 1].text = out[out.length - 1].text.trimEnd()
  }
  const runs = out.filter((r) => r.text !== '')
  if (runs.map((r) => r.text).join('') !== long.replace(/"/g, '').trim()) return undefined
  return runs.some((r) => r.bold || r.italic || r.underline) ? runs : undefined
}

/**
 * What a TA field's switches say. A later mention carries the short citation
 * alone: its long text and category are those of the mark that defines it.
 */
export function citationOf(ta: string): Citation {
  const short = /\\s\s+"([^"]*)"/.exec(ta)?.[1] ?? ''
  const long = /\\l\s+"([^"]*)"/.exec(ta)?.[1] ?? ''
  const category = Number(/\\c\s+"?(\d+)"?/.exec(ta)?.[1])
  return { long, short, category: category > 0 ? category : 1 }
}

// a quote inside a text would end it early
const literal = (text = '') => text.replace(/"/g, '').trim()

/** the switches of the TA field that defines a citation */
export const citationSwitches = ({ long, short, category }: Citation): string =>
  `\\l "${literal(long)}" ${literal(short) ? `\\s "${literal(short)}" ` : ''}\\c ${Math.round(category)}`

/** the switches of a TA field at a later mention of the citation */
export const shortCitationSwitches = (short = ''): string => `\\s "${literal(short)}"`

/** each mark as the table lists it: a later mention takes the long text and category of its short citation */
export function resolveCitations<T extends Citation>(marks: T[]): T[] {
  const defined = new Map<string, Citation>()
  for (const mark of marks)
    if (mark.long && mark.short && !defined.has(mark.short)) defined.set(mark.short, mark)
  return marks
    .map((mark) => {
      if (mark.long) return mark
      const def = defined.get(mark.short ?? '')
      return {
        ...mark,
        long: def?.long ?? mark.short ?? '',
        category: def?.category ?? mark.category,
      }
    })
    .filter((mark) => mark.long !== '')
}
