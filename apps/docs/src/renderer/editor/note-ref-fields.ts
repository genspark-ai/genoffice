/** NOTEREF fields: a cross-reference to a footnote or endnote through a bookmark on its mark. */
import type { Mark, Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { noteMarkText, type NoteKind } from '../note-format'
import { isFieldLocked } from './field-codes'

const NOTEREF_INSTR = /^\s*NOTEREF\s+(?:"([^"]+)"|([^\s\\]+))/i

/** the bookmark a NOTEREF instruction names; null for any other field */
export const noteRefBookmark = (instr: string): string | null => {
  const m = NOTEREF_INSTR.exec(instr)
  return m ? (m[1] ?? m[2]) : null
}

const names = (node: PmNode, attr: 'bookmarks' | 'hiddenBookmarks'): string[] => {
  const list = node.attrs?.[attr] as string[] | null | undefined
  return Array.isArray(list) ? list : []
}

/** the bookmark a cross-reference inserted as a hyperlink (\\h) goes to: a NOTEREF or a PAGEREF field */
export const fieldLinkBookmark = (instr: string): string | null =>
  /\\h\b/i.test(instr)
    ? (/^\s*(?:NOTEREF|PAGEREF)\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(instr)?.slice(1).find(Boolean) ??
      null)
    : null

/** what a reference mark shows: its custom mark, or its number in the format its section gives it */
export const noteMarkOf = (node: PmNode): string =>
  typeof node.attrs.mark === 'string'
    ? node.attrs.mark
    : noteMarkText(
        node.attrs.kind as NoteKind,
        Number(node.attrs.num) || 1,
        node.attrs.fmt as string | null,
      )

/** the reference mark the bookmark is on, or the first mark of the paragraph it is on */
export function noteMarkAtBookmark(
  doc: PmNode,
  name: string,
): { pos: number; node: PmNode } | null {
  let hit: { pos: number; node: PmNode } | null = null
  let host: { from: number; to: number } | null = null
  doc.descendants((node, pos) => {
    if (hit) return false
    const here =
      names(node, 'hiddenBookmarks').includes(name) || names(node, 'bookmarks').includes(name)
    if (node.type.name === 'docNoteRef') {
      if (here || (host && pos > host.from && pos < host.to)) hit = { pos, node }
      return false
    }
    if (here && !host) host = { from: pos, to: pos + node.nodeSize }
    return true
  })
  return hit
}

export interface NoteRefWords {
  above: string
  below: string
}

/** NOTEREF \\p as Word computes it: the mark, a space, then the position word */
export const relativeNoteRef = (mark: string, where: string): string => `${mark} ${where}`

/**
 * The result a NOTEREF field at `at` shows now; null when it is not one, or its note is gone.
 * Without `words` an above / below field keeps the word it shows (`shown`) and takes the number.
 */
export function noteRefFieldText(
  doc: PmNode,
  instr: string,
  at: number,
  words?: NoteRefWords,
  shown?: string,
): string | null {
  const name = noteRefBookmark(instr)
  const target = name === null ? null : noteMarkAtBookmark(doc, name)
  if (!target) return null
  const mark = noteMarkOf(target.node)
  if (!/\\p\b/i.test(instr)) return mark
  if (words) return relativeNoteRef(mark, target.pos < at ? words.above : words.below)
  const word = shown?.slice(shown.indexOf(' ') + 1).trim()
  return word ? relativeNoteRef(mark, word) : null
}

/** brings every NOTEREF result in `tr` up to date; without `words` an above / below one keeps its word. Returns how many changed. */
export function refreshNoteRefFields(tr: Transaction, words?: NoteRefWords): number {
  const jobs: Array<{
    from: number
    to: number
    instr: string
    marks: readonly Mark[]
    was: string
  }> = []
  let open: { mark: Mark; job: (typeof jobs)[number] } | null = null
  tr.doc.descendants((node, pos) => {
    if (!node.isText) return true
    const mark = node.marks.find((m) => m.type.name === 'instrField')
    if (!mark || isFieldLocked(mark) || noteRefBookmark(String(mark.attrs.instr)) === null) {
      open = null
      return false
    }
    // one field's result can be several text nodes: it is rewritten as one
    if (open && open.job.to === pos && open.mark.eq(mark)) {
      open.job.to = pos + node.nodeSize
      open.job.was += node.text ?? ''
      return false
    }
    const job = {
      from: pos,
      to: pos + node.nodeSize,
      instr: String(mark.attrs.instr),
      marks: node.marks,
      was: node.text ?? '',
    }
    jobs.push(job)
    open = { mark, job }
    return false
  })
  const changed = jobs.flatMap((j) => {
    const text = noteRefFieldText(tr.doc, j.instr, j.from, words, j.was)
    return text === null || text === j.was ? [] : [{ ...j, text }]
  })
  for (const j of changed.reverse())
    tr.replaceWith(j.from, j.to, tr.doc.type.schema.text(j.text, [...j.marks]))
  return changed.length
}
