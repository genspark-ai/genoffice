import {
  conditionHolds,
  ifBranch,
  mergeFieldName,
  mergeFieldText,
  splitInstruction,
  type MergeRecord,
} from '@genoffice/docx-engine'
import type { PmNode } from '../editor/convert'

/**
 * The template half of the mail merge: the document with its merge fields
 * filled from a record.
 *
 * The editor keeps a field as a mark (instrField) on the text that shows its
 * result, so a field is found by its mark and not by a text pattern — the
 * same way the editor itself finds them. What a field is worth comes from the
 * engine (mergeFieldText), so the switches Word writes (`\b`, `\f`, `\* Upper`)
 * behave here exactly as they do in a merge.
 *
 * Nothing here knows about the UI: the panel asks for a preview and asks for
 * a merge, and this is what answers. genoffice#2097.
 */

/** the mark a field carries */
const FIELD_MARK = 'instrField'

/**
 * What a field is worth for one record, or null when the merge leaves it
 * alone.
 *
 * A MERGEFIELD takes the column's value; an IF takes the branch its condition
 * selects. Both come from the engine, so the switches Word writes behave here
 * as they do in a merge. A field that is neither (a PAGE, a DATE) is left as
 * it was: it still evaluates in Word.
 */
export function mergeFieldValue(instr: string, record: MergeRecord): string | null {
  if (mergeFieldName(instr)) return mergeFieldText(instr, record)
  const [keyword] = splitInstruction(instr)
  if (keyword === 'IF') return ifBranch(instr, record)
  return null
}

/**
 * The document with every merge field replaced by the record's value.
 *
 * A field's result may be split over several text nodes (Word does that
 * whenever the formatting changes inside it), so the nodes that carry the
 * same field are one field: the value goes into the first of them and the
 * others are emptied. Every other field (a PAGE, a DATE) is left as it was.
 */
export function substituteMergeFields(doc: PmNode, record: MergeRecord): PmNode {
  return { ...doc, content: substituteContent(doc.content, record) }
}

function substituteContent(
  content: readonly PmNode[] | undefined,
  record: MergeRecord,
): PmNode[] | undefined {
  if (!content) return undefined
  const out: PmNode[] = []
  // the field the previous text node carried, and how many nodes it spans
  let fieldInstr: string | null = null
  let fieldSeen = false
  for (const node of content) {
    const mark = node.marks?.find((m) => m.type === FIELD_MARK)
    const instr = mark ? String(mark.attrs?.instr ?? '') : ''
    const value = mark ? mergeFieldValue(instr, record) : null
    if (value === null) {
      fieldInstr = null
      fieldSeen = false
      out.push(substituteNode(node, record))
      continue
    }
    if (instr !== fieldInstr) {
      fieldInstr = instr
      fieldSeen = false
    }
    // the field's result may span several nodes: the value goes into the
    // first of them and the others are emptied
    const text = fieldSeen ? '' : value
    fieldSeen = true
    out.push({
      ...node,
      text,
      // the field is gone: what is left is the value, with the text's own marks
      marks: (node.marks ?? []).filter((m) => m.type !== FIELD_MARK),
      content: substituteContent(node.content, record),
    })
  }
  return out
}

function substituteNode(node: PmNode, record: MergeRecord): PmNode {
  if (!node.content) return node
  return { ...node, content: substituteContent(node.content, record) }
}

/**
 * The text of a document, one line per block.
 *
 * This is the preview: the template with one record's values in it, as the
 * user reads it.
 */
export function documentText(doc: PmNode): string {
  const lines: string[] = []
  for (const block of doc.content ?? []) {
    const text = blockText(block)
    if (text !== '' || lines.length > 0) lines.push(text)
  }
  return lines.join('\n')
}

function blockText(node: PmNode): string {
  if (node.type === 'text') return node.text ?? ''
  if (!node.content) return ''
  return node.content.map(blockText).join('')
}

/**
 * Every merge field the template holds, in document order, without repeats.
 *
 * The panel offers these as the fields it can insert, and the preview tells
 * the user which of them the current record fills.
 */
export function mergeFieldNames(doc: PmNode): string[] {
  const names: string[] = []
  const walk = (node: PmNode): void => {
    for (const mark of node.marks ?? []) {
      const name = mergeFieldName(String(mark.attrs?.instr ?? ''))
      if (name && !names.includes(name)) names.push(name)
    }
    for (const child of node.content ?? []) walk(child)
  }
  walk(doc)
  return names
}

/** a question an ASK or FILLIN field puts to the user during the merge */
export interface MergePrompt {
  /** names the field, and the record too when it asks for every one */
  key: string
  prompt: string
  /** the answer when none is typed (the field's default switch) */
  fallback: string
  /** the record's place in the list, counted from 1; null for a field that asks once */
  record: number | null
}

export interface MergeRunOptions {
  /** the place in the list of the first record merged, counted from 1 (MERGEREC counts from it) */
  firstRecord?: number
  /** what the user answered, by prompt key; a prompt with no answer takes its default */
  answers?: ReadonlyMap<string, string>
}

export interface MergedDocument {
  doc: PmNode
  /** the record the document opened with */
  record: MergeRecord
}

/** the fields that steer the merge and show nothing where they stand */
const RULE_ONLY = /^\s*(?:NEXT|NEXTIF|SKIPIF|SET|ASK)(?=\s|$)/i

interface MergeWalk {
  /** index of the record being read; NEXT and NEXTIF move it on */
  index: number
  skipped: boolean
  /** ASK and FILLIN fields met so far in this pass */
  asked: number
  /** what this pass has to ask; a record that is skipped asks nothing */
  prompts: MergePrompt[]
}

/** the quoted or bare arguments of an instruction, after its keyword */
const fieldArgs = (instr: string): string[] =>
  (splitInstruction(instr)[1].match(/"[^"]*"|\S+/g) ?? []).map((arg) =>
    arg.replace(/^"([^"]*)"$/, '$1'),
  )

/** the text after a switch such as the default of an ASK */
const switchArg = (args: string[], name: string): string => {
  const at = args.findIndex((arg) => arg.toLowerCase() === name)
  return at === -1 ? '' : (args[at + 1] ?? '')
}

/**
 * The documents a merge writes, one per pass over the template.
 *
 * A pass opens with the next unread record. MERGEFIELD and IF take that
 * record's values; NEXT and a NEXTIF that holds move to the record after it
 * inside the same document; a SKIPIF that holds drops the document. MERGEREC
 * is the record's place in the list and MERGESEQ counts the documents kept.
 * SET and ASK give a bookmark its text, which REF fields to it then show;
 * FILLIN shows its answer where it stands. Questions with no answer yet come
 * back in `prompts` and take their default meanwhile.
 */
export function mergeDocuments(
  doc: PmNode,
  records: readonly MergeRecord[],
  options: MergeRunOptions = {},
): { documents: MergedDocument[]; prompts: MergePrompt[] } {
  const first = options.firstRecord ?? 1
  const documents: MergedDocument[] = []
  const prompts: MergePrompt[] = []
  const bookmarks = new Map<string, string>()
  const ruleNodes = new WeakSet<PmNode>()
  const answer = (walk: MergeWalk, args: string[], promptAt: number): string => {
    const once = args.some((arg) => arg.toLowerCase() === '\\o')
    const key = once ? `${walk.asked}` : `${walk.asked}:${first + walk.index}`
    walk.asked += 1
    const fallback = switchArg(args, '\\d')
    const given = options.answers?.get(key)
    if (given === undefined && ![...prompts, ...walk.prompts].some((p) => p.key === key))
      walk.prompts.push({
        key,
        prompt: args[promptAt] ?? '',
        fallback,
        record: once ? null : first + walk.index,
      })
    return given ?? fallback
  }
  /** what a field leaves in the text; null for a field the merge does not touch */
  const fieldText = (instr: string, walk: MergeWalk): string | null => {
    const record = records[walk.index] ?? {}
    const [keyword] = splitInstruction(instr)
    const args = fieldArgs(instr)
    switch (keyword) {
      case 'MERGEFIELD':
        return mergeFieldText(instr, record)
      case 'IF':
        return ifBranch(instr, record)
      case 'NEXT':
        walk.index += 1
        return ''
      case 'NEXTIF':
        if (conditionHolds(instr, record)) walk.index += 1
        return ''
      case 'SKIPIF':
        if (conditionHolds(instr, record)) walk.skipped = true
        return ''
      case 'MERGEREC':
        return String(first + walk.index)
      case 'MERGESEQ':
        return String(documents.length + 1)
      case 'SET':
        if (args[0]) bookmarks.set(args[0], args[1] ?? '')
        return ''
      case 'ASK': {
        const text = answer(walk, args, 1)
        if (args[0]) bookmarks.set(args[0], text)
        return ''
      }
      case 'FILLIN':
        return answer(walk, args, 0)
      case 'REF':
        return args[0] !== undefined && bookmarks.has(args[0]) ? bookmarks.get(args[0])! : null
      default:
        return null
    }
  }
  const walkContent = (content: readonly PmNode[] | undefined, walk: MergeWalk) => {
    if (!content) return undefined
    const out: PmNode[] = []
    // the field the previous node belonged to
    let open: { instr: string; id: number | null } | null = null
    for (const node of content) {
      const mark = node.marks?.find((m) => m.type === FIELD_MARK)
      const ref = node.marks?.find((m) => m.type === 'refField')
      const instr = mark
        ? String(mark.attrs?.instr ?? '')
        : ref
          ? `REF ${String(ref.attrs?.name ?? '')}`
          : ''
      const id = Number(mark?.attrs?.fieldId) > 0 ? Number(mark?.attrs?.fieldId) : null
      // A field whose result is split over nodes is worked out once: its nodes carry one
      // field id. Two fields side by side carry two, however alike their instructions
      // are, and each acts. Only nodes with no id fall back on the instruction.
      const repeat = instr !== '' && open !== null && open.instr === instr && open.id === id
      const value = instr === '' ? null : repeat ? '' : fieldText(instr, walk)
      if (value !== null && RULE_ONLY.test(instr)) ruleNodes.add(node)
      if (value === null) {
        open = null
        out.push(node.content ? { ...node, content: walkBlocks(node.content, walk) ?? [] } : node)
        continue
      }
      open = { instr, id }
      out.push({
        ...node,
        text: value,
        marks: (node.marks ?? []).filter((m) => m.type !== FIELD_MARK && m.type !== 'refField'),
      })
    }
    return out
  }
  /** blocks and inline content alike; a paragraph its rule fields emptied is left out */
  const walkBlocks = (content: readonly PmNode[] | undefined, walk: MergeWalk) => {
    const merged = walkContent(content, walk)
    return merged?.filter((_, i) => {
      const inline = content![i]!.content ?? []
      return inline.length === 0 || !inline.every((child) => ruleNodes.has(child))
    })
  }
  for (let index = 0, pass = 0; index < records.length && pass < records.length; pass += 1) {
    const walk: MergeWalk = { index, skipped: false, asked: 0, prompts: [] }
    const content = walkBlocks(doc.content, walk)
    if (!walk.skipped) {
      documents.push({ doc: { ...doc, content }, record: records[index]! })
      prompts.push(...walk.prompts)
    }
    index = walk.index + 1
  }
  return { documents, prompts }
}

/**
 * Every record's letter in one document.
 *
 * Word's "Edit Individual Documents" puts the letters one after another, each
 * on its own page; the same shape is built here: the merged documents one
 * after another, with the page break on the first block of every letter but
 * the first.
 */
export function mergeAllRecords(
  doc: PmNode,
  records: readonly MergeRecord[],
  options: MergeRunOptions = {},
): PmNode {
  return joinMergedDocuments(
    doc,
    mergeDocuments(doc, records, options).documents.map((merged) => merged.doc),
  )
}

/** the merged documents as one, each on a page of its own */
export function joinMergedDocuments(doc: PmNode, documents: readonly PmNode[]): PmNode {
  const content: PmNode[] = []
  documents.forEach((merged, index) => {
    ;(merged.content ?? []).forEach((block, blockIndex) => {
      // a bookmark, a comment anchor or a note reference exists once in the
      // package, so only the first letter keeps it
      const copy = index === 0 ? block : withoutUniqueAnchors(block)
      if (index === 0 || blockIndex !== 0) {
        content.push(copy)
        return
      }
      content.push({ ...copy, attrs: { ...(copy.attrs ?? {}), pageBreakBefore: true } })
    })
  })
  return { ...doc, content }
}

function withoutUniqueAnchors(node: PmNode): PmNode {
  const out: PmNode = { ...node }
  if (out.attrs && 'bookmarks' in out.attrs && out.attrs.bookmarks) {
    out.attrs = { ...out.attrs, bookmarks: null }
  }
  if (out.marks?.some((m) => m.type === 'comment')) {
    out.marks = out.marks.filter((m) => m.type !== 'comment')
  }
  if (node.content) {
    out.content = node.content
      .filter((child) => child.type !== 'docNoteRef')
      .map(withoutUniqueAnchors)
  }
  return out
}
