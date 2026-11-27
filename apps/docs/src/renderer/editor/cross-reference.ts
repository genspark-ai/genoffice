/**
 * Bookmarks and cross-references: the lists the Bookmark and Cross-reference
 * dialogs show, and the edits they make. A bookmark sits on a paragraph; a
 * reference to a heading, a caption or a note goes through a hidden `_Ref`
 * bookmark on it (for a note, on its reference mark), as in Word.
 */
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { NodeSelection, Selection, TextSelection, type Transaction } from '@tiptap/pm/state'
import {
  bookmarkNamesOf,
  type Block,
  type NoteInfo,
  type NumberingDef,
} from '@genoffice/docx-engine'
import { collectCaptions } from '../commands/reference-commands'
import type { ListNumberingStorage } from './extensions'
import { noteMarkOf, relativeNoteRef } from './note-ref-fields'
import { numberRefSwitches, numberedItems, paragraphNumberAt } from './number-refs'
import { RANGE_BOOKMARK_MARK as RANGE_MARK, bookmarkRange } from './bookmark-range'
import { TRACK_IGNORE } from './revisions'
import { xmlOfNode } from './toc-update'

export interface BookmarkEntry {
  name: string
  /** position of the paragraph or inline field that carries it */
  pos: number
  preview: string
  hidden: boolean
}

const BOOKMARK_HOSTS = ['docParagraph', 'docHeading', 'docListItem']

const GEN_HIDDEN_BOOKMARK = /<w:bookmarkStart\b[^>]*\bw:name="(_[^"]*)"/g

const namesOf = (node: PmNode, attr: 'bookmarks' | 'hiddenBookmarks'): string[] => {
  const names = node.attrs?.[attr] as string[] | null | undefined
  return Array.isArray(names) ? names : []
}

/** bookmarks in document order; Word's internal ones (`_Ref`, `_Toc`…) only on request */
export function collectBookmarks(doc: PmNode, withHidden = false): BookmarkEntry[] {
  const out: BookmarkEntry[] = []
  doc.descendants((node, pos) => {
    const preview = (
      node.type.name === 'docMathTypeField' ? String(node.attrs.text) : node.textContent
    ).slice(0, 40)
    if (withHidden)
      for (const name of namesOf(node, 'hiddenBookmarks'))
        out.push({ name, pos, preview, hidden: true })
    for (const name of namesOf(node, 'bookmarks')) out.push({ name, pos, preview, hidden: false })
    // a caption's `_Ref` lives in its own XML until the file is read again
    if (withHidden && typeof node.attrs?.genXml === 'string')
      for (const m of node.attrs.genXml.matchAll(GEN_HIDDEN_BOOKMARK))
        out.push({ name: m[1], pos, preview, hidden: true })
    return !node.isLeaf
  })
  return out
}

export const sortBookmarks = (list: BookmarkEntry[], by: 'name' | 'location'): BookmarkEntry[] =>
  by === 'location'
    ? list
    : [...list].sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }))

/** a letter first, then letters, digits and underscores, 40 at most: Word's rule */
export const validBookmarkName = (name: string): boolean =>
  name.length <= 40 && /^\p{L}[\p{L}\p{N}_]*$/u.test(name)

function hostAtCursor(editor: Editor): number | null {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth >= 1; depth--)
    if (BOOKMARK_HOSTS.includes($from.node(depth).type.name)) return $from.before(depth)
  return null
}

/** puts the bookmark on the cursor's paragraph; a name already in use moves there, as Word redefines it */
export function addBookmark(editor: Editor, name: string): boolean {
  const at = hostAtCursor(editor)
  if (at === null) return false
  const { tr, doc } = editor.state
  doc.descendants((node, pos) => {
    const names = namesOf(node, 'bookmarks')
    if (pos !== at && names.includes(name)) clearBookmark(tr, node, pos, 'bookmarks', name)
    return !node.isLeaf
  })
  const host = doc.nodeAt(at)
  if (!host) return false
  const names = namesOf(host, 'bookmarks')
  if (!names.includes(name)) tr.setNodeAttribute(at, 'bookmarks', [...names, name])
  const mark = editor.schema.marks[RANGE_MARK]?.create({ name })
  if (mark) {
    tr.removeMark(0, doc.content.size, mark)
    const { selection } = editor.state
    // the selected text is what the bookmark covers; with a caret it stands for its paragraph
    if (selection instanceof TextSelection && !selection.empty)
      tr.addMark(selection.from, selection.to, mark)
  }
  if (tr.docChanged) editor.view.dispatch(tr.setMeta(TRACK_IGNORE, true))
  return true
}

function clearBookmark(
  tr: Transaction,
  node: PmNode,
  pos: number,
  attr: 'bookmarks' | 'hiddenBookmarks',
  name: string,
): void {
  if (node.type.name === 'docMathTypeField') {
    const xml = String(node.attrs.xml)
    const start = [...xml.matchAll(/<w:bookmarkStart\b[^>]*\/>/g)].find(([tag]) =>
      bookmarkNamesOf(tag)[attr]?.includes(name),
    )
    const idOf = (tag: string) => /\bw:id\s*=\s*["']([^"']+)["']/.exec(tag)?.[1]
    const id = start && idOf(start[0])
    if (id) {
      const updated = xml.replace(/<w:bookmark(?:Start|End)\b[^>]*\/>/g, (tag) =>
        idOf(tag) === id ? '' : tag,
      )
      tr.setNodeAttribute(pos, 'xml', updated)
    }
  }
  const rest = namesOf(node, attr).filter((n) => n !== name)
  tr.setNodeAttribute(pos, attr, rest.length > 0 ? rest : null)
}

export function removeBookmark(editor: Editor, entry: BookmarkEntry): void {
  const node = editor.state.doc.nodeAt(entry.pos)
  if (!node) return
  const genXml = node.attrs?.genXml as string | null | undefined
  const start =
    entry.hidden && genXml
      ? new RegExp(`<w:bookmarkStart w:id="(\\d+)" w:name="${entry.name}"/>`).exec(genXml)
      : null
  if (start && genXml) {
    const xml = genXml.replace(start[0], '').replace(`<w:bookmarkEnd w:id="${start[1]}"/>`, '')
    editor.view.dispatch(editor.state.tr.setNodeAttribute(entry.pos, 'genXml', xml))
    return
  }
  const attr = entry.hidden ? 'hiddenBookmarks' : 'bookmarks'
  const tr = editor.state.tr
  clearBookmark(tr, node, entry.pos, attr, entry.name)
  const mark = editor.schema.marks[RANGE_MARK]?.create({ name: entry.name })
  if (mark && !entry.hidden) tr.removeMark(0, tr.doc.content.size, mark)
  editor.view.dispatch(tr.setMeta(TRACK_IGNORE, true))
}

/** Word's Go To: the bookmarked paragraph's text is selected and brought on screen */
export function goToBookmark(editor: Editor, entry: BookmarkEntry): void {
  const { state, view } = editor
  const node = state.doc.nodeAt(entry.pos)
  if (!node) return
  const range = entry.hidden ? null : bookmarkRange(state.doc, entry.name)
  const sel =
    node.type.name === 'docMathTypeField'
      ? NodeSelection.create(state.doc, entry.pos)
      : TextSelection.between(
          state.doc.resolve(range?.from ?? entry.pos + 1),
          state.doc.resolve(range?.to ?? entry.pos + node.nodeSize - 1),
        )
  view.dispatch(state.tr.setSelection(sel).scrollIntoView())
  const dom = view.nodeDOM(entry.pos)
  if (dom instanceof HTMLElement) dom.scrollIntoView?.({ block: 'center' })
}

/** what a reference points at: a numbered paragraph, a bookmark, a heading, a note, or the captions of one label */
export type RefType =
  | { kind: 'numbered' }
  | { kind: 'bookmark' }
  | { kind: 'heading' }
  | { kind: 'footnote' }
  | { kind: 'endnote' }
  | { kind: 'caption'; label: string }

/**
 * what the reference shows: the target's text, its page, or where it is from here;
 * a note's number as plain text or as a reference mark looks
 */
export type RefContent =
  | 'text'
  | 'page'
  | 'aboveBelow'
  | 'noteNumber'
  | 'noteNumberFormatted'
  | 'paraNumber'
  | 'paraNumberNoContext'
  | 'paraNumberFullContext'

/** the REF switch of each way a paragraph number is shown */
const PARA_NUMBER_SWITCH: Partial<Record<RefContent, string>> = {
  paraNumber: 'r',
  paraNumberNoContext: 'n',
  paraNumberFullContext: 'w',
}

export const isParaNumber = (content: RefContent): boolean => content in PARA_NUMBER_SWITCH

export interface RefTarget {
  pos: number
  text: string
  /** the bookmark a field can name; null until `anchorTarget` makes one */
  bookmark: string | null
  level?: number
  /** a footnote or endnote: what its reference mark shows */
  noteMark?: string
  /** what a reference to its text shows, when the list shows more than that */
  refText?: string
}

export type RefNotes = Record<'footnote' | 'endnote', NoteInfo[]>

const NOTE_PREVIEW_LENGTH = 60

const REF_BOOKMARK = /<w:bookmarkStart\b[^>]*w:name="(_Ref\d+)"/

export function listRefTargets(
  editor: Editor,
  blocks: Block[],
  type: RefType,
  notes?: RefNotes,
): RefTarget[] {
  const { doc } = editor.state
  if (type.kind === 'numbered') {
    const defs = editor.storage.listNumbering.defs as Map<string, NumberingDef>
    return numberedItems(doc, defs)
      .filter((item) => item.node.type.name === 'docListItem' && item.site?.text.trim())
      .map((item) => ({
        pos: item.pos,
        text: `${item.site!.text.trim()} ${item.node.textContent}`,
        refText: item.node.textContent,
        bookmark: namesOf(item.node, 'hiddenBookmarks').find((n) => n.startsWith('_Ref')) ?? null,
        level: item.site!.ilvl + 1,
      }))
  }
  if (type.kind === 'footnote' || type.kind === 'endnote') {
    const out: RefTarget[] = []
    const seen = new Set<string>()
    doc.descendants((node, pos) => {
      if (node.type.name !== 'docNoteRef') return true
      const id = String(node.attrs.id)
      if (node.attrs.kind !== type.kind || seen.has(id)) return false
      seen.add(id)
      const noteMark = noteMarkOf(node)
      const note = notes?.[type.kind].find((n) => n.id === id)
      // a custom mark opens the note's own text already
      const body = (note?.text ?? '').replace(/\s+/g, ' ').trim()
      const text = note?.customMark ? body : `${noteMark} ${body}`
      out.push({
        pos,
        text: text.slice(0, NOTE_PREVIEW_LENGTH).trim(),
        bookmark: namesOf(node, 'hiddenBookmarks').find((n) => n.startsWith('_Ref')) ?? null,
        noteMark,
      })
      return false
    })
    return out
  }
  if (type.kind === 'bookmark')
    return collectBookmarks(doc).map((b) => ({
      pos: b.pos,
      text: b.preview || b.name,
      bookmark: b.name,
    }))
  if (type.kind === 'heading') {
    const out: RefTarget[] = []
    const { defs, styles } = editor.storage.listNumbering as ListNumberingStorage
    // a heading numbered through its style is read as a list item
    const numbered = new Map(numberedItems(doc, defs).map((item) => [item.pos, item.site]))
    doc.descendants((node, pos) => {
      const style =
        node.type.name === 'docListItem' ? styles?.get(String(node.attrs.styleId ?? '')) : undefined
      // a style set back to body text is no heading, as in the navigation pane
      const styled = style?.headingOutlineOff ? undefined : style?.headingLevel
      if ((node.type.name === 'docHeading' || styled) && node.textContent.trim()) {
        const marker = numbered.get(pos)?.text.trim()
        out.push({
          pos,
          text: marker ? `${marker} ${node.textContent}` : node.textContent,
          refText: node.textContent,
          bookmark: namesOf(node, 'hiddenBookmarks').find((n) => n.startsWith('_Ref')) ?? null,
          level: styled ?? Number(node.attrs.level ?? 1),
        })
      }
      return !node.isLeaf
    })
    return out
  }
  return collectCaptions(editor, blocks)
    .filter((c) => c.label === type.label)
    .map((c) => {
      const node = doc.nodeAt(c.pos)
      const xml = node ? xmlOfNode(node as never, blocks) : ''
      return { pos: c.pos, text: c.text, bookmark: REF_BOOKMARK.exec(xml)?.[1] ?? null }
    })
}

const newRefName = (): string => `_Ref${Math.floor(1e8 + Math.random() * 9e8)}`

// a caption keeps its own XML, so its bookmark is written into it; an id this high cannot meet the ids a save hands out
const newRefId = (): number => 0x40000000 + Math.floor(Math.random() * 0x1000000)

function wrapParagraphXml(xml: string, name: string): string | null {
  const open = /^<w:p(?:\s[^>]*)?>(?:<w:pPr>[\s\S]*?<\/w:pPr>|<w:pPr\/>)?/.exec(xml)
  if (!open || !xml.endsWith('</w:p>')) return null
  const id = newRefId()
  return (
    open[0] +
    `<w:bookmarkStart w:id="${id}" w:name="${name}"/>` +
    xml.slice(open[0].length, -'</w:p>'.length) +
    `<w:bookmarkEnd w:id="${id}"/></w:p>`
  )
}

export interface CrossRefWords {
  above: string
  below: string
  onPage: (page: number) => string
}

export interface CrossRefRequest {
  target: RefTarget
  content: RefContent
  hyperlink: boolean
  /** the target was picked as a heading: one without a number shows 0 */
  heading?: boolean
  /** Word's "Include above/below", for a page number or a paragraph number */
  includeAboveBelow: boolean
  /** Word's "Separate numbers with", for a paragraph number in full context */
  separator?: string
  /** page of the target and of the cursor, when the layout knows them */
  pages: { target: number; here: number } | null
  words: CrossRefWords
}

/** the field text and instruction of a reference to `name` */
export function crossRefField(
  name: string,
  req: Omit<CrossRefRequest, 'target'> & {
    text: string
    before: boolean
    noteMark?: string
    /** the number a REF with this instruction shows for the target */
    paraNumber?: (instr: string) => string | null
  },
): {
  text: string
  mark: { type: 'refField' | 'instrField'; attrs: Record<string, unknown> }
  /** the result looks like a reference mark (NOTEREF \\f) */
  superscript?: true
} {
  const link = req.hyperlink ? ' \\h' : ''
  const where = req.before ? req.words.above : req.words.below
  if (req.noteMark !== undefined && req.content !== 'page') {
    const formatted = req.content === 'noteNumberFormatted'
    const relative = req.content === 'aboveBelow'
    return {
      text: relative ? relativeNoteRef(req.noteMark, where) : req.noteMark,
      mark: {
        type: 'instrField',
        attrs: {
          instr: `NOTEREF ${name}${formatted ? ' \\f' : ''}${relative ? ' \\p' : ''}${link}`,
          dirty: false,
        },
      },
      ...(formatted ? { superscript: true } : {}),
    }
  }
  if (req.content === 'page') {
    const relative = req.includeAboveBelow
    const page = req.pages?.target
    const text =
      relative && req.pages && req.pages.target === req.pages.here
        ? where
        : page === undefined
          ? ' '
          : relative
            ? req.words.onPage(page)
            : String(page)
    return {
      text,
      // Word lays the pages out itself, so it recomputes the number on open
      mark: {
        type: 'instrField',
        attrs: { instr: `PAGEREF ${name}${relative ? ' \\p' : ''}${link}`, dirty: true },
      },
    }
  }
  const numbered = PARA_NUMBER_SWITCH[req.content]
  if (numbered) {
    const separator =
      numbered === 'w' && req.separator ? ` \\d "${req.separator.replace(/"/g, '')}"` : ''
    const instr = ` REF ${name} \\${numbered}${req.includeAboveBelow ? ' \\p' : ''}${separator}${link} `
    return {
      text: req.paraNumber?.(instr) || ' ',
      mark: { type: 'refField', attrs: { name, instr, dirty: false } },
    }
  }
  const relative = req.content === 'aboveBelow'
  return {
    text: relative ? where : req.text,
    mark: {
      type: 'refField',
      attrs: { name, instr: ` REF ${name}${relative ? ' \\p' : ''}${link} `, dirty: false },
    },
  }
}

const FIELD_OR_LINK = new Set(['instrField', 'refField', 'link', 'comment'])

/** inserts the reference at the cursor, giving a heading or caption its `_Ref` bookmark first */
export function insertCrossReference(
  editor: Editor,
  blocks: Block[],
  req: CrossRefRequest,
): boolean {
  const { state } = editor
  const { target } = req
  const node = state.doc.nodeAt(target.pos)
  if (!node) return false
  const tr = state.tr
  let name = target.bookmark
  if (!name) {
    name = newRefName()
    if (BOOKMARK_HOSTS.includes(node.type.name) || node.type.name === 'docNoteRef') {
      tr.setNodeAttribute(target.pos, 'hiddenBookmarks', [
        ...namesOf(node, 'hiddenBookmarks'),
        name,
      ])
    } else {
      const xml = wrapParagraphXml(xmlOfNode(node as never, blocks), name)
      if (!xml) return false
      tr.setNodeMarkup(target.pos, undefined, { ...node.attrs, docxIndex: null, genXml: xml })
    }
  }
  let { from, to, $from } = state.selection
  if (!$from.parent.isTextblock) {
    // a selected picture or table stays: the reference goes to the text after it
    const near = Selection.near(state.doc.resolve(to), 1)
    // nothing after it: near falls back inside the selected block, which is not the place
    if (!near.$from.parent.isTextblock || near.from < to) return false
    from = to = near.from
    $from = near.$from
  }
  const field = crossRefField(name, {
    ...req,
    text: target.refText ?? target.text,
    before: target.pos < from,
    noteMark: target.noteMark,
    paraNumber: (instr) => {
      const switches = numberRefSwitches(instr)
      const defs = editor.storage.listNumbering.defs as Map<string, NumberingDef>
      const number = switches
        ? paragraphNumberAt(state.doc, defs, target.pos, from, switches, req.words)
        : null
      // a heading without a number: Word's result is 0
      return number ?? (req.heading ? '0' : null)
    },
  })
  const kept = (state.storedMarks ?? $from.marks()).filter((m) => !FIELD_OR_LINK.has(m.type.name))
  const mark = state.schema.marks[field.mark.type].create(field.mark.attrs)
  const style = state.schema.marks.docTextStyle
  const styled = kept.find((m) => m.type === style)
  // a note's number is raised only when asked for as a reference mark, whatever the caret stands in
  const shown =
    target.noteMark !== undefined && style && (field.superscript || styled)
      ? [
          ...kept.filter((m) => m !== styled),
          style.create({ ...styled?.attrs, vertAlign: field.superscript ? 'superscript' : null }),
        ]
      : kept
  tr.replaceWith(from, to, state.schema.text(field.text, [...shown, mark]))
  editor.view.dispatch(tr.scrollIntoView())
  return true
}
