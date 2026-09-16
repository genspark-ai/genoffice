import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import type { NoteInfo } from '@genoffice/docx-engine'
import { posAfterText } from './after-text'

export type NoteKind = 'footnote' | 'endnote'

/**
 * The App-owned footnote / endnote lists, handed to the tool executor. The
 * lists live outside the ProseMirror document (like comments); the executor
 * places and removes the in-text reference marks itself.
 */
export interface AiNotesAccess {
  list(kind: NoteKind): NoteInfo[]
  /** appends a note and returns its id */
  add(kind: NoteKind, text: string): string
  /** false when no note has that id */
  remove(kind: NoteKind, id: string): boolean
  /** block index of a protected block whose XML still holds the reference mark (nothing here can remove it) */
  protectedMarkBlock?(kind: NoteKind, id: string): number | null
}

export interface NoteAnchor {
  blockIndex: number
  pos: number
  num: number
}

/** in-text reference marks keyed "kind:id", in document order */
export function noteAnchors(doc: PmNode): Map<string, NoteAnchor> {
  const out = new Map<string, NoteAnchor>()
  const count: Record<NoteKind, number> = { footnote: 0, endnote: 0 }
  let index = 0
  doc.forEach((block, offset) => {
    block.descendants((node, pos) => {
      if (node.type.name !== 'docNoteRef') return true
      const kind = node.attrs.kind as NoteKind
      count[kind]++
      out.set(`${kind}:${node.attrs.id}`, {
        blockIndex: index,
        pos: offset + 1 + pos,
        num: count[kind],
      })
      return false
    })
    index++
  })
  return out
}

/** where a new reference mark goes: right after `afterText` (live text of the block), else the block end */
export function noteInsertPos(
  doc: PmNode,
  blockIndex: number,
  afterText?: string,
): { pos: number } | { error: string } {
  if (!Number.isInteger(blockIndex) || blockIndex < 0 || blockIndex >= doc.childCount) {
    return {
      error: `block index invalid or out of range (the document has ${doc.childCount} blocks); call get_document_context for fresh indexes`,
    }
  }
  const block = doc.child(blockIndex)
  let start = 0
  doc.forEach((node, offset, i) => {
    if (i === blockIndex) start = offset
  })
  if (block.type.name === 'docProtected') {
    return { error: `block ${blockIndex} is protected content; a note mark cannot be placed in it` }
  }
  if (afterText === undefined) {
    let end = start + block.nodeSize - 1
    if (!block.isTextblock) {
      // containers (tables): the last paragraph inside
      block.descendants((node, pos) => {
        if (node.isTextblock) end = start + 1 + pos + node.nodeSize - 1
        return !node.isTextblock
      })
    }
    return { pos: end }
  }
  return posAfterText(block, start + 1, afterText, `block ${blockIndex}`)
}

function renumber(tr: Transaction, kind: NoteKind): void {
  let num = 0
  tr.doc.descendants((node, pos) => {
    if (node.type.name !== 'docNoteRef' || node.attrs.kind !== kind) return true
    num++
    if (node.attrs.num !== num) tr.setNodeMarkup(pos, undefined, { ...node.attrs, num })
    return false
  })
}

/** insert the reference mark and renumber the marks of that kind in document order */
export function insertNoteRef(editor: Editor, kind: NoteKind, id: string, pos: number): void {
  const type = editor.schema.nodes.docNoteRef!
  const tr = editor.state.tr.insert(pos, type.create({ kind, id, num: 1 }))
  renumber(tr, kind)
  editor.view.dispatch(tr)
}

/** remove every reference mark of that note; returns how many were removed */
export function removeNoteRefs(editor: Editor, kind: NoteKind, id: string): number {
  const tr = editor.state.tr
  const hits: Array<{ pos: number; size: number }> = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'docNoteRef') return true
    if (node.attrs.kind === kind && String(node.attrs.id) === id)
      hits.push({ pos, size: node.nodeSize })
    return false
  })
  for (const hit of hits.reverse()) tr.delete(hit.pos, hit.pos + hit.size)
  renumber(tr, kind)
  if (tr.steps.length > 0) editor.view.dispatch(tr)
  return hits.length
}

/** reference marks inside protected blocks live only in the block XML */
export function protectedNoteMarkBlock(
  doc: PmNode,
  xmlOf: (block: PmNode) => string,
  kind: NoteKind,
  id: string,
): number | null {
  const re = new RegExp(
    `<w:${kind}Reference\\b[^>]*\\bw:id="${id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`,
  )
  let found: number | null = null
  doc.forEach((block, _offset, index) => {
    if (found === null && block.type.name === 'docProtected' && re.test(xmlOf(block))) found = index
  })
  return found
}

const clip = (s: string, max: number) => (s.length > max ? `${s.slice(0, max - 1)}…` : s)

/** model-facing list: one line per note with id, number, anchored block and text */
export function buildNotesContext(
  doc: PmNode,
  footnotes: NoteInfo[],
  endnotes: NoteInfo[],
  protectedBlockOf?: (kind: NoteKind, id: string) => number | null,
): string {
  const anchors = noteAnchors(doc)
  const lines: string[] = []
  const section = (kind: NoteKind, notes: NoteInfo[]) => {
    if (notes.length === 0) return
    lines.push(`${kind === 'footnote' ? 'Footnotes' : 'Endnotes'} (${notes.length}):`)
    for (const note of notes) {
      const a = anchors.get(`${kind}:${note.id}`)
      const locked = a ? null : protectedBlockOf?.(kind, note.id)
      const where = a
        ? `mark ${a.num} in block ${a.blockIndex}`
        : locked !== null && locked !== undefined
          ? `mark in protected block ${locked}, cannot be deleted here`
          : 'no reference mark in the text'
      lines.push(`- id ${note.id} (${where}): ${clip(note.text.replace(/\s+/g, ' ').trim(), 300)}`)
    }
  }
  section('footnote', footnotes)
  section('endnote', endnotes)
  return lines.length ? lines.join('\n') : '(the document has no footnotes or endnotes)'
}
