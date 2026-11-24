import type { Editor } from '@tiptap/core'
import { Slice } from '@tiptap/pm/model'
import {
  carriableBlockXml,
  carriableSectionBreakXml,
  docxBookmarkNames,
  docxBookmarkRange,
  docxRootNamespaces,
  includeTextInstruction,
  nextNoteId,
  parseDocx,
  type NoteInfo,
  type ParsedDoc,
} from '@genoffice/docx-engine'
import { renumber } from '../ai/note-ops'
import { fetchDocBytes } from '../doc-bytes'
import { blockFingerprint, whileCarrying, type CarriedNodes } from './carried-nodes'
import { blocksToPmDoc, pmDocOptions, type PmNode } from './convert'
import { pastePayload } from './paste-actions'

type NoteKind = 'footnote' | 'endnote'

/** the notes of the document being inserted into, and the way to add to them */
export interface InsertFileNotes {
  list(kind: NoteKind): NoteInfo[]
  add(kind: NoteKind, notes: NoteInfo[]): void
}

export interface InsertFileSource {
  parsed: ParsedDoc
  /** namespace declarations of the other document's root, for the objects taken as raw XML */
  namespaces?: Record<string, string>
  /** the range lies inside one paragraph: it goes in as text, not as a paragraph of its own */
  inline?: boolean
  /** only the text goes in, taking the formatting of where it lands */
  plain?: boolean
  /** section breaks come along with their page setup, and are counted here */
  sections?: { count: number }
}

interface Payload {
  html: string
  text: string
  leftOut: number
  carried: CarriedNodes
  notes: Record<NoteKind, NoteInfo[]>
}

/**
 * A protected block as the paste can bring it in. A picture travels in the
 * markup itself. A shape, a text box, an equation or a field that needs
 * nothing but its own XML is looked up while the markup is parsed; a chart,
 * a diagram or anything else with parts of its own in the other file stays
 * behind.
 */
function carriedBlock(
  node: PmNode,
  source: InsertFileSource,
  carried: CarriedNodes,
): PmNode | null {
  const attrs = node.attrs ?? {}
  if (attrs.blockType === 'image') {
    const src = attrs.imageDataUrl
    return typeof src === 'string' && src.startsWith('data:') ? node : null
  }
  const index = attrs.docxIndex
  const original = source.parsed.blocks.find((b) => b.docxIndex === index && !b.hidden)
  const raw = attrs.blockType === 'passthrough' && source.namespaces ? original?.originalXml : null
  if (!raw || !source.namespaces) return null
  const section =
    source.sections && !source.plain ? carriableSectionBreakXml(raw, source.namespaces) : null
  const xml = section ?? carriableBlockXml(raw, source.namespaces)
  if (!xml) return null
  if (section) source.sections!.count++
  const own = { ...attrs, docxIndex: null, genXml: xml, strayList: null }
  carried.blocks.set(String(index), { ...own, genPristine: blockFingerprint(own) })
  return node
}

/** note references get an id of this document, and their notes are queued to join its list */
function carryNotes(
  nodes: PmNode[],
  source: InsertFileSource,
  host: InsertFileNotes | undefined,
  out: Payload,
): PmNode[] {
  const taken: Record<NoteKind, NoteInfo[]> = { footnote: [], endnote: [] }
  const given = { footnote: new Map<string, string>(), endnote: new Map<string, string>() }
  const visit = (node: PmNode): PmNode | null => {
    if (node.type === 'docNoteRef') {
      const kind: NoteKind = node.attrs?.kind === 'endnote' ? 'endnote' : 'footnote'
      const from = (kind === 'footnote' ? source.parsed.footnotes : source.parsed.endnotes).find(
        (n) => n.id === String(node.attrs?.id),
      )
      if (!host || !from) return null
      // marks that share a note in the other document share one here
      let id = given[kind].get(from.id)
      if (id === undefined) {
        id = nextNoteId([...host.list(kind), ...taken[kind]])
        given[kind].set(from.id, id)
        const note: NoteInfo = { id, text: from.text, created: true }
        if (from.richParas) note.richParas = from.richParas
        if (from.noRefMark) note.noRefMark = true
        taken[kind].push(note)
      }
      const attrs = { ...node.attrs, id }
      out.carried.notes.set(id, [...(out.carried.notes.get(id) ?? []), attrs])
      return { ...node, attrs }
    }
    if (!node.content) return node
    return { ...node, content: node.content.map(visit).filter((n): n is PmNode => n !== null) }
  }
  const mapped = nodes.map(visit).filter((n): n is PmNode => n !== null)
  out.notes = taken
  return mapped
}

/**
 * The body of a parsed document as a paste: the markup, its plain text, how
 * many blocks stay behind, and what the markup alone cannot say.
 */
export function fileBodyPayload(
  editor: Editor,
  source: InsertFileSource,
  host?: InsertFileNotes,
): Payload {
  const { parsed } = source
  const out: Payload = {
    html: '',
    text: '',
    leftOut: 0,
    carried: { blocks: new Map(), notes: new Map() },
    notes: { footnote: [], endnote: [] },
  }
  const blocks = blocksToPmDoc(parsed.blocks, undefined, pmDocOptions(parsed)).content ?? []
  const kept = blocks
    .map((node) => (node.type === 'docProtected' ? carriedBlock(node, source, out.carried) : node))
    .filter((n): n is PmNode => n !== null)
  out.leftOut = blocks.length - kept.length
  if (kept.length === 0) return out
  const doc = editor.schema.nodeFromJSON({
    type: 'doc',
    content: carryNotes(kept, source, host, out),
  })
  const open = source.inline && doc.childCount === 1 && doc.firstChild!.isTextblock ? 1 : 0
  out.html = editor.view.serializeForClipboard(new Slice(doc.content, open, open)).dom.innerHTML
  out.text = doc.textBetween(0, doc.content.size, '\n\n')
  return out
}

/**
 * Word's Insert > File: the body of another document goes in at the
 * selection, keeping its formatting, the way pasting it would. Returns how
 * many blocks could not be brought in, or null when there was nothing to insert.
 */
export function insertFileBody(
  editor: Editor,
  source: InsertFileSource,
  host?: InsertFileNotes,
): number | null {
  const { html, text, leftOut, carried, notes } = fileBodyPayload(editor, source, host)
  if (!html) return null
  // the notes first: a reference must never be in the body without its note
  for (const kind of ['footnote', 'endnote'] as const)
    if (notes[kind].length > 0) host?.add(kind, notes[kind])
  const pasted = whileCarrying(carried, () =>
    pastePayload(editor, { html, text, image: null }, source.plain ? 'text' : 'source'),
  )
  if (!pasted) return null
  if (carried.notes.size > 0) {
    // the marks arrive with the numbers they had in their own document
    const tr = editor.state.tr
    renumber(tr, 'footnote')
    renumber(tr, 'endnote')
    // joins the paste's undo step
    if (tr.docChanged) editor.view.dispatch(tr.setMeta('appendedTransaction', tr))
  }
  return leftOut
}

export type InsertFileResult =
  | { status: 'canceled' | 'password' | 'empty' }
  | { status: 'inserted'; name: string; leftOut: number; unlinked?: true; sections?: number }
  | { status: 'failed'; error: string }

export interface InsertFileHost {
  notes: InsertFileNotes
  /**
   * Asks how the picked file goes in: the range (a bookmark's name, or null
   * for the whole document) and whether it stays linked to the file;
   * undefined to give up.
   */
  askOptions(
    fileName: string,
    bookmarks: string[],
  ): Promise<{ range: string | null; link: boolean } | undefined>
  /** inserts the body as the result of a link field with this instruction */
  insertLinked(
    source: InsertFileSource,
    notes: InsertFileNotes,
    instr: string,
  ): { leftOut: number | null; linked: boolean }
}

/** Asks for a document and inserts its body, or the range of one of its bookmarks, at the selection. */
export async function pickAndInsertFile(
  editor: Editor,
  host: InsertFileHost,
): Promise<InsertFileResult> {
  const picked = await window.desktop.pickDocx()
  if (!picked) return { status: 'canceled' }
  // a protected document is not run through the password prompt here
  if ('needsPassword' in picked) return { status: 'password' }
  try {
    let bytes = await fetchDocBytes(picked.dataUrl)
    let inline = false
    const options = await host.askOptions(picked.name, await docxBookmarkNames(bytes))
    if (options === undefined) return { status: 'canceled' }
    if (options.range !== null) {
      const range = await docxBookmarkRange(bytes, options.range)
      if (!range) return { status: 'empty' }
      bytes = range.bytes
      inline = range.inline
    }
    const namespaces = await docxRootNamespaces(bytes)
    const parsed = await parseDocx(bytes)
    const source = { parsed, namespaces, inline }
    // a linked body is a field result, which holds no section of its own
    const sections = { count: 0 }
    const { leftOut, linked } = options.link
      ? host.insertLinked(source, host.notes, includeTextInstruction(picked.path, options.range))
      : { leftOut: insertFileBody(editor, { ...source, sections }, host.notes), linked: true }
    if (leftOut === null) return { status: 'empty' }
    return {
      status: 'inserted',
      name: picked.name,
      leftOut,
      ...(linked ? {} : { unlinked: true }),
      ...(sections.count > 0 ? { sections: sections.count } : {}),
    }
  } catch (err) {
    return { status: 'failed', error: String(err) }
  }
}
