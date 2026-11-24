import type { Editor } from '@tiptap/core'
import {
  docxAutoTextBody,
  docxAutoTextEntries,
  docxRootNamespaces,
  parseDocx,
} from '@genoffice/docx-engine'
import { insertFileBody, type InsertFileNotes } from './insert-file'
import { pastePayload } from './paste-actions'

const KEY = 'aidocs.autoText'

/** a piece of the document kept under a name, to insert again: formatting and all */
export interface AutoTextEntry {
  name: string
  /** the saved selection as the clipboard would carry it, so it goes back in as a paste does */
  html: string
  text: string
  /** kept in the open document's glossary part, not on this computer; its content is read on insert */
  from?: 'document'
}

const isEntry = (value: unknown): value is AutoTextEntry =>
  typeof value === 'object' &&
  value !== null &&
  typeof (value as AutoTextEntry).name === 'string' &&
  typeof (value as AutoTextEntry).html === 'string' &&
  typeof (value as AutoTextEntry).text === 'string'

/** the saved entries, by name */
export function autoTextEntries(): AutoTextEntry[] {
  try {
    const parsed: unknown = JSON.parse(localStorage.getItem(KEY) ?? '[]')
    return Array.isArray(parsed) ? parsed.filter(isEntry) : []
  } catch {
    return []
  }
}

function store(entries: AutoTextEntry[]): boolean {
  try {
    localStorage.setItem(KEY, JSON.stringify(entries))
    return true
  } catch {
    return false
  }
}

/** the start of the selection: what a new entry is called unless renamed ('' with no selection) */
export function selectionSnippet(editor: Editor): string {
  const { from, to } = editor.state.selection
  return editor.state.doc.textBetween(from, to, ' ').trim().slice(0, 32)
}

/** the selection as text and clipboard HTML, when it stays inside one paragraph */
export function selectionClip(editor: Editor): { text: string; html: string } | null {
  const { selection, doc } = editor.state
  if (selection.empty || !selection.$from.sameParent(selection.$to)) return null
  if (!selection.$from.parent.isTextblock) return null
  return {
    text: doc.textBetween(selection.from, selection.to, ' '),
    html: editor.view.serializeForClipboard(selection.content()).dom.innerHTML,
  }
}

/** Keep the selection under `name`, replacing an entry of that name. False when nothing is selected. */
export function saveAutoText(editor: Editor, name: string): boolean {
  const { selection, doc } = editor.state
  const title = name.trim()
  if (selection.empty || !title) return false
  const entry: AutoTextEntry = {
    name: title,
    html: editor.view.serializeForClipboard(selection.content()).dom.innerHTML,
    text: doc.textBetween(selection.from, selection.to, '\n\n'),
  }
  const others = autoTextEntries().filter((e) => e.name !== title)
  return store([...others, entry].sort((a, b) => a.name.localeCompare(b.name)))
}

export function deleteAutoText(name: string): void {
  store(autoTextEntries().filter((e) => e.name !== name))
}

/** Put an entry in at the selection, the way pasting it would */
export function insertAutoText(editor: Editor, entry: AutoTextEntry): boolean {
  return pastePayload(editor, { html: entry.html, text: entry.text, image: null })
}

/** the AutoText entries the document's own glossary part holds (a document made from a template keeps the template's) */
export async function documentAutoText(bytes: Uint8Array): Promise<AutoTextEntry[]> {
  try {
    return (await docxAutoTextEntries(bytes)).map((e) => ({ ...e, html: '', from: 'document' }))
  } catch {
    return []
  }
}

/** Put an entry of the document's glossary part in at the selection, with its formatting */
export async function insertDocumentAutoText(
  editor: Editor,
  bytes: Uint8Array,
  name: string,
  notes: InsertFileNotes,
): Promise<boolean> {
  try {
    const body = await docxAutoTextBody(bytes, name)
    if (!body) return false
    const namespaces = await docxRootNamespaces(body.bytes)
    const parsed = await parseDocx(body.bytes)
    return insertFileBody(editor, { parsed, namespaces, inline: body.inline }, notes) !== null
  } catch {
    return false
  }
}
