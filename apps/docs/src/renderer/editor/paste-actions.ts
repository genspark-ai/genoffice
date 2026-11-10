/**
 * Clipboard pastes started from the UI (ribbon Paste, its split menu, Paste
 * Special, the context menu). Every entry runs the editor's full paste
 * pipeline with a synthesized clipboard event, so App's handlePaste lanes and
 * the paste-mode handshake behave exactly as on a native Ctrl+V.
 */
import type { Editor } from '@tiptap/core'
import { insertImageFromDataUrl } from '../components/ribbon-tabs'
import {
  beginForeignPaste,
  defaultPasteMode,
  forceNextPasteMode,
  stashPastePayload,
  type PasteMode,
} from './paste-options'
import { isForeignPasteHtml } from './paste-web-html'

export type ClipboardPayload = { html: string; text: string; image: Blob | null }

/** A paste-special choice: one of the paste modes, or the bitmap flavor. */
export type PasteChoice = PasteMode | 'picture'

export async function readClipboardPayload(): Promise<ClipboardPayload> {
  const payload: ClipboardPayload = { html: '', text: '', image: null }
  try {
    for (const item of await navigator.clipboard.read()) {
      if (!payload.text && item.types.includes('text/plain')) {
        payload.text = await (await item.getType('text/plain')).text()
      }
      if (!payload.html && item.types.includes('text/html')) {
        payload.html = await (await item.getType('text/html')).text()
      }
      const imageType = item.types.find((type) => type.startsWith('image/'))
      if (!payload.image && imageType) payload.image = await item.getType(imageType)
    }
  } catch {
    try {
      payload.text = await navigator.clipboard.readText()
    } catch {
      /* clipboard unavailable or denied */
    }
  }
  return payload
}

function clipboardEvent(html: string | null, text: string): ClipboardEvent {
  const data = new DataTransfer()
  if (html) data.setData('text/html', html)
  if (text) data.setData('text/plain', text)
  return new ClipboardEvent('paste', { clipboardData: data })
}

function htmlText(html: string): string {
  try {
    const body = new window.DOMParser().parseFromString(html, 'text/html').body
    for (const block of body.querySelectorAll('p,div,li,tr,h1,h2,h3,h4,h5,h6,br')) {
      block.after(body.ownerDocument.createTextNode('\n'))
    }
    return (body.textContent ?? '').replace(/\n{2,}/g, '\n').trim()
  } catch {
    return ''
  }
}

function insertBlobImage(editor: Editor, blob: Blob): void {
  const reader = new FileReader()
  reader.onload = () => {
    if (typeof reader.result === 'string') {
      void insertImageFromDataUrl(editor, reader.result, 'Image (pasted)')
    }
  }
  reader.readAsDataURL(blob)
}

/**
 * Paste an already-read payload. Without a mode the user's default applies
 * (plain Ctrl+V); a mode pastes like picking that entry on the paste-options
 * chip. Returns false when the payload has nothing the choice can paste.
 */
export function pastePayload(
  editor: Editor,
  payload: ClipboardPayload,
  mode?: PasteChoice,
): boolean {
  const { html, text, image } = payload
  if (mode === 'picture') {
    if (!image) return false
    insertBlobImage(editor, image)
    return true
  }
  if (html) {
    // Only the document editor installs the transformPasted / handlePaste
    // lanes that consume the mode handshake; text-box and header/footer
    // sub-editors must not arm it (the mode would leak into the next paste)
    const consumes = Boolean(editor.view.someProp('transformPasted'))
    if (consumes) {
      if (mode) forceNextPasteMode(mode)
      if (beginForeignPaste(html)) {
        stashPastePayload({ html, text, mode: mode ?? defaultPasteMode() })
        editor.view.pasteHTML(html, clipboardEvent(html, text))
        editor.commands.focus()
        return true
      }
    }
    // the default mode only ever applies to HTML from other programs
    const textOnly = mode
      ? mode === 'text'
      : !consumes && isForeignPasteHtml(html) && defaultPasteMode() === 'text'
    // our own clipboard HTML keeps its exact runs; Keep Text Only still
    // strips them like Word does for an in-document copy
    if (!textOnly) {
      editor.view.pasteHTML(html, clipboardEvent(html, text))
      editor.commands.focus()
      return true
    }
    const plain = text || htmlText(html)
    if (!plain) return false
    editor.view.pasteText(plain, clipboardEvent(null, plain))
    editor.commands.focus()
    return true
  }
  // image priority mirrors Ctrl+V: a bitmap wins over missing or
  // whitespace-only plain text (OS clipboards often advertise an empty
  // text/plain beside image/png)
  if (image && !text.trim() && mode !== 'text') {
    insertBlobImage(editor, image)
    return true
  }
  if (!text) return false
  editor.view.pasteText(text, clipboardEvent(null, text))
  editor.commands.focus()
  return true
}

export async function pasteFromClipboard(editor: Editor, mode?: PasteMode): Promise<boolean> {
  return pastePayload(editor, await readClipboardPayload(), mode)
}
