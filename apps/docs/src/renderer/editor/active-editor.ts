import type { Editor } from '@tiptap/core'

/**
 * Tracks which textbox sub-editor (if any) should receive ribbon commands.
 * The last-focused sub-editor stays active until the main editor takes focus
 * back, so clicking a ribbon button (which briefly steals focus) still routes
 * the command into the textbox — mirroring how Word keeps the shape selected.
 */

type Listener = (docChanged: boolean) => void

let active: Editor | null = null
const listeners = new Set<Listener>()

function emit(docChanged: boolean): void {
  for (const fn of listeners) fn(docChanged)
}

/** the sub-editor ribbon commands should target, or null for the main editor */
export function getActiveSubEditor(): Editor | null {
  if (active?.isDestroyed) active = null
  return active
}

/** pass null when the main editor regains focus */
export function setActiveSubEditor(editor: Editor | null): void {
  if (active === editor) return
  active = editor
  emit(false)
}

/** deactivate a sub-editor that is being destroyed */
export function dropActiveSubEditor(editor: Editor): void {
  if (active === editor) {
    active = null
    emit(false)
  }
}

/** sub-editors call this on every transaction so the ribbon re-renders */
export function notifySubEditorState(docChanged: boolean): void {
  emit(docChanged)
}

/** App subscribes to re-render the ribbon and mark the document dirty */
export function subscribeSubEditorState(fn: Listener): () => void {
  listeners.add(fn)
  return () => {
    listeners.delete(fn)
  }
}
