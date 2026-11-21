import type { Editor } from '@tiptap/core'

const listeners = new WeakMap<Editor, Set<() => void>>()

/**
 * Tiptap's setEditable emits `update` by default, which every subscriber reads
 * as a content change (the document went dirty on a protection or Read Mode
 * switch). Editability is view state: it is announced on its own channel.
 */
export function setEditorEditable(editor: Editor, editable: boolean): void {
  if (editor.isDestroyed) return
  const was = editor.isEditable
  editor.setEditable(editable, false)
  if (was === editor.isEditable) return
  for (const fn of [...(listeners.get(editor) ?? [])]) fn()
}

export function onEditableChange(editor: Editor, fn: () => void): () => void {
  let set = listeners.get(editor)
  if (!set) listeners.set(editor, (set = new Set()))
  set.add(fn)
  return () => {
    set.delete(fn)
  }
}
