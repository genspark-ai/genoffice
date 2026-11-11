/**
 * The file name a never-saved markdown document takes from its own content:
 * the first heading, else the first few words.
 *
 * Lives outside `App.tsx` so the file-naming module can use it as its fallback
 * without importing the component, which would be a cycle: App imports the
 * naming module. Re-exported from `App` so existing importers keep working.
 */
import type { Editor } from '@tiptap/core'

/** File name for an AI-generated untitled document: first heading, else first words */
export function deriveAutoFileName(editor: Editor): string {
  const doc = editor.state.doc
  for (let i = 0; i < doc.childCount; i++) {
    const node = doc.child(i)
    const text = node.textContent.replace(/\s+/g, ' ').trim()
    if (!text) continue
    if (node.type.name === 'heading') return text.slice(0, 60)
    return text.split(' ').slice(0, 8).join(' ').slice(0, 60)
  }
  return ''
}
