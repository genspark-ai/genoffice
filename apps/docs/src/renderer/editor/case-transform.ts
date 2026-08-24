import type { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'

export type CaseMode = 'upper' | 'lower' | 'title' | 'sentence'

export function transformCase(s: string, mode: CaseMode): string {
  switch (mode) {
    case 'upper':
      return s.toUpperCase()
    case 'lower':
      return s.toLowerCase()
    case 'title':
      return s.toLowerCase().replace(/(^|\s)(\p{L})/gu, (m) => m.toUpperCase())
    case 'sentence':
      return s.toLowerCase().replace(/(^\s*\p{L})|([.!?。!?]\s*\p{L})/gu, (m) => m.toUpperCase())
  }
}

/**
 * Word's Shift+F3 cycle: lowercase → UPPERCASE → Capitalize Each Word. Which
 * step comes next is read off the selection, so repeated presses walk the ring
 * (mixed-case text enters it at lowercase, like Word).
 */
export function nextCaseMode(text: string): CaseMode {
  const letters = text.replace(/\P{L}/gu, '')
  if (!letters) return 'lower'
  if (letters === letters.toLowerCase()) return 'upper'
  if (letters === letters.toUpperCase()) return 'title'
  return 'lower'
}

/** rewrite every text run in the selection, keeping its marks */
export function applyCase(editor: Editor, mode: CaseMode): boolean {
  const { from, to } = editor.state.selection
  if (from === to) return false
  return editor
    .chain()
    .focus()
    .command(({ state, tr }) => {
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText || !node.text) return
        const start = Math.max(from, pos)
        const end = Math.min(to, pos + node.nodeSize)
        const slice = node.text.slice(start - pos, end - pos)
        const next = transformCase(slice, mode)
        if (next !== slice) {
          tr.replaceWith(
            tr.mapping.map(start),
            tr.mapping.map(end),
            state.schema.text(next, node.marks),
          )
        }
      })
      // Restore the range: without this, replacing the leading run pushes the
      // anchor to its end and the next Shift+F3 only sees the tail of the
      // phrase. Every rewrite happened inside the selection, so the document's
      // growth is the selection's growth (ß → SS and friends lengthen it).
      if (tr.docChanged) {
        const grew = tr.doc.content.size - state.doc.content.size
        tr.setSelection(TextSelection.create(tr.doc, from, to + grew))
      }
      return true
    })
    .run()
}

/** the selection's text, used to decide where the Shift+F3 ring resumes */
export function selectionText(editor: Editor): string {
  const { from, to } = editor.state.selection
  return editor.state.doc.textBetween(from, to, '\n', '\n')
}
