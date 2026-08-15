import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

const key = new PluginKey<DecorationSet>('aiHighlight')

/** Mark a freshly AI-written range on a transaction; decorations map with later edits */
export function markAiRange(tr: Transaction, from: number, to: number): Transaction {
  return tr.setMeta('aiHighlightAdd', { from, to })
}

/** Remove every AI-change highlight (called when a run finishes) */
export function clearAiHighlights(editor: Editor): void {
  editor.view.dispatch(editor.state.tr.setMeta('aiHighlightClear', true))
}

/**
 * Transient yellow wash over content the AI just wrote, so the user can see
 * what changed during a run. Display-only: decorations never touch the
 * document or the serialized markdown.
 */
export const AiHighlight = Extension.create({
  name: 'aiHighlight',

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            if (tr.getMeta('aiHighlightClear')) return DecorationSet.empty
            let next = set.map(tr.mapping, tr.doc)
            const add = tr.getMeta('aiHighlightAdd') as { from: number; to: number } | undefined
            if (add) {
              const from = Math.max(0, Math.min(add.from, tr.doc.content.size))
              const to = Math.max(from, Math.min(add.to, tr.doc.content.size))
              if (to > from) {
                next = next.add(tr.doc, [Decoration.inline(from, to, { class: 'ai-changed' })])
              }
            }
            return next
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
      }),
    ]
  },
})
