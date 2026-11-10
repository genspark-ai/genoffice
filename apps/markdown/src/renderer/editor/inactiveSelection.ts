import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

const key = new PluginKey<boolean>('inactiveSelection')

/**
 * Focusing the Ask-AI input relocates the DOM selection into the input, so the
 * document highlight vanishes even though the editor state still holds the
 * range the AI will act on. While a caller opts in, decorate the state
 * selection so it stays visible (docs parity).
 */
export const InactiveSelection = Extension.create({
  name: 'inactiveSelection',

  addProseMirrorPlugins() {
    return [
      new Plugin<boolean>({
        key,
        state: {
          init: () => false,
          apply(tr, shown) {
            const next = tr.getMeta(key) as boolean | undefined
            return next ?? shown
          },
        },
        props: {
          decorations(state) {
            if (!key.getState(state)) return DecorationSet.empty
            const sel = state.selection
            // a NodeSelection keeps its selectednode class without focus
            if (sel.empty || sel instanceof NodeSelection) return DecorationSet.empty
            const decos = sel.ranges
              .filter((r) => r.$from.pos < r.$to.pos)
              .map((r) =>
                Decoration.inline(r.$from.pos, r.$to.pos, { class: 'md-inactive-selection' }),
              )
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})

export function setInactiveSelectionShown(editor: Editor | null, shown: boolean): void {
  if (!editor || editor.isDestroyed) return
  if ((key.getState(editor.state) ?? false) === shown) return
  editor.view.dispatch(editor.state.tr.setMeta(key, shown).setMeta('addToHistory', false))
}
