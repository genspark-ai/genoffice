import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

const key = new PluginKey<boolean>('inactiveSelection')

/**
 * Word keeps the document selection visibly (gray) highlighted while a ribbon
 * combobox has keyboard focus. In the browser, focusing the font family/size
 * <input> RELOCATES the DOM selection into the input, so the document
 * highlight vanishes entirely even though the editor state still holds the
 * selection and formatting will apply to it (alpha ledger r119: users read
 * that as "my text got deselected" and abort). While a ribbon control opts in
 * (focus → on, blur → off), decorate the state selection so it stays visible.
 */
export const InactiveSelectionExtension = Extension.create({
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
            // every range of any selection kind (text, ctrl-A AllSelection,
            // table CellSelection); a NodeSelection keeps its selectednode
            // class without focus, so it needs no decoration
            if (sel.empty || sel instanceof NodeSelection) return DecorationSet.empty
            const decos = sel.ranges
              .filter((r) => r.$from.pos < r.$to.pos)
              .map((r) =>
                Decoration.inline(r.$from.pos, r.$to.pos, { class: 'doc-inactive-selection' }),
              )
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})

/** ribbon controls call this on focus(true)/blur(false) of their text inputs */
export function setInactiveSelectionShown(editor: Editor | null, shown: boolean): void {
  if (!editor || editor.isDestroyed) return
  if ((key.getState(editor.state) ?? false) === shown) return
  editor.view.dispatch(editor.state.tr.setMeta(key, shown).setMeta('addToHistory', false))
}
