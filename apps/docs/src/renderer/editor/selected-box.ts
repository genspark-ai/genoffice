import type { Editor } from '@tiptap/core'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state'
import { getActiveSubEditor } from './active-editor'

/**
 * Which box of a multi-drawing docProtected node the user pressed. A paragraph
 * anchoring several shapes packs them into one node and the NodeSelection
 * covers all of them, so the ribbon needs this to show and edit the pressed
 * box instead of always the first one.
 */
interface SelectedBoxState {
  pos: number
  index: number
}

const key = new PluginKey<SelectedBoxState | null>('docSelectedBox')

export function selectedBoxPlugin(): Plugin<SelectedBoxState | null> {
  return new Plugin<SelectedBoxState | null>({
    key,
    state: {
      init: () => null,
      apply(tr, prev) {
        const meta = tr.getMeta(key) as SelectedBoxState | null | undefined
        if (meta !== undefined) return meta
        if (!prev) return null
        const pos = tr.mapping.map(prev.pos)
        const sel = tr.selection
        if (!(sel instanceof NodeSelection) || sel.from !== pos) return null
        return pos === prev.pos ? prev : { pos, index: prev.index }
      },
    },
  })
}

/** Select the docProtected node at `pos` with `index` as the box the ribbon addresses. */
export function selectBoxAt(tr: Transaction, pos: number, index: number): Transaction {
  return tr.setSelection(NodeSelection.create(tr.doc, pos)).setMeta(key, { pos, index })
}

/** the box whose text the active sub-editor edits: its ordinal among the node's rendered boxes */
function subEditorBoxIndex(sub: Editor): number | null {
  const el = sub.view.dom.closest('.doc-textbox')
  const wrapper = el?.closest('.doc-protected')
  if (!el || !wrapper) return null
  return Array.from(wrapper.querySelectorAll('.doc-textbox')).indexOf(el)
}

/**
 * Index of the addressed box when the selection is a docProtected node; 0
 * otherwise. Typing into a sibling text box moves the address to that box: the
 * press that recorded the index did not go through the floating-object handler.
 */
export function selectedBoxIndex(state: EditorState): number {
  const sel = state.selection
  if (!(sel instanceof NodeSelection)) return 0
  const boxes = sel.node.attrs.textboxes as unknown[] | null
  if (!Array.isArray(boxes)) return 0
  const sub = getActiveSubEditor()
  const subIndex = sub ? subEditorBoxIndex(sub) : null
  if (subIndex !== null && subIndex >= 0 && subIndex < boxes.length) return subIndex
  const s = key.getState(state)
  if (!s || s.pos !== sel.from) return 0
  return s.index < boxes.length ? s.index : 0
}
