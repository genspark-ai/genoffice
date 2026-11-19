import type { Editor } from '@tiptap/core'

/**
 * Word's Alt+Shift+Up / Alt+Shift+Down: move the blocks the selection touches
 * past their neighbor.
 *
 * The neighbor is what actually travels — cutting one node and re-inserting it
 * on the other side of the range leaves the moved paragraphs' own content (and
 * the caret riding along with it, through ProseMirror's selection mapping)
 * untouched, which a delete-and-reinsert of the selection would not.
 */
export function moveBlocks(editor: Editor, dir: -1 | 1): boolean {
  return editor
    .chain()
    .focus()
    .command(({ state, tr, dispatch }) => {
      const { doc, selection } = state
      const first = doc.resolve(selection.from).index(0)
      const $to = doc.resolve(selection.to)
      // a selection ending exactly at a block start does not include that block
      const last =
        $to.parentOffset === 0 && $to.depth > 0 && $to.index(0) > first
          ? $to.index(0) - 1
          : $to.index(0)
      const neighborIndex = dir < 0 ? first - 1 : last + 1
      if (neighborIndex < 0 || neighborIndex >= doc.childCount) return false
      const neighbor = doc.child(neighborIndex)
      let neighborStart = 0
      for (let i = 0; i < neighborIndex; i++) neighborStart += doc.child(i).nodeSize
      let rangeStart = 0
      for (let i = 0; i < first; i++) rangeStart += doc.child(i).nodeSize
      let rangeEnd = rangeStart
      for (let i = first; i <= last; i++) rangeEnd += doc.child(i).nodeSize
      if (!dispatch) return true
      tr.delete(neighborStart, neighborStart + neighbor.nodeSize)
      tr.insert(dir < 0 ? tr.mapping.map(rangeEnd) : tr.mapping.map(rangeStart), neighbor)
      dispatch(tr)
      return true
    })
    .run()
}
