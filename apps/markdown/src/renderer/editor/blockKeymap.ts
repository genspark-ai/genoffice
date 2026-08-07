import { Extension } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    blockKeymap: {
      /** duplicate the top-level block at the caret and move the caret into the copy */
      duplicateBlock: () => ReturnType
      /** delete the top-level block at the caret (the last block becomes an empty paragraph) */
      deleteBlock: () => ReturnType
      moveBlockUp: () => ReturnType
      moveBlockDown: () => ReturnType
    }
  }
}

/** [start, end) of the depth-1 block containing the selection head; null at doc level */
function rootRange(
  state: import('@tiptap/pm/state').EditorState,
): { from: number; to: number } | null {
  const sel = state.selection
  // a top-level NodeSelection (block handle, atom blocks like hr/image) IS the range
  if (sel instanceof NodeSelection && sel.$from.depth === 0) {
    return { from: sel.from, to: sel.to }
  }
  const $head = sel.$head
  if ($head.depth < 1) return null
  return { from: $head.before(1), to: $head.after(1) }
}

function moveBlock(direction: -1 | 1) {
  return ({
    state,
    dispatch,
  }: {
    state: import('@tiptap/pm/state').EditorState
    dispatch?: (tr: import('@tiptap/pm/state').Transaction) => void
  }) => {
    const range = rootRange(state)
    if (!range) return false
    const index = state.doc.resolve(range.from).index(0)
    const siblingIndex = index + direction
    if (siblingIndex < 0 || siblingIndex >= state.doc.childCount) return false
    const block = state.doc.child(index)
    const sibling = state.doc.child(siblingIndex)
    const caretOffset = state.selection.head - range.from
    let tr = state.tr
    if (direction === -1) {
      const siblingFrom = range.from - sibling.nodeSize
      tr = tr.delete(range.from, range.to).insert(siblingFrom, block)
      tr = tr.setSelection(TextSelection.near(tr.doc.resolve(siblingFrom + caretOffset)))
    } else {
      const insertAt = range.to + sibling.nodeSize - block.nodeSize
      tr = tr.delete(range.from, range.to).insert(insertAt, block)
      tr = tr.setSelection(TextSelection.near(tr.doc.resolve(insertAt + caretOffset)))
    }
    dispatch?.(tr.scrollIntoView())
    return true
  }
}

/**
 * SecondBrain-parity block shortcuts: ⌘D duplicate, ⌘⇧⌫ delete,
 * ⌘⇧↑/↓ move the current top-level block.
 */
export const BlockKeymap = Extension.create({
  name: 'blockKeymap',

  addCommands() {
    return {
      duplicateBlock:
        () =>
        ({ state, dispatch }) => {
          const range = rootRange(state)
          if (!range) return false
          const block = state.doc.resolve(range.from).nodeAfter
          if (!block) return false
          if (dispatch) {
            let tr = state.tr.insert(range.to, block)
            const caretOffset = state.selection.head - range.from
            tr = tr.setSelection(TextSelection.near(tr.doc.resolve(range.to + caretOffset)))
            dispatch(tr.scrollIntoView())
          }
          return true
        },
      deleteBlock:
        () =>
        ({ state, dispatch, editor }) => {
          const range = rootRange(state)
          if (!range) return false
          if (dispatch) {
            let tr = state.tr.delete(range.from, range.to)
            if (tr.doc.childCount === 0) {
              tr = tr.insert(0, editor.schema.nodes.paragraph!.create())
            }
            tr = tr.setSelection(
              TextSelection.near(tr.doc.resolve(Math.min(range.from, tr.doc.content.size))),
            )
            dispatch(tr.scrollIntoView())
          }
          return true
        },
      moveBlockUp: () => moveBlock(-1),
      moveBlockDown: () => moveBlock(1),
    }
  },

  addKeyboardShortcuts() {
    return {
      'Mod-d': () => this.editor.commands.duplicateBlock(),
      'Mod-D': () => this.editor.commands.duplicateBlock(),
      'Mod-Shift-Backspace': () => this.editor.commands.deleteBlock(),
      'Mod-Shift-ArrowUp': () => this.editor.commands.moveBlockUp(),
      'Mod-Shift-ArrowDown': () => this.editor.commands.moveBlockDown(),
      'Mod-k': () => {
        const editor = this.editor
        if (editor.isActive('link')) {
          return editor.chain().focus().extendMarkRange('link').unsetLink().run()
        }
        const href = window.prompt('URL:')
        if (!href?.trim()) return true
        return editor.chain().focus().extendMarkRange('link').setLink({ href: href.trim() }).run()
      },
    }
  },
})
