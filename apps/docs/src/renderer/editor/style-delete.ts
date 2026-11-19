import type { Editor } from '@tiptap/core'

const STYLED_BLOCKS = new Set(['docParagraph', 'docHeading', 'docListItem'])

/**
 * Text that uses the style goes back to the default paragraph style (a
 * character style simply comes off), as when Word deletes a style. One
 * undo step; false when nothing used it.
 */
export function unstyleDocument(
  editor: Editor,
  styleId: string,
  type: 'paragraph' | 'character',
): boolean {
  const { state } = editor
  const { docParagraph, docHeading } = state.schema.nodes
  const mark = state.schema.marks.docTextStyle
  const tr = state.tr
  state.doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    if (type === 'paragraph') {
      if (!STYLED_BLOCKS.has(node.type.name) || node.attrs.styleId !== styleId) return false
      const $pos = tr.doc.resolve(pos)
      // a heading that was one only through the style is body text again
      const plain =
        node.type === docHeading &&
        !node.attrs.outlineOnly &&
        docParagraph &&
        $pos.parent.canReplaceWith($pos.index(), $pos.index() + 1, docParagraph)
      tr.setNodeMarkup(pos, plain ? docParagraph : undefined, { ...node.attrs, styleId: null })
      return false
    }
    node.forEach((child, offset) => {
      const own = mark?.isInSet(child.marks)
      if (!own || own.attrs.styleId !== styleId) return
      const from = pos + 1 + offset
      tr.addMark(from, from + child.nodeSize, mark.create({ ...own.attrs, styleId: null }))
    })
    return false
  })
  if (!tr.docChanged) return false
  editor.view.dispatch(tr)
  return true
}
