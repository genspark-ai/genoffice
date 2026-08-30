import type { ChainedCommands, Editor } from '@tiptap/core'

/**
 * Word's Ctrl+Enter / Insert→Page Break: the paragraph content AFTER the
 * caret starts the next page — no extra empty line. Inserting a fresh empty
 * break-paragraph left a blank line coupled to the break (deleting the line
 * killed the break — alpha ledger r154). Mid-paragraph we split and the
 * second half takes the pageBreakBefore attribute; an empty second half
 * (caret at paragraph end) matches Word, whose break then lives on an empty
 * line at the next page top; at a paragraph start the block itself takes the
 * attribute. The DOCUMENT start is special: pagination deliberately ignores
 * breaks that land on a blank page except for a leading w:br on the first
 * content (Word's blank-first-page case), and the attribute alone would be a
 * no-op there — only that spot inserts a break character instead. Every path
 * dispatches once, so a single undo reverts the whole break.
 */
export function insertPageBreak(editor: Editor): boolean {
  const { $from, empty } = editor.state.selection
  const start = (): ChainedCommands => {
    const chain = editor.chain().focus()
    return empty ? chain : chain.deleteSelection()
  }
  const markCaretBlock = (chain: ChainedCommands): ChainedCommands =>
    chain.command(({ state, tr, dispatch }) => {
      const caret = state.selection.$from
      if (!caret.parent.isTextblock) return false
      if (!('pageBreakBefore' in caret.parent.attrs)) return false
      if (dispatch) {
        tr.setNodeMarkup(caret.before(caret.depth), undefined, {
          ...caret.parent.attrs,
          pageBreakBefore: true,
        })
        dispatch(tr.scrollIntoView())
      }
      return true
    })

  if ($from.parent.isTextblock && 'pageBreakBefore' in $from.parent.attrs) {
    if ($from.parentOffset === 0) {
      if ($from.before($from.depth) === 0) {
        return start()
          .insertContent({ type: 'hardBreak', attrs: { pageBreak: true } })
          .run()
      }
      return markCaretBlock(start()).run()
    }
    return markCaretBlock(start().splitBlock()).run()
  }
  // non-textblock contexts keep the explicit break paragraph
  return start()
    .insertContent({ type: 'docParagraph', attrs: { pageBreakBefore: true } })
    .run()
}
