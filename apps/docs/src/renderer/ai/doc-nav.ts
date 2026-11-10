import type { Editor } from '@tiptap/core'
import { Selection, TextSelection } from '@tiptap/pm/state'
import { blockRangePositions } from './protocol'

/**
 * In-answer citations: the model links passages as [label](docnav://block/N);
 * clicking one selects and scrolls to that top-level block.
 */

export const DOC_NAV_SCHEME = 'docnav://'

/** docnav://block/N -> N; null for anything else */
export function parseDocNavHref(href: string): number | null {
  const m = /^docnav:\/\/block\/(\d+)$/.exec(href)
  return m ? Number(m[1]) : null
}

/** select block `index` (clamped) and scroll it into view */
export function navigateToBlock(editor: Editor, index: number): void {
  const doc = editor.state.doc
  const clamped = Math.min(Math.max(index, 0), doc.childCount - 1)
  const node = doc.child(clamped)
  const { from, to } = blockRangePositions(editor, clamped, clamped)
  // atoms (images/charts/fields) and tables are not valid TextSelection
  // endpoints — TextSelection.create throws there; settle near them instead
  const selection = node.isTextblock
    ? TextSelection.create(doc, from + 1, to - 1)
    : Selection.near(doc.resolve(from), 1)
  editor.view.dispatch(editor.state.tr.setSelection(selection).scrollIntoView())
  editor.view.focus()
}
