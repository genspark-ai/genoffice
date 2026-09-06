/**
 * Word pastes plain text with the formatting typing there would produce
 * ("Keep Text Only" takes the insertion point's format): pending stored marks
 * win, else the marks at the insertion position. ProseMirror's default
 * clipboardTextParser only reads the position's marks — an emptied paragraph
 * has none and its pilcrow memory lives in storedMarks (caret-marks.ts), so
 * pasting after select→Delete fell back to the theme font (alpha ledger r172).
 */
import type { ResolvedPos } from '@tiptap/pm/model'
import { Fragment, Slice } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'

export function pasteTextSlice(text: string, $context: ResolvedPos, view: EditorView): Slice {
  // same precedence as typing (storedMarks ?? marks at caret); an explicit
  // empty array from a user toggle is respected, like typing would
  const marks = view.state.storedMarks ?? $context.marks()
  const schema = view.state.schema
  const paragraph = schema.nodes.docParagraph
  // the default parser's line handling: consecutive breaks collapse to one split
  const blocks = text
    .split(/(?:\r\n?|\n)+/)
    .map((line) => paragraph.create(null, line ? schema.text(line, marks) : null))
  // open ends so single-line text merges inline into the destination paragraph
  return new Slice(Fragment.from(blocks), 1, 1)
}
