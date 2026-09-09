/**
 * Word pastes plain text with the formatting typing there would produce
 * ("Keep Text Only" takes the insertion point's format): pending stored marks
 * win, else the marks at the insertion position. ProseMirror's default
 * clipboardTextParser only reads the position's marks — an emptied paragraph
 * has none and its pilcrow memory lives in storedMarks (caret-marks.ts), so
 * pasting after select→Delete fell back to the theme font.
 */
import type { ResolvedPos } from '@tiptap/pm/model'
import { Fragment, Slice } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'

/**
 * The text of a paste whose entire payload is one single-cell table carrying
 * no formatting of its own, or null for everything else. Companion to the
 * single-cell unwrap: that paste is semantically text, so
 * it must take the insertion point's formatting like typing — the HTML parse
 * lane keeps no marks for a bare cell and fell back to the theme font (alpha
 * ledger r176: a Sheets cell pasted into Docs as the theme font). A cell with
 * formatting elements of its own (a rich-text cell) keeps the HTML lane; a
 * style attribute on the td itself is not that — the unwrap never kept it.
 */
export function singleCellPasteText(html: string): string | null {
  if (!/<table/i.test(html)) return null
  try {
    const doc = new window.DOMParser().parseFromString(html, 'text/html')
    const body = doc.body
    const tables = body.querySelectorAll('table')
    if (tables.length !== 1) return null
    const table = tables[0]!
    // same whole-payload guards as the unwrap: prose or media anywhere
    // outside the table keeps the HTML lane
    if ((body.textContent ?? '').trim() !== (table.textContent ?? '').trim()) return null
    if (
      [...body.querySelectorAll('img,svg,video,hr')].some((element) => !table.contains(element))
    ) {
      return null
    }
    const cells = table.querySelectorAll('td,th')
    if (cells.length !== 1) return null
    const cell = cells[0]!
    // text nodes and explicit line breaks only; any other element is
    // formatting the HTML lane should keep
    let text = ''
    for (const node of cell.childNodes) {
      if (node.nodeType === 3 /* TEXT_NODE */) text += node.textContent ?? ''
      else if (node.nodeName === 'BR') text += '\n'
      else return null
    }
    return text.trim() ? text : null
  } catch {
    return null
  }
}

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
