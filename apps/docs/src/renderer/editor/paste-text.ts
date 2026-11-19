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
import { PARA_FORMAT_ATTRS } from './caret-marks'

/**
 * The text of a paste whose entire payload is one single-COLUMN table of
 * cells carrying no formatting of their own, or null for everything else.
 * Companion to the single-cell unwrap: such a paste is semantically lines of
 * text, so it must take the insertion point's formatting like typing — the
 * HTML parse lane keeps no marks for bare cells and built an invisible
 * borderless table instead (fixed layout, source row heights, bottom
 * alignment), which reads as gapped, indented stray lines (r176 follow-up:
 * a column of Sheets cells pasted into Docs; the 1x1 case was r176 itself).
 * The result equals what the plain-text clipboard flavor would paste, so
 * Ctrl+V and paste-as-text agree for unformatted cells. A cell with
 * formatting elements of its own (a rich-text cell), a multi-column table,
 * or prose/media around the table keeps the HTML lane; a style attribute on
 * the td itself is not that — the unwrap never kept it.
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
    // nested tables are structure the HTML lane must keep
    if (table.querySelector('table')) return null
    const rows = [...table.querySelectorAll('tr')]
    if (rows.length === 0) return null
    const lines: string[] = []
    for (const row of rows) {
      const cells = row.querySelectorAll('td,th')
      // a second column means real tabular structure: keep the HTML lane
      if (cells.length !== 1) return null
      // text nodes and explicit line breaks only; any other element is
      // formatting the HTML lane should keep
      let text = ''
      for (const node of cells[0]!.childNodes) {
        if (node.nodeType === 3 /* TEXT_NODE */) text += node.textContent ?? ''
        else if (node.nodeName === 'BR') text += '\n'
        else return null
      }
      lines.push(text)
    }
    const text = lines.join('\n')
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
  // Word Keep Text Only: a paragraph this paste creates formats like pressing
  // Enter at the insertion point — clone the insertion paragraph's formatting
  // (the r178 allowlist; identity/annotation attrs and pageBreakBefore never
  // clone). Only the first line merges into the destination paragraph, so
  // without this, lines 2+ landed as DEFAULT paragraphs and dropped the
  // surrounding style/indent/line-spacing/alignment.
  let attrs: Record<string, unknown> | null = null
  for (let depth = $context.depth; depth >= 1; depth--) {
    const node = $context.node(depth)
    if (node.type === paragraph) {
      const picked: Record<string, unknown> = {}
      for (const key of PARA_FORMAT_ATTRS) if (key in node.attrs) picked[key] = node.attrs[key]
      attrs = picked
      break
    }
  }
  // the default parser's line handling: consecutive breaks collapse to one split
  const blocks = text
    .split(/(?:\r\n?|\n)+/)
    .map((line) => paragraph.create(attrs, line ? schema.text(line, marks) : null))
  // open ends so single-line text merges inline into the destination paragraph
  return new Slice(Fragment.from(blocks), 1, 1)
}
