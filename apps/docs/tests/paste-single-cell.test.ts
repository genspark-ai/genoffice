/**
 * A copy of one unformatted spreadsheet cell (Sheets/Excel put a 1x1 table on
 * the clipboard) pastes as text with the insertion point's formatting, like
 * typing — not as unmarked text that falls back to the theme font (alpha
 * ledger r176: cell copied in AI Sheets pasted into AI Docs showed the Aptos
 * theme font instead of the destination's Calibri).
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { pasteTextSlice, singleCellPasteText } from '../src/renderer/editor/paste-text'

// jsdom has no ClipboardEvent; the paste path only needs an Event shell
class FakeClipboardEvent extends Event {
  clipboardData = null
}
;(globalThis as Record<string, unknown>).ClipboardEvent = FakeClipboardEvent

const CALIBRI = { font: 'Calibri', fontAscii: 'Calibri', sizeHalfPoints: 24 }

/** the shape Univer's USMToHtmlService writes for one unformatted cell */
const UNIVER_SINGLE_CELL =
  '<google-sheets-html-origin><table data-copy-id="x" xmlns="http://www.w3.org/1999/xhtml" ' +
  'cellspacing="0" cellpadding="0" dir="ltr" ' +
  'style="table-layout:fixed;font-size:10pt;font-family:Arial;width:0px;border-collapse:collapse;border:none">' +
  '<colgroup><col width="72" /></colgroup>' +
  '<tbody><tr style="height:19px;"><td style="vertical-align:bottom;">Budget 2026</td></tr></tbody>' +
  '</table></google-sheets-html-origin>'

describe('singleCellPasteText', () => {
  it('extracts the text of a lone unformatted cell (Univer clipboard shape)', () => {
    expect(singleCellPasteText(UNIVER_SINGLE_CELL)).toBe('Budget 2026')
  })

  it('turns in-cell explicit line breaks into newlines', () => {
    expect(singleCellPasteText('<table><tr><td>a<br>b</td></tr></table>')).toBe('a\nb')
  })

  it('keeps the HTML lane for multi-cell tables', () => {
    expect(singleCellPasteText('<table><tr><td>a</td><td>b</td></tr></table>')).toBeNull()
  })

  it('keeps the HTML lane when prose accompanies the table', () => {
    expect(singleCellPasteText('<p>intro</p><table><tr><td>a</td></tr></table>')).toBeNull()
  })

  it('keeps the HTML lane for a rich-text cell (formatting elements inside)', () => {
    expect(
      singleCellPasteText(
        '<table><tr><td><span style="font-weight:bold;">a</span></td></tr></table>',
      ),
    ).toBeNull()
  })

  it('keeps the HTML lane for images and empty cells', () => {
    expect(
      singleCellPasteText('<table><tr><td><img src="http://x/y.png"></td></tr></table>'),
    ).toBeNull()
    expect(singleCellPasteText('<table><tr><td> </td></tr></table>')).toBeNull()
    expect(singleCellPasteText('no tables here')).toBeNull()
  })
})

describe('single-cell paste takes the insertion point font', () => {
  function makeEditor() {
    return new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Report ',
                marks: [{ type: 'docTextStyle', attrs: CALIBRI }],
              },
            ],
          },
        ],
      },
    })
  }

  it('lands the cell text with the caret marks via pasteTextSlice', () => {
    const editor = makeEditor()
    const { state } = editor.view
    // caret at the end of the Calibri run
    editor.view.dispatch(
      state.tr.setSelection(TextSelection.create(state.doc, 1 + 'Report '.length)),
    )
    const cellText = singleCellPasteText(UNIVER_SINGLE_CELL)
    expect(cellText).toBe('Budget 2026')
    const slice = pasteTextSlice(cellText!, editor.view.state.selection.$from, editor.view)
    editor.view.dispatch(editor.view.state.tr.replaceSelection(slice))
    const para = editor.state.doc.child(0)
    expect(para.textContent).toBe('Report Budget 2026')
    const last = para.child(para.childCount - 1)
    expect(last.marks.find((m) => m.type.name === 'docTextStyle')?.attrs).toMatchObject({
      font: 'Calibri',
    })
    editor.destroy()
  })
})
