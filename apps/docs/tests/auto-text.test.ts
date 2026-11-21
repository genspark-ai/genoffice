/** AutoText keeps a selection under a name and puts it back in, formatting and all. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  autoTextEntries,
  deleteAutoText,
  insertAutoText,
  saveAutoText,
} from '../src/renderer/editor/auto-text'
import {
  fontStyleKey,
  revealedFormatting,
  revealedStyleName,
} from '../src/renderer/components/RevealFormattingDialog'
import type { RibbonFormatState } from '../src/renderer/components/ribbon-format-state'

class FakeDataTransfer {
  private data = new Map<string, string>()
  setData(type: string, value: string) {
    this.data.set(type, value)
  }
  getData(type: string) {
    return this.data.get(type) ?? ''
  }
  get types() {
    return [...this.data.keys()]
  }
}
class FakeClipboardEvent extends Event {
  clipboardData: FakeDataTransfer | null
  constructor(type: string, init?: { clipboardData?: FakeDataTransfer }) {
    super(type)
    this.clipboardData = init?.clipboardData ?? null
  }
}
const g = globalThis as Record<string, unknown>
g.DataTransfer = FakeDataTransfer
g.ClipboardEvent = FakeClipboardEvent

describe('AutoText', () => {
  let editor: Editor
  beforeEach(() => {
    localStorage.clear()
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'Yours ' },
              { type: 'text', text: 'sincerely', marks: [{ type: 'bold' }] },
            ],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'end' }] },
        ],
      },
    })
  })
  afterEach(() => editor.destroy())

  it('saves nothing without a selection or a name', () => {
    editor.commands.setTextSelection(3)
    expect(saveAutoText(editor, 'Closing')).toBe(false)
    editor.commands.setTextSelection({ from: 1, to: 16 })
    expect(saveAutoText(editor, '  ')).toBe(false)
    expect(autoTextEntries()).toEqual([])
  })

  it('keeps the selection by name, replaces a same-named entry and lists by name', () => {
    editor.commands.setTextSelection({ from: 1, to: 16 })
    expect(saveAutoText(editor, 'Closing')).toBe(true)
    editor.commands.setTextSelection({ from: 1, to: 6 })
    expect(saveAutoText(editor, 'Address')).toBe(true)
    expect(saveAutoText(editor, ' Closing ')).toBe(true)
    expect(autoTextEntries().map((e) => [e.name, e.text])).toEqual([
      ['Address', 'Yours'],
      ['Closing', 'Yours'],
    ])
    deleteAutoText('Address')
    expect(autoTextEntries().map((e) => e.name)).toEqual(['Closing'])
  })

  it('inserts an entry at the caret with its formatting', () => {
    editor.commands.setTextSelection({ from: 1, to: 16 })
    saveAutoText(editor, 'Closing')
    const end = editor.state.doc.content.size - 1
    editor.commands.setTextSelection(end)
    expect(insertAutoText(editor, autoTextEntries()[0])).toBe(true)
    const last = editor.state.doc.lastChild!
    expect(last.textContent).toBe('endYours sincerely')
    expect(last.lastChild!.marks.map((m) => m.type.name)).toContain('bold')
  })

  it('ignores what storage holds that is not an entry', () => {
    localStorage.setItem('aidocs.autoText', JSON.stringify([{ name: 1 }, 'x']))
    expect(autoTextEntries()).toEqual([])
    localStorage.setItem('aidocs.autoText', '{')
    expect(autoTextEntries()).toEqual([])
  })
})

describe('Reveal Formatting', () => {
  const format = {
    fontFamily: 'Calibri',
    fontSizePt: 11,
    bold: true,
    italic: true,
    underline: true,
    textColor: 'FF0000',
    align: null,
    bidi: false,
    indentLeft: 720,
    indentRight: 0,
    spaceBefore: 120,
    spaceAfter: 160,
    lineSpacing: 1.15,
  } as unknown as RibbonFormatState
  const length = (twips: number) => `${twips / 1440} in`

  it('names the font style as the Font dialog does', () => {
    expect(fontStyleKey({ bold: true, italic: true })).toBe('appFontBoldItalic')
    expect(fontStyleKey({ bold: true, italic: false })).toBe('appFontBold')
    expect(fontStyleKey({ bold: false, italic: true })).toBe('appFontItalic')
    expect(fontStyleKey({ bold: false, italic: false })).toBe('appFontRegular')
  })

  it('names a style by the document, then the built-in table, then its id', () => {
    const styles = new Map([['Quote', { name: 'Block Quote' }]])
    expect(revealedStyleName('Quote', styles)).toBe('Block Quote')
    expect(revealedStyleName('Normal', new Map())).toBe('Normal')
    expect(revealedStyleName('Mystery', null)).toBe('Mystery')
  })

  it('lists the font, then the paragraph', () => {
    const shown = revealedFormatting(format, 'Normal', length)
    expect(shown.font).toEqual([
      ['appFontDialogTitle', 'Calibri'],
      ['appFontSizeLabel', '11 pt'],
      ['appFontStyleLabel', ['appFontBoldItalic']],
      ['appUnderline', '✓'],
      ['appFontColor', '#FF0000'],
    ])
    expect(shown.paragraph).toEqual([
      ['ribbonPbStyle', 'Normal'],
      ['appAlignment', ['appAlignLeft']],
      ['appIndentLeft', '0.5 in'],
      ['appIndentRight', '0 in'],
      ['appSpaceBefore', '6 pt'],
      ['appSpaceAfter', '8 pt'],
      ['appLineSpacingLabel', '1.15'],
    ])
  })

  it('leaves out what is not set, and reads a right-to-left paragraph as right aligned', () => {
    const plain = { ...format, underline: false, textColor: null, lineSpacing: null, bidi: true }
    const shown = revealedFormatting(plain as RibbonFormatState, 'Normal', length)
    expect(shown.font.map(([key]) => key)).toEqual([
      'appFontDialogTitle',
      'appFontSizeLabel',
      'appFontStyleLabel',
    ])
    expect(shown.paragraph[1]).toEqual(['appAlignment', ['appAlignRight']])
    expect(shown.paragraph.map(([key]) => key)).not.toContain('appLineSpacingLabel')
  })
})
