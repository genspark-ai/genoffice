/** Word's Drop Cap dialog reads the paragraph's drop cap and keeps the line count in range. */
import { describe, expect, it } from 'vitest'
import { dropCapValueOf } from '../src/renderer/components/DropCapDialog'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  clampDropCapLines,
  dropCapFontOf,
  dropCapGapOf,
  dropCapLinesOf,
  setDropCapFont,
} from '../src/renderer/editor/drop-cap'

describe('drop cap dialog', () => {
  it('opens on the paragraph as it is', () => {
    expect(dropCapValueOf(null)).toEqual({ position: 'none', lines: 3, gap: 0 })
    expect(dropCapValueOf('{"type":"drop","lines":4}')).toEqual({
      position: 'drop',
      lines: 4,
      gap: 0,
    })
    expect(dropCapValueOf('{"type":"drop","lines":4,"hSpace":144}').gap).toBe(144)
    expect(dropCapValueOf('{"type":"margin"}')).toEqual({
      position: 'margin',
      lines: 3,
      gap: 0,
    })
    expect(dropCapValueOf('not json')).toEqual({ position: 'none', lines: 3, gap: 0 })
    expect(dropCapValueOf('{"type":"sideways","lines":2}')).toEqual({
      position: 'none',
      lines: 3,
      gap: 0,
    })
  })

  it('keeps the lines to drop between 1 and 10', () => {
    expect(clampDropCapLines('0')).toBe(1)
    expect(clampDropCapLines('25')).toBe(10)
    expect(clampDropCapLines('2.6')).toBe(3)
    expect(clampDropCapLines('')).toBe(1)
    expect(clampDropCapLines('abc')).toBe(3)
  })

  it('gives the page the lines the letter spans, three when the paragraph does not say', () => {
    expect(dropCapLinesOf('{"type":"drop","lines":5}')).toBe(5)
    expect(dropCapLinesOf('{"type":"drop"}')).toBe(3)
    expect(dropCapLinesOf('{"type":"drop","lines":40}')).toBe(10)
    expect(dropCapLinesOf('not json')).toBe(3)
    expect(dropCapLinesOf(null)).toBe(3)
  })

  it('gives the page the distance from the text, none when the paragraph does not say', () => {
    expect(dropCapGapOf('{"type":"drop","hSpace":144}')).toBe(144)
    expect(dropCapGapOf('{"type":"drop"}')).toBe(0)
    expect(dropCapGapOf('{"type":"drop","hSpace":-5}')).toBe(0)
    expect(dropCapGapOf('{"type":"drop","hSpace":99999}')).toBe(31680)
    expect(dropCapGapOf(null)).toBe(0)
  })

  it('reads and sets the font of the first letter, in the slot the letter is drawn from', () => {
    const paragraph = (text: string, attrs: object) => ({
      type: 'docParagraph',
      content: [{ type: 'text', text, marks: [{ type: 'docTextStyle', attrs }] }],
    })
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          paragraph('Once', { fontAscii: 'Arial', eastAsiaFont: 'SimSun' }),
          paragraph('\u6708 rises', { fontAscii: 'Arial', eastAsiaFont: 'SimSun' }),
        ],
      },
    })
    editor.commands.setTextSelection(2)
    expect(dropCapFontOf(editor.state)).toEqual({ font: 'Arial', eastAsian: false })
    setDropCapFont(editor, 'Georgia')
    expect(dropCapFontOf(editor.state)).toEqual({ font: 'Georgia', eastAsian: false })
    // only the letter changed
    expect(editor.state.doc.firstChild?.child(1).marks[0].attrs.fontAscii).toBe('Arial')
    editor.commands.setTextSelection(9)
    expect(dropCapFontOf(editor.state)).toEqual({ font: 'SimSun', eastAsian: true })
    editor.destroy()
  })
})
