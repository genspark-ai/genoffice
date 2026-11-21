/** Fit Text, Two Lines in One and Horizontal in Vertical on selected text, and what is saved for them. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  eaLayoutDecls,
  fitTextDecls,
  fitTextOf,
  parseEaLayoutAttr,
  setFitText,
  textStretch,
  toggleEaLayout,
  twoLinesWidthEm,
} from '../src/renderer/editor/asian-layout'

describe('asian layout of selected text', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'head abcde tail' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'second' }] },
        ],
      },
    })
  })
  afterEach(() => editor.destroy())

  const layout = () => parseEaLayoutAttr(editor.getAttributes('docTextStyle').eaLayout)
  const selectWord = () => editor.commands.setTextSelection({ from: 6, to: 11 })

  it('needs a stretch of text inside one paragraph', () => {
    selectWord()
    expect(textStretch(editor.state)).toEqual({ from: 6, to: 11, text: 'abcde' })
    editor.commands.setTextSelection(6)
    expect(textStretch(editor.state)).toBeNull()
    expect(toggleEaLayout(editor, 'combine')).toBe(false)
    expect(setFitText(editor, 1440)).toBe(false)
    editor.commands.setTextSelection({ from: 6, to: 20 })
    expect(textStretch(editor.state)).toBeNull()
  })

  it('turns two lines in one on, sized to half the text, and off again as an explicit none', () => {
    selectWord()
    expect(toggleEaLayout(editor, 'combine')).toBe(true)
    expect(layout()).toMatchObject({ combine: true, w: 3 })
    expect(layout()!.id).toBeGreaterThan(0)
    selectWord()
    toggleEaLayout(editor, 'combine')
    expect(layout()).toEqual({})
  })

  it('horizontal in vertical replaces two lines in one: the two exclude each other', () => {
    selectWord()
    toggleEaLayout(editor, 'combine')
    selectWord()
    toggleEaLayout(editor, 'vert')
    expect(layout()).toMatchObject({ vert: true })
    expect(layout()!.combine).toBeUndefined()
  })

  it('fits the text into a width, changes it and takes it off', () => {
    selectWord()
    expect(fitTextOf(editor)).toBe(0)
    setFitText(editor, 1440.4)
    selectWord()
    expect(fitTextOf(editor)).toBe(1440)
    setFitText(editor, 0)
    selectWord()
    expect(fitTextOf(editor)).toBe(0)
    expect(editor.getAttributes('docTextStyle').fitText).toBe(0)
  })

  it('draws them: a fixed width with spread characters, two half-size lines, upright in vertical text', () => {
    expect(fitTextDecls(1440)).toEqual(
      expect.arrayContaining(['display:inline-block', 'width:72pt', 'text-align-last:justify']),
    )
    expect(eaLayoutDecls({ combine: true, w: 3 }, 24)).toEqual(
      expect.arrayContaining(['display:inline-block', 'font-size:6pt', 'width:3em']),
    )
    expect(eaLayoutDecls({ combine: true, w: 2 }, null)).toContain('font-size:50%')
    expect(eaLayoutDecls({ vert: true }, null)).toEqual(['text-combine-upright:all'])
    expect(eaLayoutDecls({}, null)).toEqual([])
    expect(twoLinesWidthEm('abcde')).toBe(3)
    expect(twoLinesWidthEm('')).toBe(1)
  })
})
