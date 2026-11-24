/** Fit Text, Two Lines in One and Horizontal in Vertical on selected text, and what is saved for them. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  eaLayoutDecls,
  fitTextDecls,
  fitTextOf,
  parseEaLayoutAttr,
  setEaLayout,
  setFitText,
  textStretch,
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
    expect(setEaLayout(editor, { kind: 'combine', brackets: null })).toBe(false)
    expect(setFitText(editor, 1440)).toBe(false)
    editor.commands.setTextSelection({ from: 6, to: 20 })
    expect(textStretch(editor.state)).toBeNull()
  })

  it('turns two lines in one on, sized to half the text, and off again as an explicit none', () => {
    selectWord()
    expect(setEaLayout(editor, { kind: 'combine', brackets: null })).toBe(true)
    expect(layout()).toMatchObject({ combine: true, w: 3 })
    expect(layout()!.brackets).toBeUndefined()
    expect(layout()!.id).toBeGreaterThan(0)
    selectWord()
    setEaLayout(editor, null)
    expect(layout()).toEqual({})
  })

  it('puts brackets around two lines in one, changes them, and keeps the region id', () => {
    selectWord()
    setEaLayout(editor, { kind: 'combine', brackets: 'square' })
    const { id } = layout()!
    expect(layout()).toMatchObject({ combine: true, brackets: 'square' })
    expect(editor.view.dom.querySelector('[data-ea-brackets="square"]')?.textContent).toBe('abcde')
    selectWord()
    setEaLayout(editor, { kind: 'combine', brackets: 'curly' })
    expect(layout()).toMatchObject({ combine: true, brackets: 'curly', id })
    selectWord()
    setEaLayout(editor, { kind: 'combine', brackets: null })
    expect(editor.view.dom.querySelector('[data-ea-brackets]')).toBeNull()
  })

  it('horizontal in vertical replaces two lines in one: the two exclude each other', () => {
    selectWord()
    setEaLayout(editor, { kind: 'combine', brackets: 'round' })
    selectWord()
    setEaLayout(editor, { kind: 'vert', fit: true })
    expect(layout()).toMatchObject({ vert: true, vertCompress: true })
    expect(layout()!.combine).toBeUndefined()
    expect(layout()!.brackets).toBeUndefined()
    // the run is marked for the stylesheet, fitted or not, and carries no drawing of its own
    const run = () => editor.view.dom.querySelector<HTMLElement>('[data-ea-vert]')
    expect(run()?.textContent).toBe('abcde')
    expect(run()!.getAttribute('style') ?? '').not.toMatch(/writing-mode|text-combine/)
    selectWord()
    setEaLayout(editor, { kind: 'vert', fit: false })
    expect(layout()!.vertCompress).toBeUndefined()
    expect(run()?.textContent).toBe('abcde')
    selectWord()
    setEaLayout(editor, { kind: 'combine', brackets: null })
    expect(run()).toBeNull()
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
    // horizontal in vertical is the stylesheet's: fitted or not, no inline drawing
    expect(eaLayoutDecls({ vert: true, vertCompress: true }, null)).toEqual([])
    expect(eaLayoutDecls({ vert: true }, null)).toEqual([])
    expect(eaLayoutDecls({}, null)).toEqual([])
    expect(twoLinesWidthEm('abcde')).toBe(3)
    expect(twoLinesWidthEm('')).toBe(1)
  })
})
