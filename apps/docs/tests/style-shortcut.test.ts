/**
 * Issue #126 wave 2: paragraph-style and line-spacing commands backing the
 * ⌥⌘0-3 / ⌘1·2·5 shortcuts. applyParagraphStyle was extracted from the ribbon
 * gallery closure; these pin its Word-like behaviors (node switch + shedding
 * the runs' direct font/size/color).
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildBlankDocx, parseDocx } from '@genoffice/docx-engine'
import {
  applyParagraphStyle,
  applyParagraphStyleId,
  setParaAttrs,
} from '../src/renderer/components/ribbon-tabs'

const styled = { type: 'docTextStyle', attrs: { sizeHalfPoints: 48, color: 'FF0000' } }

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'chapter title', marks: [styled] }],
        },
        { type: 'docParagraph', content: [{ type: 'text', text: 'body text' }] },
      ],
    },
  })
}

describe('applyParagraphStyle', () => {
  it('switches the block to a heading and sheds direct size/color covering the whole text', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'h2')
    const block = editor.state.doc.child(0)
    expect(block.type.name).toBe('docHeading')
    expect(block.attrs.level).toBe(2)
    expect(block.attrs.styleId).toBeNull()
    const marks = block.firstChild!.marks.filter((m) => m.type.name === 'docTextStyle')
    expect(marks.length).toBe(0)
    editor.destroy()
  })

  it('returns a heading to a normal paragraph', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'h1')
    applyParagraphStyle(editor, 'p')
    expect(editor.state.doc.child(0).type.name).toBe('docParagraph')
    editor.destroy()
  })

  it('carries the document\u2019s own style id for the level and the default style', async () => {
    const { styles } = await parseDocx(await buildBlankDocx())
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'h1', styles)
    expect(editor.state.doc.child(0).attrs.styleId).toBe('Heading1')
    applyParagraphStyle(editor, 'p', styles)
    expect(editor.state.doc.child(0).type.name).toBe('docParagraph')
    expect(editor.state.doc.child(0).attrs.styleId).toBe('Normal')
    editor.destroy()
  })
})

describe('Word\u2019s 50% rule for direct character formatting', () => {
  const bold = { type: 'bold' }
  const words = (...runs: Array<[string, unknown[]]>) => ({
    type: 'docParagraph',
    content: runs.map(([text, marks]) => ({ type: 'text', text, marks })),
  })
  function editorWith(paragraph: unknown): Editor {
    return new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [paragraph] } as never,
    })
  }
  const textStyleOf = (editor: Editor, index: number) =>
    editor.state.doc
      .child(0)
      .child(index)
      .marks.find((m) => m.type.name === 'docTextStyle')?.attrs

  it('keeps a bold word and a red word that cover less than half of the paragraph', () => {
    const editor = editorWith(words(['one ', [bold]], ['two ', [styled]], ['three four five', []]))
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 2)))
    applyParagraphStyleId(editor, 'Title', null)
    const block = editor.state.doc.child(0)
    expect(block.attrs.styleId).toBe('Title')
    expect(block.child(0).marks.some((m) => m.type.name === 'bold')).toBe(true)
    expect(textStyleOf(editor, 1)?.color).toBe('FF0000')
    editor.destroy()
  })

  it('drops a property from the whole paragraph once it covers more than half of the text', () => {
    const editor = editorWith(
      words(['mostly bold text ', [bold, styled]], ['and italic', [bold, { type: 'italic' }]]),
    )
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 2)))
    applyParagraphStyleId(editor, 'Heading1', 1)
    const block = editor.state.doc.child(0)
    expect(block.type.name).toBe('docHeading')
    expect(block.attrs.styleId).toBe('Heading1')
    for (let i = 0; i < block.childCount; i++)
      expect(block.child(i).marks.some((m) => m.type.name === 'bold')).toBe(false)
    // size/color covered 17 of 27 characters: gone; italic covered 10: kept
    expect(textStyleOf(editor, 0)).toBeUndefined()
    expect(block.lastChild!.marks.some((m) => m.type.name === 'italic')).toBe(true)
    editor.destroy()
  })

  it('exactly half is not a majority', () => {
    const editor = editorWith(words(['abcd', [bold]], ['efgh', []]))
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 2)))
    applyParagraphStyleId(editor, 'Quote', null)
    expect(
      editor.state.doc
        .child(0)
        .child(0)
        .marks.some((m) => m.type.name === 'bold'),
    ).toBe(true)
    editor.destroy()
  })

  it('clears only the majority attribute of a docTextStyle mark and keeps the rest', () => {
    const editor = editorWith(
      words(
        ['big ', [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 48, highlight: 'yellow' } }]],
        ['big too ', [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 48 } }]],
        ['end', []],
      ),
    )
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 2)))
    applyParagraphStyleId(editor, 'Normal', null)
    expect(textStyleOf(editor, 0)).toMatchObject({ sizeHalfPoints: null, highlight: 'yellow' })
    expect(textStyleOf(editor, 1)).toBeUndefined()
    editor.destroy()
  })

  it('applies to every paragraph the selection touches and keeps list numbering', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'first' }] },
          {
            type: 'docListItem',
            attrs: { kind: 'bullet', numId: '1', ilvl: 0 },
            content: [{ type: 'text', text: 'second' }],
          },
        ],
      },
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, 2, editor.state.doc.content.size - 2),
      ),
    )
    applyParagraphStyleId(editor, 'Subtitle', null)
    expect(editor.state.doc.child(0).attrs.styleId).toBe('Subtitle')
    expect(editor.state.doc.child(1).type.name).toBe('docListItem')
    expect(editor.state.doc.child(1).attrs.styleId).toBe('Subtitle')
    applyParagraphStyleId(editor, 'Heading2', 2)
    expect(editor.state.doc.child(1).type.name).toBe('docHeading')
    expect(editor.state.doc.child(1).attrs.level).toBe(2)
    editor.destroy()
  })
})

describe('line spacing via setParaAttrs (⌘1/⌘2/⌘5 path)', () => {
  it('sets the multiple on every paragraph in the selection and clears exact rules', () => {
    const editor = makeEditor()
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3, 20)),
    )
    setParaAttrs(editor, { lineSpacing: 1.5, lineRule: null, lineRawTwips: null })
    expect(editor.state.doc.child(0).attrs.lineSpacing).toBe(1.5)
    expect(editor.state.doc.child(1).attrs.lineSpacing).toBe(1.5)
    expect(editor.state.doc.child(0).attrs.lineRule).toBeNull()
    editor.destroy()
  })
})
