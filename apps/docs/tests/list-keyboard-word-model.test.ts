/**
 * Word's list keys: Backspace at the start of an item removes the marker, then
 * the indent, then joins; Tab changes the level only from the item start;
 * Enter on an empty item climbs one level per press and finally leaves the
 * list as an unindented Normal paragraph.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import type { NumberingDef } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const para = (text: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docParagraph',
  attrs,
  ...(text ? { content: [{ type: 'text', text }] } : {}),
})
const item = (text: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docListItem',
  attrs: { kind: 'bullet', numId: '3', ilvl: 0, ...attrs },
  ...(text ? { content: [{ type: 'text', text }] } : {}),
})

const def: NumberingDef = {
  numId: '3',
  abstractNumId: '0',
  levels: {
    0: { numFmt: 'bullet', lvlText: '•', start: 1, indentLeft: 720, hanging: 360 },
    1: { numFmt: 'bullet', lvlText: 'o', start: 1, indentLeft: 1440, hanging: 360 },
    2: { numFmt: 'bullet', lvlText: '▪', start: 1, indentLeft: 2160, hanging: 360 },
  },
  startOverrides: {},
}

function createEditor(content: JsonNode[], withDef = true): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content } as never,
  })
  if (withDef) editor.storage.listNumbering.defs = new Map([['3', def]])
  return editor
}

function press(editor: Editor, key: string, shiftKey = false) {
  const { view } = editor
  const event = new KeyboardEvent('keydown', { key, shiftKey, bubbles: true })
  return view.someProp('handleKeyDown', (f) => f(view, event))
}

const blocks = (editor: Editor) =>
  editor.state.doc.content.content.map((n) => ({
    type: n.type.name,
    text: n.textContent,
    indentLeft: n.attrs.indentLeft,
    indentFirstLine: n.attrs.indentFirstLine,
    ilvl: n.attrs.ilvl,
  }))

/** caret at the start of the block with this index */
const caretAtStart = (editor: Editor, index: number) => {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  editor.commands.setTextSelection(pos + 1)
}

describe('Backspace at the start of a list item', () => {
  it('removes the marker, then the indent, then joins', () => {
    const editor = createEditor([para('above'), item('entry', { spaceAfter: 120 })])
    caretAtStart(editor, 1)

    expect(press(editor, 'Backspace')).toBe(true)
    expect(blocks(editor)[1]).toMatchObject({
      type: 'docParagraph',
      text: 'entry',
      indentLeft: 720,
    })
    const plain = editor.state.doc.child(1)
    expect(plain.attrs.spaceAfter).toBe(120)
    expect(plain.attrs.indentFirstLine).toBeNull()
    expect(plain.attrs.docxIndex).toBeNull()

    expect(press(editor, 'Backspace')).toBe(true)
    expect(blocks(editor)[1]).toMatchObject({
      type: 'docParagraph',
      text: 'entry',
      indentLeft: null,
    })

    expect(press(editor, 'Backspace')).toBe(true)
    expect(blocks(editor)).toEqual([
      {
        type: 'docParagraph',
        text: 'aboveentry',
        indentLeft: null,
        indentFirstLine: null,
        ilvl: undefined,
      },
    ])
    editor.destroy()
  })

  it('keeps the indent of the level the item was on', () => {
    const editor = createEditor([item('deep', { ilvl: 2 })])
    caretAtStart(editor, 0)
    press(editor, 'Backspace')
    expect(blocks(editor)[0]).toMatchObject({ type: 'docParagraph', indentLeft: 2160 })
    editor.destroy()
  })

  it('an item with its own w:ind keeps that position', () => {
    const editor = createEditor([item('own', { indentLeft: 1000, indentFirstLine: -300 })])
    caretAtStart(editor, 0)
    press(editor, 'Backspace')
    expect(blocks(editor)[0]).toMatchObject({
      type: 'docParagraph',
      indentLeft: 1000,
      indentFirstLine: null,
    })
    editor.destroy()
  })

  it('an explicit zero indent stays when the style has an indent of its own', () => {
    const editor = createEditor([item('flush', { indentLeft: 0, styleId: 'ListParagraph' })])
    editor.storage.listNumbering.styles = new Map([
      [
        'ListParagraph',
        {
          styleId: 'ListParagraph',
          name: 'List Paragraph',
          type: 'paragraph',
          display: { indentLeftTwips: 720 },
        },
      ],
    ])
    caretAtStart(editor, 0)
    press(editor, 'Backspace')
    expect(blocks(editor)[0]).toMatchObject({ type: 'docParagraph', indentLeft: 0 })
    editor.destroy()
  })

  it('falls back to the rendered default when the list has no definition', () => {
    const editor = createEditor([item('loose', { numId: null, ilvl: 1 })], false)
    caretAtStart(editor, 0)
    press(editor, 'Backspace')
    expect(blocks(editor)[0]).toMatchObject({ type: 'docParagraph', indentLeft: 792 + 432 })
    editor.destroy()
  })

  it('also works on an empty item', () => {
    const editor = createEditor([para('above'), item('')])
    caretAtStart(editor, 1)
    press(editor, 'Backspace')
    expect(blocks(editor)).toHaveLength(2)
    expect(blocks(editor)[1]).toMatchObject({ type: 'docParagraph', indentLeft: 720 })
    editor.destroy()
  })

  it('in the middle of the text or with a selection is left to the default handling', () => {
    const editor = createEditor([item('entry')])
    // a single character mid-text is the browser's own delete: no handler claims it
    editor.commands.setTextSelection(3)
    expect(press(editor, 'Backspace')).toBeFalsy()
    expect(blocks(editor)[0]).toMatchObject({ type: 'docListItem', text: 'entry' })

    editor.commands.setTextSelection({ from: 1, to: 3 })
    press(editor, 'Backspace')
    expect(blocks(editor)[0]).toMatchObject({ type: 'docListItem', text: 'try' })
    editor.destroy()
  })

  it('a plain paragraph without indent still joins the previous one', () => {
    const editor = createEditor([para('above'), para('below')])
    caretAtStart(editor, 1)
    press(editor, 'Backspace')
    expect(blocks(editor).map((b) => b.text)).toEqual(['abovebelow'])
    editor.destroy()
  })

  it('a first-line indent alone is not a left indent: the paragraph joins', () => {
    const editor = createEditor([para('above'), para('below', { indentFirstLine: 420 })])
    caretAtStart(editor, 1)
    press(editor, 'Backspace')
    expect(blocks(editor).map((b) => b.text)).toEqual(['abovebelow'])
    editor.destroy()
  })

  it('a hanging indent leaves together with the left indent', () => {
    const editor = createEditor([
      para('above'),
      para('ref', { indentLeft: 720, indentFirstLine: -720 }),
    ])
    caretAtStart(editor, 1)
    press(editor, 'Backspace')
    expect(blocks(editor)[1]).toMatchObject({
      text: 'ref',
      indentLeft: null,
      indentFirstLine: null,
    })
    editor.destroy()
  })
})

describe('Tab in a list item', () => {
  it('changes the level only at the start of the item', () => {
    const editor = createEditor([item('entry', { ilvl: 1 })])
    caretAtStart(editor, 0)
    press(editor, 'Tab')
    expect(blocks(editor)[0].ilvl).toBe(2)
    press(editor, 'Tab', true)
    press(editor, 'Tab', true)
    expect(blocks(editor)[0].ilvl).toBe(0)

    editor.commands.setTextSelection(3)
    expect(press(editor, 'Tab')).toBe(true)
    expect(blocks(editor)[0]).toMatchObject({ text: 'en\ttry', ilvl: 0 })
    expect(press(editor, 'Tab', true)).toBe(true)
    expect(blocks(editor)[0]).toMatchObject({ text: 'en\ttry', ilvl: 0 })
    editor.destroy()
  })

  it('a selection across items changes every item', () => {
    const editor = createEditor([item('one'), item('two')])
    editor.commands.setTextSelection({ from: 2, to: 8 })
    press(editor, 'Tab')
    expect(blocks(editor).map((b) => b.ilvl)).toEqual([1, 1])
    editor.destroy()
  })
})

describe('Enter on an empty list item', () => {
  it('climbs one level per press, then leaves as an unindented Normal paragraph', () => {
    const editor = createEditor([
      item('', { ilvl: 2, styleId: 'ListParagraph', spaceAfter: 160, indentLeft: 2000 }),
    ])
    caretAtStart(editor, 0)
    press(editor, 'Enter')
    expect(blocks(editor)[0]).toMatchObject({ type: 'docListItem', ilvl: 1 })
    press(editor, 'Enter')
    expect(blocks(editor)[0]).toMatchObject({ type: 'docListItem', ilvl: 0 })
    press(editor, 'Enter')
    expect(blocks(editor)).toEqual([
      { type: 'docParagraph', text: '', indentLeft: null, indentFirstLine: null, ilvl: undefined },
    ])
    const left = editor.state.doc.child(0)
    expect(left.attrs.styleId).toBeNull()
    expect(left.attrs.spaceAfter).toBe(160)
    expect(left.attrs.numId).toBeUndefined()
    editor.destroy()
  })

  it('a non-empty item still splits into a sibling', () => {
    const editor = createEditor([item('entry', { ilvl: 1 })])
    editor.commands.setTextSelection(6)
    press(editor, 'Enter')
    expect(blocks(editor).map((b) => [b.type, b.ilvl])).toEqual([
      ['docListItem', 1],
      ['docListItem', 1],
    ])
    editor.destroy()
  })
})
