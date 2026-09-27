/**
 * Toggle Field Codes is a display-time view: the cached result stays in the
 * document while decorations hide it and show `{ INSTR }`; ⌥F9 flips every
 * field and per-field toggles are exceptions to it. Edit Field… rewrites the
 * instruction and result in place, keeping the run's other marks.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  collectFields,
  fieldCodeText,
  fieldCodesPluginKey,
  fieldRangeAt,
  isFieldCodeShown,
  setFieldInstr,
  toggleAllFieldCodes,
  toggleFieldCodes,
} from '../src/renderer/editor/field-codes'

const live: Editor[] = []
afterEach(() => {
  for (const e of live.splice(0)) e.destroy()
})

const field = (text: string, instr: string, extra: Array<{ type: string }> = []) => ({
  type: 'text',
  text,
  marks: [...extra, { type: 'instrField', attrs: { instr } }],
})

function makeEditor() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'Page ' },
            field('3', 'PAGE'),
            { type: 'text', text: ' of ' },
            field('1', 'NUMPAGES', [{ type: 'bold' }]),
            field('2', 'NUMPAGES'),
          ],
        },
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'Dated ' }, field('2026-01-01', 'DATE \\@ "yyyy-MM-dd"')],
        },
      ],
    },
  })
  live.push(editor)
  return editor
}

const codeWidgets = (editor: Editor) =>
  fieldCodesPluginKey
    .getState(editor.state)!
    .decos.find(undefined, undefined, (spec) => typeof spec.key === 'string')
    .map((d) => (d.spec.key as string).split(':').slice(2).join(':'))

describe('field ranges', () => {
  it('finds the field around a position, across runs with different formatting', () => {
    const editor = makeEditor()
    // "Page " 1..6, PAGE 6..7, " of " 7..11, NUMPAGES 11..13
    expect(fieldRangeAt(editor.state, 6)).toMatchObject({ from: 6, to: 7, instr: 'PAGE' })
    expect(fieldRangeAt(editor.state, 7)).toMatchObject({ from: 6, to: 7, instr: 'PAGE' })
    expect(fieldRangeAt(editor.state, 12)).toMatchObject({ from: 11, to: 13, instr: 'NUMPAGES' })
    expect(fieldRangeAt(editor.state, 9)).toBeNull()
    expect(collectFields(editor.state.doc).map((f) => f.instr)).toEqual([
      'PAGE',
      'NUMPAGES',
      'DATE \\@ "yyyy-MM-dd"',
    ])
  })

  it('formats the code view like Word', () => {
    expect(fieldCodeText('  DATE \\@ "d MMMM yyyy" ')).toBe('{ DATE \\@ "d MMMM yyyy" }')
  })
})

describe('toggle field codes', () => {
  it('shows one field as code without touching the document or history', () => {
    const editor = makeEditor()
    const before = editor.state.doc
    const page = fieldRangeAt(editor.state, 6)!
    toggleFieldCodes(editor.view, page)
    expect(editor.state.doc.eq(before)).toBe(true)
    expect(isFieldCodeShown(editor.state, page)).toBe(true)
    expect(codeWidgets(editor)).toEqual(['{ PAGE }'])
    expect(editor.can().undo()).toBe(false)
    toggleFieldCodes(editor.view, page)
    expect(isFieldCodeShown(editor.state, page)).toBe(false)
    expect(codeWidgets(editor)).toEqual([])
  })

  it('⌥F9 flips every field; a per-field toggle then hides that one again', () => {
    const editor = makeEditor()
    toggleAllFieldCodes(editor.view)
    expect(codeWidgets(editor)).toEqual(['{ PAGE }', '{ NUMPAGES }', '{ DATE \\@ "yyyy-MM-dd" }'])
    const page = fieldRangeAt(editor.state, 6)!
    toggleFieldCodes(editor.view, page)
    expect(isFieldCodeShown(editor.state, page)).toBe(false)
    expect(codeWidgets(editor)).toEqual(['{ NUMPAGES }', '{ DATE \\@ "yyyy-MM-dd" }'])
    toggleAllFieldCodes(editor.view)
    expect(codeWidgets(editor)).toEqual([])
  })

  it('keeps a toggled field through edits before it', () => {
    const editor = makeEditor()
    const date = fieldRangeAt(editor.state, 22)!
    toggleFieldCodes(editor.view, date)
    editor.commands.insertContentAt(1, 'First ')
    const moved = fieldRangeAt(editor.state, date.from + 6)!
    expect(moved.instr).toBe('DATE \\@ "yyyy-MM-dd"')
    expect(isFieldCodeShown(editor.state, moved)).toBe(true)
    expect(codeWidgets(editor)).toEqual(['{ DATE \\@ "yyyy-MM-dd" }'])
  })
})

describe('toggle field codes', () => {
  it('drops the code view when the toggled field is deleted', () => {
    const editor = makeEditor()
    const page = fieldRangeAt(editor.state, 6)!
    toggleFieldCodes(editor.view, page)
    expect(codeWidgets(editor)).toEqual(['{ PAGE }'])
    editor.commands.deleteRange({ from: page.from, to: page.to })
    expect(codeWidgets(editor)).toEqual([])
    expect(fieldCodesPluginKey.getState(editor.state)!.decos.find()).toEqual([])
  })
})

describe('setFieldInstr', () => {
  it('rewrites instruction and result, keeps other marks, flags the field dirty', () => {
    const editor = makeEditor()
    const numpages = fieldRangeAt(editor.state, 12)!
    expect(setFieldInstr(editor, numpages, 'SECTIONPAGES', '9')).toBe(true)
    expect(editor.state.doc.textContent).toBe('Page 3 of 9Dated 2026-01-01')
    const run = editor.state.doc.nodeAt(11)!
    const mark = run.marks.find((m) => m.type.name === 'instrField')!
    expect(mark.attrs).toMatchObject({ instr: 'SECTIONPAGES', dirty: true })
    expect(run.marks.some((m) => m.type.name === 'bold')).toBe(true)
  })

  it('keeps the old result when no value can be computed', () => {
    const editor = makeEditor()
    const page = fieldRangeAt(editor.state, 6)!
    expect(setFieldInstr(editor, page, 'REF _Ref1 \\h', '')).toBe(true)
    expect(fieldRangeAt(editor.state, 6)).toMatchObject({ instr: 'REF _Ref1 \\h' })
    expect(editor.state.doc.textBetween(6, 7)).toBe('3')
    expect(setFieldInstr(editor, page, '   ', 'x')).toBe(false)
  })
})
