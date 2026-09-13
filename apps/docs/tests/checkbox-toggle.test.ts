/**
 * Clicking a checkbox glyph flips it (genoffice#248): legacy FORMCHECKBOX
 * fields and w14:checkbox content controls toggle, plain box characters do not.
 */
import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { toggleCheckboxAt } from '../src/renderer/editor/checkbox-toggle'
import { editorExtensions } from '../src/renderer/editor/extensions'

const SDT_PR =
  '<w:sdtPr><w14:checkbox><w14:checked w14:val="0"/>' +
  '<w14:checkedState w14:val="2611" w14:font="MS Gothic"/>' +
  '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>'

// destroyed after each test: ProseMirror's DOM observer flushes on a timer
// and would touch `document` after the jsdom environment is torn down
const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

const editorWith = (content: Array<Record<string, unknown>>) => {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content }],
    } as never,
  })
  editors.push(editor)
  return editor
}

const text = (editor: Editor) => editor.state.doc.textContent

describe('toggleCheckboxAt', () => {
  it('flips a FORMCHECKBOX glyph on either side of the caret slot', () => {
    const editor = editorWith([
      { type: 'text', text: 'item: ' },
      {
        type: 'text',
        text: '☐',
        marks: [{ type: 'instrField', attrs: { instr: 'FORMCHECKBOX', beginXml: '<w:r/>' } }],
      },
    ])
    // glyph sits at doc position 7 (paragraph open + 6 characters)
    let tr = toggleCheckboxAt(editor.state, 7, 'after')
    expect(tr).not.toBeNull()
    editor.view.dispatch(tr!)
    expect(text(editor)).toBe('item: ☒')
    tr = toggleCheckboxAt(editor.state, 8, 'before')
    editor.view.dispatch(tr!)
    expect(text(editor)).toBe('item: ☐')
    const glyph = editor.state.doc.nodeAt(7)!
    expect(glyph.marks.map((m) => m.type.name)).toEqual(['instrField'])
  })

  it('flips a content control between the glyphs its properties declare', () => {
    const editor = editorWith([
      { type: 'text', text: '☐', marks: [{ type: 'ctrlCheckbox', attrs: { sdtPr: SDT_PR } }] },
    ])
    editor.view.dispatch(toggleCheckboxAt(editor.state, 1, 'after')!)
    expect(text(editor)).toBe('☑')
    expect(editor.state.doc.nodeAt(1)!.marks[0].attrs.sdtPr).toBe(SDT_PR)
    editor.view.dispatch(toggleCheckboxAt(editor.state, 2, 'before')!)
    expect(text(editor)).toBe('☐')
  })

  it('leaves a plain box character and other fields alone', () => {
    const editor = editorWith([
      { type: 'text', text: '☐ ' },
      { type: 'text', text: '3', marks: [{ type: 'instrField', attrs: { instr: 'PAGE' } }] },
    ])
    expect(toggleCheckboxAt(editor.state, 1, 'after')).toBeNull()
    expect(toggleCheckboxAt(editor.state, 3, 'after')).toBeNull()
    expect(toggleCheckboxAt(editor.state, 0, 'before')).toBeNull()
  })
})
