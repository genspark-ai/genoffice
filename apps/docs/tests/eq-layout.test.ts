/** Enclose Characters and Combine Characters: put on selected text, changed, taken off, and saved as their field. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { DOMParser as PmDOMParser, DOMSerializer } from '@tiptap/pm/model'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  applyEqLayout,
  combineLines,
  eqLayoutFontOf,
  eqLayoutTargetOf,
  eqLayoutText,
} from '../src/renderer/editor/eq-layout'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { leafTextForWordCount } from '../src/renderer/word-count'

describe('enclosed and combined characters', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'see ' },
              { type: 'text', text: 'ABCDEF', marks: [{ type: 'bold' }] },
              { type: 'text', text: ' rise' },
            ],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'second' }] },
        ],
      },
    })
  })
  afterEach(() => editor.destroy())

  const layouts = () => {
    const out: Array<Record<string, unknown>> = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docEqLayout')
        out.push({ ...node.attrs, bold: node.marks.some((m) => m.type.name === 'bold') })
    })
    return out
  }

  it('splits combined characters over two lines, the upper one first', () => {
    expect(combineLines('abcdef')).toEqual({ a: 'abc', b: 'def' })
    expect(combineLines('abcde')).toEqual({ a: 'abc', b: 'de' })
    expect(combineLines('a')).toEqual({ a: 'a', b: '' })
    expect(eqLayoutText({ kind: 'combine', a: 'abc', b: 'de' })).toBe('abcde')
    expect(eqLayoutText({ kind: 'enclose', a: '○', b: 'A' })).toBe('A')
  })

  it('takes a caret or a short stretch of text in one paragraph', () => {
    editor.commands.setTextSelection({ from: 5, to: 6 })
    expect(eqLayoutTargetOf(editor.state, 'enclose')).toEqual({
      from: 5,
      to: 6,
      text: 'A',
      existing: null,
    })
    // a caret: the dialog asks for the text
    editor.commands.setTextSelection(5)
    expect(eqLayoutTargetOf(editor.state, 'enclose')?.text).toBe('')
    // more than the layout holds, and a selection into the next paragraph
    editor.commands.setTextSelection({ from: 5, to: 8 })
    expect(eqLayoutTargetOf(editor.state, 'enclose')).toBeNull()
    expect(eqLayoutTargetOf(editor.state, 'combine')?.text).toBe('ABC')
    editor.commands.setTextSelection({ from: 5, to: 20 })
    expect(eqLayoutTargetOf(editor.state, 'combine')).toBeNull()
  })

  it('encloses the selected character in its formatting, changes it, and takes the enclosure off', () => {
    editor.commands.setTextSelection({ from: 5, to: 6 })
    const target = eqLayoutTargetOf(editor.state, 'enclose')!
    applyEqLayout(editor, target, {
      kind: 'enclose',
      text: 'A',
      enclosure: '○',
      style: 'shrink',
    })
    expect(layouts()).toEqual([
      { kind: 'enclose', a: '○', b: 'A', style: 'shrink', xml: '', bold: true },
    ])
    expect(editor.state.doc.textContent).toBe('see BCDEF risesecond')

    // it is selected: the dialog opens on it again, and not as combined characters
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    const again = eqLayoutTargetOf(editor.state, 'enclose')!
    expect(again.existing).toEqual({ kind: 'enclose', a: '○', b: 'A', style: 'shrink' })
    expect(eqLayoutTargetOf(editor.state, 'combine')).toBeNull()
    applyEqLayout(editor, again, {
      kind: 'enclose',
      text: 'A',
      enclosure: '□',
      style: 'enlarge',
    })
    expect(layouts()).toMatchObject([{ a: '□', b: 'A', style: 'enlarge' }])

    applyEqLayout(editor, eqLayoutTargetOf(editor.state, 'enclose')!, null)
    expect(layouts()).toEqual([])
    expect(editor.state.doc.textContent).toBe('see ABCDEF risesecond')
  })

  it('combines the selected characters and counts them as the characters they are', () => {
    editor.commands.setTextSelection({ from: 5, to: 11 })
    applyEqLayout(editor, eqLayoutTargetOf(editor.state, 'combine')!, {
      kind: 'combine',
      text: 'ABCDEF',
    })
    expect(layouts()).toMatchObject([{ kind: 'combine', a: 'ABC', b: 'DEF', bold: true }])
    const node = (editor.state.selection as NodeSelection).node
    expect(leafTextForWordCount(node)).toBe('ABCDEF')
  })

  it("takes the font and size picked for it, and keeps the text's own when none is", () => {
    editor.commands.setTextSelection({ from: 5, to: 9 })
    applyEqLayout(
      editor,
      eqLayoutTargetOf(editor.state, 'combine')!,
      { kind: 'combine', text: 'ABCD' },
      { font: 'Arial', sizePt: 18 },
    )
    const node = editor.state.doc.nodeAt(5)!
    expect(node.type.name).toBe('docEqLayout')
    expect(node.marks.some((m) => m.type.name === 'bold')).toBe(true)
    const style = node.marks.find((m) => m.type.name === 'docTextStyle')!
    expect(style.attrs).toMatchObject({ fontAscii: 'Arial', sizeHalfPoints: 36 })
    expect(eqLayoutFontOf(editor.state, eqLayoutTargetOf(editor.state, 'combine')!)).toEqual({
      font: 'Arial',
      sizePt: 18,
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, [])
    const first = plan.saveBlocks[0]
    const runs = first.kind === 'generated' ? (first.block.runs ?? []) : []
    expect(runs[1]).toMatchObject({ sizeHalfPoints: 36, eqLayout: { kind: 'combine' } })

    // opened again with nothing picked: the size stays
    applyEqLayout(editor, eqLayoutTargetOf(editor.state, 'combine')!, {
      kind: 'combine',
      text: 'ABC',
    })
    expect(
      editor.state.doc.nodeAt(5)!.marks.find((m) => m.type.name === 'docTextStyle')!.attrs
        .sizeHalfPoints,
    ).toBe(36)
  })

  it('comes back from the clipboard markup as what it was', () => {
    editor.commands.setTextSelection({ from: 5, to: 6 })
    applyEqLayout(editor, eqLayoutTargetOf(editor.state, 'enclose')!, {
      kind: 'enclose',
      text: 'A',
      enclosure: '\u25a1',
      style: 'enlarge',
    })
    const { dom } = editor.view.serializeForClipboard(editor.state.selection.content())
    const parsed = PmDOMParser.fromSchema(editor.schema).parse(dom)
    const pasted: Array<Record<string, unknown>> = []
    parsed.descendants((node) => {
      if (node.type.name === 'docEqLayout') pasted.push(node.attrs)
    })
    expect(pasted).toEqual([{ kind: 'enclose', a: '\u25a1', b: 'A', style: 'enlarge', xml: '' }])

    // a field read from a file keeps its markup through the copy
    const node = editor.schema.nodes.docEqLayout.create({
      kind: 'combine',
      a: 'ab',
      b: 'cd',
      xml: '<w:r><w:instrText> eq \\o(\\s\\up 5(ab),\\s\\do 2(cd))</w:instrText></w:r>',
    })
    const html = DOMSerializer.fromSchema(editor.schema).serializeNode(node) as HTMLElement
    const holder = document.createElement('p')
    holder.appendChild(html)
    const back = PmDOMParser.fromSchema(editor.schema).parse(holder).firstChild?.firstChild
    expect(back?.attrs).toEqual(node.attrs)
  })

  it('saves as the run that writes the field, and reads a file back as the same node', async () => {
    editor.commands.setTextSelection({ from: 5, to: 6 })
    applyEqLayout(editor, eqLayoutTargetOf(editor.state, 'enclose')!, {
      kind: 'enclose',
      text: 'A',
      enclosure: '○',
      style: 'enlarge',
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, [])
    const first = plan.saveBlocks[0]
    const runs = first.kind === 'generated' ? (first.block.runs ?? []) : []
    expect(runs.map((r) => r.text)).toEqual(['see ', 'A', 'BCDEF', ' rise'])
    expect(runs[1]).toMatchObject({
      bold: true,
      eqLayout: { kind: 'enclose', a: '○', b: 'A', style: 'enlarge' },
    })

    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> eq \\o\\ac(○,B)</w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: `<w:p>${field}</w:p>` }))
    const doc = blocksToPmDoc(parsed.blocks) as PmNode
    const node = doc.content?.[0]?.content?.[0]
    expect(node?.type).toBe('docEqLayout')
    expect(node?.attrs).toMatchObject({ kind: 'enclose', a: '○', b: 'B' })
    expect(String(node?.attrs?.xml)).toContain('w:instrText')
  })
})
