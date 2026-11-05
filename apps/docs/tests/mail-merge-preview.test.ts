// Regression: the merged document must survive `editor.setContent` —
// PM rejects empty text nodes, which used to void the whole document
// and leave the editor blank after clicking Preview.
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { PmMark, PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { substituteMergeFields } from '../src/renderer/mail-merge/template'

const mark = (instr: string): PmMark => ({ type: 'instrField', attrs: { instr } })
const mergeMark = (name: string): PmMark => mark(`MERGEFIELD "${name}" \\* MERGEFORMAT`)
const text = (value: string, marks?: PmMark[]): PmNode => ({ type: 'text', text: value, marks })
const para = (...content: PmNode[]): PmNode => ({ type: 'docParagraph', content })
const doc = (...content: PmNode[]): PmNode => ({ type: 'doc', content })

function makeEditor(content: PmNode): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: content as never,
  })
}

describe('preview keeps the editor filled', () => {
  it('split-field merged doc round-trips through setContent', () => {
    const template = doc(
      para(text('«FirstName»', [mergeMark('FirstName')]), text(' tail', [mergeMark('FirstName')])),
    )
    const merged = substituteMergeFields(template, { FirstName: 'Ada' })
    const editor = makeEditor(template)
    editor.commands.setContent(merged as never, { emitUpdate: false })
    expect(editor.getText()).toContain('Ada')
    editor.destroy()
  })

  it('empty value merged doc round-trips through setContent', () => {
    const template = doc(para(text('Dear '), text('«FirstName»', [mergeMark('FirstName')])))
    const merged = substituteMergeFields(template, { FirstName: '' })
    const editor = makeEditor(template)
    editor.commands.setContent(merged as never, { emitUpdate: false })
    expect(editor.getText()).toContain('Dear')
    editor.destroy()
  })

  it('plain merged doc round-trips through setContent', () => {
    const template = doc(
      para(text('Dear '), text('«FirstName»', [mergeMark('FirstName')]), text('!')),
    )
    const merged = substituteMergeFields(template, { FirstName: 'Ada' })
    const editor = makeEditor(template)
    editor.commands.setContent(merged as never, { emitUpdate: false })
    expect(editor.getText()).toContain('Dear Ada!')
    editor.destroy()
  })
})
