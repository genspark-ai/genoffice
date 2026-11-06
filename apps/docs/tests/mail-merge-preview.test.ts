// Regression: the merged document must survive `editor.setContent` —
// PM rejects empty text nodes, which used to void the whole document
// and leave the editor blank after clicking Preview.
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { PmMark, PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { mailMergePreviewActive, originalTemplateDoc } from '../src/renderer/mail-merge/template'
import {
  previewTemplate,
  setPreviewReadOnly,
  setPreviewTemplate,
} from '../src/renderer/mail-merge/preview-extension'
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

describe('the preview storage holds the original template', () => {
  it('readers see the template with its field codes while the preview shows values', () => {
    const template = doc(para(text('Dear '), text('«FirstName»', [mergeMark('FirstName')])))
    const editor = makeEditor(template)

    setPreviewTemplate(editor, template)
    expect(previewTemplate(editor)).not.toBeNull()
    expect(mailMergePreviewActive(editor)).toBe(true)

    // the editor now shows the merged values, but the readers must not
    editor.commands.setContent(substituteMergeFields(template, { FirstName: 'Ada' }) as never, {
      emitUpdate: false,
    })
    expect(editor.getText()).toContain('Ada')
    const seen = JSON.stringify(originalTemplateDoc(editor))
    expect(seen).toContain('«FirstName»')
    expect(seen).not.toContain('Ada')

    // handing the document back ends the preview
    setPreviewTemplate(editor, null)
    expect(mailMergePreviewActive(editor)).toBe(false)
    expect(JSON.stringify(originalTemplateDoc(editor))).toContain('Ada')
    editor.destroy()
  })

  it('the storage lives on the extension slot, not at the top level', () => {
    const template = doc(para(text('Dear ')))
    const editor = makeEditor(template)
    setPreviewTemplate(editor, template)
    // TipTap keys an extension's storage by the extension name: a reader that
    // looks at editor.storage.mailmergePreviewJson finds nothing and would
    // silently hand back the preview document instead
    expect(editor.storage.mailmergePreview.mailmergePreviewJson).not.toBeNull()
    expect(
      (editor.storage as unknown as Record<string, unknown>).mailmergePreviewJson,
    ).toBeUndefined()
    editor.destroy()
  })

  it('the read-only switch stops the view from accepting input', () => {
    const template = doc(para(text('Dear ')))
    const editor = makeEditor(template)
    setPreviewReadOnly(editor, true)
    expect(editor.isEditable).toBe(false)
    expect(editor.view.editable).toBe(false)
    setPreviewReadOnly(editor, false)
    expect(editor.isEditable).toBe(true)
    expect(editor.view.editable).toBe(true)
    editor.destroy()
  })
})
