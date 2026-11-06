// Regression: the merged document must survive `editor.setContent` —
// PM rejects empty text nodes, which used to void the whole document
// and leave the editor blank after clicking Preview.
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { PmMark, PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { mailMergePreviewActive, originalTemplateDoc } from '../src/renderer/mail-merge/template'
import {
  MAIL_MERGE_PREVIEW_BYPASS,
  previewTemplate,
  setPreviewLocked,
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

describe('the App-level preview lock fences every edit', () => {
  it('refuses doc-changing transactions from any entry point', () => {
    const editor = makeEditor(doc(para(text('Dear '))))
    // an ordinary, undoable edit first — undo during the preview must not
    // remove it
    editor.view.dispatch(editor.state.tr.insertText('Z'))
    setPreviewLocked(editor, true)

    // the view: what typing, keyboard shortcuts and the context menu dispatch
    editor.view.dispatch(editor.state.tr.insertText('X', 0))
    expect(editor.state.doc.textContent).not.toContain('X')

    // a TipTap command chain: what the ribbon commands and the AI tools run
    editor.commands.insertContent('Y')
    expect(editor.state.doc.textContent).not.toContain('Y')

    // undo replays recorded steps as a new doc-changing transaction
    editor.commands.undo()
    expect(editor.state.doc.textContent).toContain('Z')

    // off the lock again, the same dispatch goes through
    setPreviewLocked(editor, false)
    editor.view.dispatch(editor.state.tr.insertText('X', 0))
    expect(editor.state.doc.textContent).toContain('X')
    editor.destroy()
  })

  it('keeps selection-only transactions coming through', () => {
    const editor = makeEditor(doc(para(text('Dear '))))
    setPreviewLocked(editor, true)
    // walking the preview document (clicking into it, scrolling to a field)
    // stays possible: only the doc itself is frozen
    expect(editor.commands.setTextSelection(1)).toBe(true)
    expect(editor.state.selection.from).toBe(1)
    expect(editor.state.doc.textContent).toBe('Dear ')
    editor.destroy()
  })

  it("lets the preview's own swaps through on the bypass meta", () => {
    const editor = makeEditor(doc(para(text('Dear '))))
    setPreviewLocked(editor, true)
    editor
      .chain()
      .setMeta(MAIL_MERGE_PREVIEW_BYPASS, true)
      .setContent(doc(para(text('Merged'))) as never, { emitUpdate: false })
      .run()
    expect(editor.getText()).toContain('Merged')
    editor.destroy()
  })
})

describe('the read-only switch keeps the pre-preview state', () => {
  it('does not lift Read Mode when the preview ends', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: doc(para(text('Dear '))) as never,
      editable: false,
    })
    setPreviewReadOnly(editor, true)
    expect(editor.isEditable).toBe(false)
    setPreviewReadOnly(editor, false)
    // the document was read-only before the preview: it stays that way
    expect(editor.isEditable).toBe(false)
    expect(editor.view.editable).toBe(false)
    editor.destroy()
  })
})
