import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function openImageDoc() {
  const source = await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, source }
}

describe('image wrap in the editor', () => {
  it('keeps an untouched image byte-identical', async () => {
    const { editor, parsed, source } = await openImageDoc()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('position preset (margin align) round-trips and stays clean on reopen', async () => {
    const { editor, parsed } = await openImageDoc()
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    // Layout tab position gallery: bottom center + square wrap
    editor.commands.updateAttributes('docProtected', {
      imageWrap: 'square-left',
      imagePosH: 'center',
      imagePosV: 'bottom',
      imageOffsetXEmu: null,
      imageOffsetYEmu: null,
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].imagePosH).toBe('center')
    expect(reparsed.blocks[0].imagePosV).toBe('bottom')
    expect(reparsed.blocks[0].imageWrap).toBe('square-left')

    // reopen: same attrs come back from parse, so an untouched save stays byte-identical
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks) as never,
    })
    const plan2 = pmDocToSavePlan(editor2.getJSON() as PmNode, reparsed.blocks)
    expect(plan2.changedCount).toBe(0)
    expect(await saveDocx(reparsed, plan2.saveBlocks)).toEqual(saved)
    editor.destroy()
    editor2.destroy()
  })

  it('setting imageWrap floats the image and round-trips through save', async () => {
    const { editor, parsed } = await openImageDoc()
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    editor.commands.updateAttributes('docProtected', { imageWrap: 'square-right' })
    expect(editor.view.dom.querySelector('.doc-protected.img-wrap-square-right')).toBeTruthy()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].imageWrap).toBe('square-right')

    // back to inline from the reparsed doc
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks) as never,
    })
    editor2.view.dispatch(editor2.state.tr.setSelection(NodeSelection.create(editor2.state.doc, 0)))
    editor2.commands.updateAttributes('docProtected', { imageWrap: null })
    const plan2 = pmDocToSavePlan(editor2.getJSON() as PmNode, reparsed.blocks)
    expect(plan2.changedCount).toBe(1)
    const p3 = await parseDocx(await saveDocx(reparsed, plan2.saveBlocks))
    expect(p3.blocks[0].imageWrap).toBeUndefined()
    editor.destroy()
    editor2.destroy()
  })
})
