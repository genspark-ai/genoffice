import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function openImageDoc(bodyXml = IMAGE_PARAGRAPH_XML, extraRels?: string) {
  const source = await buildDocx({ bodyXml, withImage: true, extraRels })
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

  it('rotation + flip round-trip through a:xfrm and clear cleanly', async () => {
    // the minimal shared fixture has no pic:spPr; rotation lives on the pic xfrm
    const withXfrm = IMAGE_PARAGRAPH_XML.replace(
      '</pic:blipFill></pic:pic>',
      '</pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>',
    )
    const source = await buildDocx({ bodyXml: withXfrm, withImage: true })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    editor.commands.updateAttributes('docProtected', { imageRotDeg: 90, imageFlipH: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)

    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].imageRotDeg).toBe(90)
    expect(reparsed.blocks[0].imageFlipH).toBe(true)
    expect(reparsed.blocks[0].imageFlipV).toBeUndefined()

    // untouched re-save stays byte-identical; clearing removes the attributes
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks) as never,
    })
    const plan2 = pmDocToSavePlan(editor2.getJSON() as PmNode, reparsed.blocks)
    expect(plan2.changedCount).toBe(0)
    expect(await saveDocx(reparsed, plan2.saveBlocks)).toEqual(saved)

    editor2.view.dispatch(editor2.state.tr.setSelection(NodeSelection.create(editor2.state.doc, 0)))
    editor2.commands.updateAttributes('docProtected', { imageRotDeg: null, imageFlipH: false })
    const plan3 = pmDocToSavePlan(editor2.getJSON() as PmNode, reparsed.blocks)
    expect(plan3.changedCount).toBe(1)
    const cleared = await saveDocx(reparsed, plan3.saveBlocks)
    const p3 = await parseDocx(cleared)
    expect(p3.blocks[0].imageRotDeg).toBeUndefined()
    expect(p3.blocks[0].imageFlipH).toBeUndefined()
    expect(cleared.length).toBeGreaterThan(0)
    editor.destroy()
    editor2.destroy()
  })

  it('imageReplace swaps bytes in place, keeping the drawing XML and wrap', async () => {
    // 1x1 transparent GIF
    const GIF_B64 = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'
    const { editor, parsed } = await openImageDoc()
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    editor.commands.updateAttributes('docProtected', {
      imageWrap: 'square-right',
      imageReplace: { base64: GIF_B64, mime: 'image/gif' },
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)

    const reparsed = await parseDocx(saved)
    const blk = reparsed.blocks[0]
    // still an original image block (docxIndex-anchored), wrap applied, new bytes served
    expect(blk.type).toBe('image')
    expect(blk.imageWrap).toBe('square-right')
    expect(blk.imageDataUrl?.startsWith('data:image/gif;base64,')).toBe(true)

    // an untouched re-save of the reparsed doc stays byte-identical
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

  it('flip-only saves keep a Word-authored effectExtent untouched', async () => {
    // Flips do not change the bounding box: Word-authored padding for shadow /
    // glow effects must survive a mirror toggle byte-for-byte
    const wordEffectExtent = '<wp:effectExtent l="9525" t="19050" r="28575" b="38100"/>'
    const withEffect = IMAGE_PARAGRAPH_XML.replace(
      '<wp:extent cx="914400" cy="914400"/>',
      `<wp:extent cx="914400" cy="914400"/>${wordEffectExtent}`,
    ).replace(
      '</pic:blipFill></pic:pic>',
      '</pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>',
    )
    const source = await buildDocx({ bodyXml: withEffect, withImage: true })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    editor.commands.updateAttributes('docProtected', { imageFlipH: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const docXml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(docXml).toContain('flipH="1"')
    expect(docXml).toContain(wordEffectExtent)
    editor.destroy()
  })

  it('90° rotation of a non-square image records the bounding-box overflow in effectExtent', async () => {
    // 2:1 landscape picture (1828800 x 914400 EMU): turned 90° it overflows the
    // unrotated footprint by (w-h)/2 = 457200 EMU on top/bottom, and needs that
    // recorded in wp:effectExtent or Word crops / misplaces the drawing
    const nonSquare = IMAGE_PARAGRAPH_XML.replace(
      '<wp:extent cx="914400" cy="914400"/>',
      '<wp:extent cx="1828800" cy="914400"/>',
    ).replace(
      '</pic:blipFill></pic:pic>',
      '</pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>' +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>',
    )
    const source = await buildDocx({ bodyXml: nonSquare, withImage: true })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    editor.commands.updateAttributes('docProtected', { imageRotDeg: 90 })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)

    const zip = await JSZip.loadAsync(saved)
    const docXml = await zip.file('word/document.xml')!.async('string')
    // (bboxW - cx)/2 = (914400 - 1828800)/2 clamps to 0; (bboxH - cy)/2 = 457200
    expect(docXml).toContain('<wp:effectExtent l="0" t="457200" r="0" b="457200"/>')
    // clearing the rotation zeroes the extra space again
    const reparsed = await parseDocx(saved)
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks) as never,
    })
    editor2.view.dispatch(editor2.state.tr.setSelection(NodeSelection.create(editor2.state.doc, 0)))
    editor2.commands.updateAttributes('docProtected', { imageRotDeg: null })
    const cleared = await saveDocx(
      reparsed,
      pmDocToSavePlan(editor2.getJSON() as PmNode, reparsed.blocks).saveBlocks,
    )
    const clearedXml = await (
      await JSZip.loadAsync(cleared)
    )
      .file('word/document.xml')!
      .async('string')
    expect(clearedXml).toContain('<wp:effectExtent l="0" t="0" r="0" b="0"/>')
    editor.destroy()
    editor2.destroy()
  })

  it('imageReplace strips a stale external r:link and svgBlip extension', async () => {
    const GIF_B64 = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'
    // Word "insert and link" SVG picture: the blip carries the raster fallback
    // (r:embed), an external link (r:link), and the Office-2016 svgBlip extension
    const linkedBody = IMAGE_PARAGRAPH_XML.replace(
      '<a:blip r:embed="rId10"/>',
      '<a:blip r:embed="rId10" r:link="rId11">' +
        '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
        '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId10"/>' +
        '</a:ext></a:extLst></a:blip>',
    )
    const linkRel =
      '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="file:///C:/old.png" TargetMode="External"/>'
    const { editor, parsed } = await openImageDoc(linkedBody, linkRel)
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    editor.commands.updateAttributes('docProtected', {
      imageReplace: { base64: GIF_B64, mime: 'image/gif' },
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)

    // Word would refresh a surviving r:link from the old file, and prefers a
    // surviving svgBlip over the retargeted embed — both must be gone
    const zip = await JSZip.loadAsync(saved)
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).not.toContain('r:link')
    expect(docXml).not.toContain('svgBlip')
    const reparsed2 = await parseDocx(saved)
    expect(reparsed2.blocks[0].imageDataUrl?.startsWith('data:image/gif;base64,')).toBe(true)
    editor.destroy()
  })
})
