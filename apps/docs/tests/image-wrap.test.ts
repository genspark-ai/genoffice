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

  it('imageReplace resets a non-default a:fillRect fill window', async () => {
    const GIF_B64 = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'
    // A Word-authored fill window clips like a crop: it must not survive onto
    // the swapped bytes (the editor clears its imageFillRect attr in step)
    const withFill = IMAGE_PARAGRAPH_XML.replace(
      '</pic:blipFill>',
      '<a:stretch><a:fillRect l="10000" t="10000" r="20000" b="5000"/></a:stretch></pic:blipFill>',
    )
    const { editor, parsed } = await openImageDoc(withFill)
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    editor.commands.updateAttributes('docProtected', {
      imageReplace: { base64: GIF_B64, mime: 'image/gif' },
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const docXml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(docXml).toContain('<a:fillRect/>')
    expect(docXml).not.toContain('<a:fillRect l=')
    const reparsed2 = await parseDocx(saved)
    expect(reparsed2.blocks[0].imageFillRect).toBeUndefined()
    editor.destroy()
  })

  // LibreOffice-style anchors: arbitrary raw relativeHeight (1, 7) instead of
  // Word's 251658240+rank encoding. Parse compresses them to compact ranks.
  const WILD_ANCHOR_XML =
    '<w:p><w:r><w:drawing>' +
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="914400" cy="914400"/>' +
    '<wp:wrapNone/>' +
    '<wp:docPr id="1" name="pic 1"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'
  const WILD_TWO_ANCHORS_XML =
    WILD_ANCHOR_XML +
    WILD_ANCHOR_XML.replace('relativeHeight="1"', 'relativeHeight="7"').replace(
      'id="1" name="pic 1"',
      'id="2" name="pic 2"',
    )

  it('untouched wild-relativeHeight anchors stay byte-identical on save', async () => {
    const { editor, parsed, source } = await openImageDoc(WILD_TWO_ANCHORS_XML)
    expect(parsed.blocks[0].imageZOrder).toBeUndefined() // compact rank 0
    expect(parsed.blocks[1].imageZOrder).toBe(1)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('a z-order edit harmonizes wild sibling anchors to base+rank encoding', async () => {
    const { editor, parsed } = await openImageDoc(WILD_TWO_ANCHORS_XML)
    // send the SECOND picture (rank 1, on top) to the back, like the Arrange menu
    let pos = -1
    editor.state.doc.descendants((n, p) => {
      if (n.type.name === 'docProtected' && n.attrs.imageZOrder === 1) pos = p
      return pos === -1
    })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    editor.commands.updateAttributes('docProtected', { imageZOrder: -1 })

    // both anchors rewritten: the edited one AND its wild sibling (Word paints
    // by raw relativeHeight, so mixed encodings would invert the order)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(2)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const docXml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    const rels = [...docXml.matchAll(/relativeHeight="(\d+)"/g)].map((m) => Number(m[1]))
    expect(rels).toEqual([251658240, 251658239]) // base+0, base-1 — raw order matches intent
    // harmonization is surgical: page anchoring and wrap bytes survive on BOTH
    // anchors (a full rebuild would reset relativeFrom to the editor default)
    expect(docXml.match(/relativeFrom="page"/g)?.length).toBe(4)
    expect(docXml.match(/<wp:wrapNone\/>/g)?.length).toBe(2)

    // reopen: ranks decode small (no wild values left), order preserved
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].imageZOrder).toBeUndefined()
    expect(reparsed.blocks[1].imageZOrder).toBe(-1)
    editor.destroy()
  })

  it('a wrap-only change keeps an existing stacking rank', async () => {
    const xml = IMAGE_PARAGRAPH_XML // inline image, then float it with a rank
    const { editor, parsed } = await openImageDoc(xml)
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    editor.commands.updateAttributes('docProtected', { imageWrap: 'front', imageZOrder: 3 })
    const saved = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].imageZOrder).toBe(3)

    // now change ONLY the wrap mode; the rank must survive the rewrite
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks) as never,
    })
    editor2.view.dispatch(editor2.state.tr.setSelection(NodeSelection.create(editor2.state.doc, 0)))
    editor2.commands.updateAttributes('docProtected', { imageWrap: 'behind' })
    const saved2 = await saveDocx(
      reparsed,
      pmDocToSavePlan(editor2.getJSON() as PmNode, reparsed.blocks).saveBlocks,
    )
    const reparsed2 = await parseDocx(saved2)
    expect(reparsed2.blocks[0].imageWrap).toBe('behind')
    expect(reparsed2.blocks[0].imageZOrder).toBe(3)
    editor.destroy()
    editor2.destroy()
  })

  it('harmonizes a geometry-only edited sibling carrying a wild relativeHeight', async () => {
    // Bugbot #767: when one picture's z-order changes, a SIBLING that was only
    // resized/aligned/rotated must also be re-encoded from its wild producer
    // relativeHeight to base+rank — otherwise Word paints the raw-valued
    // sibling above the re-encoded pictures and inverts the stacking.
    const { editor, parsed } = await openImageDoc(WILD_TWO_ANCHORS_XML)
    // resize the FIRST picture (rank 0) — a pure geometry patch, no wrap/z change
    let firstPos = -1
    let secondPos = -1
    editor.state.doc.descendants((n, p) => {
      if (n.type.name === 'docProtected') {
        if (n.attrs.imageZOrder === 1) secondPos = p
        else if (firstPos === -1) firstPos = p
      }
      return true
    })
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(firstPos, undefined, {
        ...editor.state.doc.nodeAt(firstPos)!.attrs,
        imageWidthPx: 48,
        imageHeightPx: 48,
      }),
    )
    // and send the SECOND picture to the back (the z-order edit that triggers
    // the harmonize pre-scan)
    editor.view.dispatch(
      editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, secondPos)),
    )
    editor.commands.updateAttributes('docProtected', { imageZOrder: -1 })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const docXml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    const rels = [...docXml.matchAll(/relativeHeight="(\d+)"/g)].map((m) => Number(m[1]))
    // BOTH anchors carry base+rank encoding: no wild 1/7 value survives
    expect(rels.every((r) => r >= 251658000)).toBe(true)
    expect(rels).toEqual([251658240, 251658239]) // first base+0, second base-1
    // the re-encode is surgical: page anchoring survives on both anchors
    expect(docXml.match(/relativeFrom="page"/g)?.length).toBe(4)

    const reparsed = await parseDocx(saved)
    // geometry patch preserved on the first picture
    expect(reparsed.blocks[0].imageWidthPx).toBe(48)
    // ranks decode small (no wild values left), order preserved
    expect(reparsed.blocks[0].imageZOrder).toBeUndefined()
    expect(reparsed.blocks[1].imageZOrder).toBe(-1)
    editor.destroy()
  })
})
