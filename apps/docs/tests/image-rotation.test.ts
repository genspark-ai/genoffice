/**
 * Pictures with a pic a:xfrm rot: Word turns the bitmap about the centre of
 * its unrotated extent and lets the flow reserve the rotated bounding box, so
 * a 90° landscape photo stands upright in a portrait slot. Both the block
 * image path and the run-level docInlineImage path (table cells, paragraphs
 * mixing text and anchors) must apply the turn.
 */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { foldQuarterTurnMargins, quarterTurnInsetPx } from '../src/renderer/editor/image-rotation'

const PIC = (xfrmAttrs: string) =>
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  `<pic:spPr><a:xfrm${xfrmAttrs}><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'

/** 200×100 px landscape picture turned a quarter clockwise */
const INLINE_ROTATED =
  '<w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/>' +
  PIC(' rot="5400000"') +
  '</wp:inline></w:drawing></w:r>'

const anchorRotated = (wrap: string, posOffsetEmu: number) =>
  '<w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${posOffsetEmu}</wp:posOffset></wp:positionH>` +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1905000" cy="952500"/>' +
  wrap +
  '<wp:docPr id="1" name="pic 1"/>' +
  PIC(' rot="5400000"') +
  '</wp:anchor></w:drawing></w:r>'

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

describe('quarterTurnInsetPx', () => {
  it('is half the side difference for quarter turns only', () => {
    expect(quarterTurnInsetPx(200, 100, 90)).toBe(50)
    expect(quarterTurnInsetPx(200, 100, 270)).toBe(50)
    expect(quarterTurnInsetPx(200, 100, -90)).toBe(50)
    expect(quarterTurnInsetPx(100, 200, 90)).toBe(-50)
    expect(quarterTurnInsetPx(200, 100, 180)).toBe(0)
    expect(quarterTurnInsetPx(200, 100, 45)).toBe(0)
    expect(quarterTurnInsetPx(200, 100, null)).toBe(0)
  })
})

describe('foldQuarterTurnMargins', () => {
  it('adds the inset to declared margins and leaves other declarations alone', () => {
    expect(foldQuarterTurnMargins(50, 'margin-top:-12.0px;margin-left:146.0px')).toBe(
      'margin-top:38.0px;margin-right:-50.0px;margin-bottom:50.0px;margin-left:96.0px',
    )
    expect(foldQuarterTurnMargins(50, 'margin-left:calc(28px + 1em)')).toBe(
      'margin-top:50.0px;margin-right:-50.0px;margin-bottom:50.0px;margin-left:calc(28px + 1em + -50.0px)',
    )
    expect(foldQuarterTurnMargins(0, 'margin-top:-12.0px')).toBe('margin-top:-12.0px')
    expect(foldQuarterTurnMargins(0, '')).toBe('')
  })
})

describe('rotated pictures', () => {
  it('turns a table-cell picture about its centre inside a swapped footprint', async () => {
    const editor = await open(
      '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p>' +
        INLINE_ROTATED +
        '</w:p></w:tc></w:tr></w:tbl>',
    )
    const img = editor.view.dom.querySelector<HTMLImageElement>('img.doc-inline-img')!
    expect(img.style.width).toBe('200px')
    expect(img.style.height).toBe('100px')
    expect(img.style.transform).toBe('rotate(90deg)')
    expect(img.style.marginTop).toBe('50px')
    expect(img.style.marginBottom).toBe('50px')
    expect(img.style.marginLeft).toBe('-50px')
    expect(img.style.marginRight).toBe('-50px')
    editor.destroy()
  })

  it('leaves an unrotated table-cell picture untouched', async () => {
    const editor = await open(
      '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p>' +
        INLINE_ROTATED.replace(' rot="5400000"', '') +
        '</w:p></w:tc></w:tr></w:tbl>',
    )
    const img = editor.view.dom.querySelector<HTMLImageElement>('img.doc-inline-img')!
    expect(img.style.transform).toBe('')
    expect(img.style.margin).toBe('')
    editor.destroy()
  })

  it('keeps a run-level tight float pinned at its offset and grows toward the text', async () => {
    // 1in offset: the extent box still ends at 96+200px, the text side gives
    // way by the inset (12px distance - 50px); the extent box top stays on the
    // anchor line and the turned visual reaches 50px above it (Word)
    const editor = await open(
      '<w:p><w:r><w:t>before</w:t></w:r>' +
        anchorRotated('<wp:wrapTight wrapText="left"/>', 914400) +
        '<w:r><w:t>after</w:t></w:r></w:p>',
    )
    const img = editor.view.dom.querySelector<HTMLImageElement>(
      'img.doc-inline-img--wrap-tight-right',
    )!
    expect(img.style.transform).toBe('rotate(90deg)')
    expect(img.style.marginRight).toMatch(/calc\(100% - 296(\.0)?px\)/)
    expect(img.style.marginLeft).toBe('-38px')
    expect(img.style.marginTop).toBe('0px')
    expect(img.style.marginBottom).toBe('50px')
    editor.destroy()
  })

  it('composes the turn with the overlay translate of a run-level front anchor', async () => {
    const editor = await open(
      '<w:p><w:r><w:t>before</w:t></w:r>' +
        anchorRotated('<wp:wrapNone/>', 914400) +
        '<w:r><w:t>after</w:t></w:r></w:p>',
    )
    const img = editor.view.dom.querySelector<HTMLImageElement>(
      '.doc-inline-img-anchor > img.doc-inline-img',
    )!
    expect(img.style.transform).toBe('translate(96.0px, 0.0px) rotate(90deg)')
    expect(img.style.marginTop).toBe('')
    editor.destroy()
  })

  it('keeps the band gap of a run-level topBottom picture', async () => {
    const editor = await open(
      '<w:p><w:r><w:t>before</w:t></w:r>' +
        anchorRotated('<wp:wrapTopAndBottom/>', 0).replace(' distT="0" distB="0"', '') +
        '<w:r><w:t>after</w:t></w:r></w:p>',
    )
    const img = editor.view.dom.querySelector<HTMLImageElement>(
      'img.doc-inline-img--wrap-topBottom',
    )!
    expect(img.style.transform).toBe('rotate(90deg)')
    expect(img.style.marginTop).toBe('52px')
    expect(img.style.marginBottom).toBe('52px')
    expect(img.style.marginLeft).toBe('')
    editor.destroy()
  })

  it('pins a run-level right float at its box when the sliver guard fires', async () => {
    // 120x100 picture (inset 10) at x=20px: visual left 30px sits in the sliver
    const editor = await open(
      '<w:p><w:r><w:t>before</w:t></w:r>' +
        anchorRotated('<wp:wrapTight wrapText="left"/>', 190500)
          .replace('<wp:extent cx="1905000" cy="952500"/>', '<wp:extent cx="1143000" cy="952500"/>')
          .replace('<a:ext cx="1905000" cy="952500"/>', '<a:ext cx="1143000" cy="952500"/>') +
        '<w:r><w:t>after</w:t></w:r></w:p>',
    )
    const img = editor.view.dom.querySelector<HTMLImageElement>(
      'img.doc-inline-img--wrap-tight-right',
    )!
    expect(img.style.marginRight).toMatch(/calc\(100% - 140(\.0)?px\)/)
    expect(img.style.marginLeft).toBe('20px')
    editor.destroy()
  })

  it('folds the footprint into the positioning margins of a cropped topBottom picture', async () => {
    const cropped = anchorRotated('<wp:wrapTopAndBottom/>', 914400).replace(
      '<pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>',
      '<pic:blipFill><a:blip r:embed="rId10"/><a:srcRect l="10000" r="10000"/></pic:blipFill>',
    )
    const editor = await open('<w:p>' + cropped + '</w:p>')
    const wrap = editor.view.dom.querySelector<HTMLElement>('.doc-img-crop')!
    expect(wrap.style.transform).toBe('rotate(90deg)')
    expect(wrap.style.marginLeft).toBe('96px')
    // the 50px lift of the turned visual folds into the footprint growth
    expect(wrap.style.marginTop).toBe('0px')
    expect(wrap.style.marginRight).toBe('-50px')
    expect(wrap.style.marginBottom).toBe('50px')
    editor.destroy()
  })

  it('swaps the footprint of a block-level inline picture', async () => {
    const editor = await open('<w:p>' + INLINE_ROTATED + '</w:p>')
    const img = editor.view.dom.querySelector<HTMLImageElement>(
      '.doc-protected[data-doc-protected="image"] img',
    )!
    expect(img.style.transform).toBe('rotate(90deg)')
    expect(img.style.margin).toBe('50px -50px')
    editor.destroy()
  })
})
