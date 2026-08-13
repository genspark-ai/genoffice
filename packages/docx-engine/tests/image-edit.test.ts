import { describe, expect, it } from 'vitest'
import { applyImageWrap, parseDocx, patchImageParagraphXml } from '../src/index'
import { buildDocx, IMAGE_PARAGRAPH_XML } from './helpers/build-docx'

/** A minimal floating image paragraph with numeric posOffset */
const ANCHOR_IMAGE_XML =
  '<w:p><w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="914400"/>' +
  '<wp:wrapSquare wrapText="bothSides"/>' +
  '<wp:docPr id="1" name="图片 1"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic>' +
  '</wp:anchor>' +
  '</w:drawing></w:r></w:p>'

describe('image metadata parsing', () => {
  it('exposes display size from wp:extent', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true }))
    const img = doc.blocks[0]
    expect(img.type).toBe('image')
    // 914400 EMU = 96 px
    expect(img.imageWidthPx).toBe(96)
    expect(img.imageHeightPx).toBe(96)
    expect(img.imageAlign).toBeUndefined()
  })

  it('exposes centered alignment', async () => {
    const centered = IMAGE_PARAGRAPH_XML.replace(
      '<w:p>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: centered, withImage: true }))
    expect(doc.blocks[0].imageAlign).toBe('center')
  })
})

describe('patchImageParagraphXml', () => {
  it('rewrites wp:extent and a:ext size', () => {
    const xml =
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
      '<a:graphic><pic:pic><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr></pic:pic></a:graphic>' +
      '</wp:inline></w:drawing></w:r></w:p>'
    const out = patchImageParagraphXml(xml, { widthPx: 192, heightPx: 96 })
    expect(out).toContain('<wp:extent cx="1828800" cy="914400"/>')
    expect(out).toContain('<a:ext cx="1828800" cy="914400"/>')
  })

  it('adds, replaces and removes w:jc', () => {
    const xml = '<w:p><w:r><w:t>x</w:t></w:r></w:p>'
    const centered = patchImageParagraphXml(xml, { align: 'center' })
    expect(centered).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>')

    const right = patchImageParagraphXml(centered, { align: 'right' })
    expect(right).toContain('<w:jc w:val="right"/>')
    expect(right).not.toContain('w:val="center"')

    const cleared = patchImageParagraphXml(right, { align: null })
    expect(cleared).not.toContain('<w:jc')
  })

  it('inserts w:jc before the paragraph-mark rPr', () => {
    const xml = '<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
    const out = patchImageParagraphXml(xml, { align: 'center' })
    expect(out).toContain('<w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr>')
  })

  it('round-trips size + align through the parser', async () => {
    const patched = patchImageParagraphXml(IMAGE_PARAGRAPH_XML, {
      widthPx: 200,
      heightPx: 150,
      align: 'center',
    })
    const doc = await parseDocx(await buildDocx({ bodyXml: patched, withImage: true }))
    expect(doc.blocks[0].imageWidthPx).toBe(200)
    expect(doc.blocks[0].imageHeightPx).toBe(150)
    expect(doc.blocks[0].imageAlign).toBe('center')
  })
})

describe('image posOffset (free-position drag)', () => {
  it('parses posOffset from anchor images', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: ANCHOR_IMAGE_XML, withImage: true }))
    const img = doc.blocks[0]
    expect(img.type).toBe('image')
    expect(img.imageOffsetXEmu).toBe(914400)
    expect(img.imageOffsetYEmu).toBe(457200)
    expect(img.imageWrap).toBe('square-left')
  })

  it('patchImageParagraphXml rewrites posOffset values', () => {
    const out = patchImageParagraphXml(ANCHOR_IMAGE_XML, {
      posOffsetX: 1828800,
      posOffsetY: 914400,
    })
    expect(out).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>1828800</wp:posOffset></wp:positionH>',
    )
    expect(out).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>914400</wp:posOffset></wp:positionV>',
    )
    // Original structure untouched
    expect(out).toContain('relativeHeight="251658240"')
  })

  it('patchImageParagraphXml leaves images without posOffset unchanged', () => {
    // IMAGE_PARAGRAPH_XML is inline (no posOffset) — no rewrite should happen
    const unchanged = patchImageParagraphXml(IMAGE_PARAGRAPH_XML, {
      posOffsetX: 100000,
    })
    expect(unchanged).toBe(IMAGE_PARAGRAPH_XML)
  })

  it('applyImageWrap with posOffset uses numeric posOffset instead of align', () => {
    const out = applyImageWrap(IMAGE_PARAGRAPH_XML, 'square-left', { x: 914400, y: 457200 })
    expect(out).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>',
    )
    expect(out).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>',
    )
    expect(out).not.toContain('<wp:align>')
  })

  it('applyImageWrap without posOffset uses named alignment', () => {
    const out = applyImageWrap(IMAGE_PARAGRAPH_XML, 'square-left')
    expect(out).toContain('<wp:align>left</wp:align>')
    // V position uses posOffset(0) when no numeric offset provided (existing behavior)
    expect(out).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
    )
    // H position uses named align, not posOffset
    expect(out).not.toContain('<wp:positionH relativeFrom="column"><wp:posOffset>')
  })

  it('round-trips posOffset through the parser', async () => {
    // Apply wrap with posOffset, then parse back
    const wrapped = applyImageWrap(IMAGE_PARAGRAPH_XML, 'square-left', { x: 1828800, y: 914400 })
    const doc = await parseDocx(await buildDocx({ bodyXml: wrapped, withImage: true }))
    const img = doc.blocks[0]
    expect(img.imageWrap).toBe('square-left')
    expect(img.imageOffsetXEmu).toBe(1828800)
    expect(img.imageOffsetYEmu).toBe(914400)
  })
})
