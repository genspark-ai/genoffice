import { describe, expect, it } from 'vitest'
import { applyImageWrap, applyImageZOrder, parseDocx, saveDocx } from '../src/index'
import { buildDocx, IMAGE_PARAGRAPH_XML } from './helpers/build-docx'

const ANCHOR_SQUARE_RIGHT_XML =
  '<w:p><w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="914400"/>' +
  '<wp:wrapSquare wrapText="bothSides"/>' +
  '<wp:docPr id="1" name="图片 1"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

describe('image text wrap (wp:anchor)', () => {
  it('parses wrap mode from anchored images', async () => {
    const bytes = await buildDocx({ bodyXml: ANCHOR_SQUARE_RIGHT_XML, withImage: true })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].type).toBe('image')
    expect(doc.blocks[0].imageWrap).toBe('square-right')
  })

  it('parses behind / topAndBottom / front variants', async () => {
    // behind = behindDoc with NO wrap element; an explicit wrap element wins
    // (Word draws behindDoc+wrapSquare behind the text and still wraps around it)
    const behind = ANCHOR_SQUARE_RIGHT_XML.replace('behindDoc="0"', 'behindDoc="1"').replace(
      '<wp:wrapSquare wrapText="bothSides"/>',
      '<wp:wrapNone/>',
    )
    const behindWithWrap = ANCHOR_SQUARE_RIGHT_XML.replace('behindDoc="0"', 'behindDoc="1"')
    expect(
      (await parseDocx(await buildDocx({ bodyXml: behindWithWrap, withImage: true }))).blocks[0]
        .imageWrap,
    ).toBe('square-right')
    const topBottom = ANCHOR_SQUARE_RIGHT_XML.replace(
      '<wp:wrapSquare wrapText="bothSides"/>',
      '<wp:wrapTopAndBottom/>',
    )
    const front = ANCHOR_SQUARE_RIGHT_XML.replace(
      '<wp:wrapSquare wrapText="bothSides"/>',
      '<wp:wrapNone/>',
    )
    for (const [xml, expected] of [
      [behind, 'behind'],
      [topBottom, 'topBottom'],
      [front, 'front'],
    ] as const) {
      const doc = await parseDocx(await buildDocx({ bodyXml: xml, withImage: true }))
      expect(doc.blocks[0].imageWrap).toBe(expected)
    }
    const inline = await parseDocx(
      await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true }),
    )
    expect(inline.blocks[0].imageWrap).toBeUndefined()
  })

  it('parses the picture outline (pic:spPr a:ln solid fill, tdf#162551)', async () => {
    const withLn = ANCHOR_SQUARE_RIGHT_XML.replace(
      '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>',
      '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
        '<pic:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="FFD428"/></a:solidFill></a:ln></pic:spPr></pic:pic>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: withLn, withImage: true }))
    expect(doc.blocks[0].imageBorder).toEqual({ color: 'FFD428', widthPt: 2.25 })
    const noFill = withLn.replace(
      '<a:solidFill><a:srgbClr val="FFD428"/></a:solidFill>',
      '<a:noFill/>',
    )
    expect(
      (await parseDocx(await buildDocx({ bodyXml: noFill, withImage: true }))).blocks[0]
        .imageBorder,
    ).toBeUndefined()
  })

  it('derives the float side from wrapText and far posOffset (tdf#97090)', async () => {
    const withPosH = (posH: string, wrapText = 'bothSides') =>
      ANCHOR_SQUARE_RIGHT_XML.replace(
        '<wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>',
        posH,
      ).replace('wrapText="bothSides"', `wrapText="${wrapText}"`)
    const cases: Array<[string, string, string]> = [
      // text on the left only -> object floats right
      [
        '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>',
        'left',
        'square-right',
      ],
      // text on the right only -> object floats left even with a far offset
      [
        '<wp:positionH relativeFrom="column"><wp:posOffset>5000000</wp:posOffset></wp:positionH>',
        'right',
        'square-left',
      ],
      // bothSides + absolute X past mid-body -> right
      [
        '<wp:positionH relativeFrom="column"><wp:posOffset>3187064</wp:posOffset></wp:positionH>',
        'bothSides',
        'square-right',
      ],
      // bothSides + near X -> left
      [
        '<wp:positionH relativeFrom="column"><wp:posOffset>100000</wp:posOffset></wp:positionH>',
        'bothSides',
        'square-left',
      ],
    ]
    for (const [posH, wrapText, expected] of cases) {
      const doc = await parseDocx(
        await buildDocx({ bodyXml: withPosH(posH, wrapText), withImage: true }),
      )
      expect(doc.blocks[0].imageWrap, `${wrapText} ${posH}`).toBe(expected)
    }
  })

  it('converts inline -> square anchor with position and wrap elements', () => {
    const out = applyImageWrap(IMAGE_PARAGRAPH_XML, 'square-left')
    expect(out).toContain('<wp:anchor')
    expect(out).not.toContain('<wp:inline')
    expect(out).toContain('behindDoc="0"')
    expect(out).toContain('<wp:align>left</wp:align>')
    expect(out).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    // schema order: positionH before extent, wrap after extent and before the graphic
    expect(out.indexOf('<wp:positionH')).toBeLessThan(out.indexOf('<wp:extent'))
    expect(out.indexOf('<wp:wrapSquare')).toBeGreaterThan(out.indexOf('<wp:extent'))
    expect(out.indexOf('<wp:wrapSquare')).toBeLessThan(out.indexOf('<a:graphic'))
  })

  it('converts anchor -> inline', () => {
    const out = applyImageWrap(ANCHOR_SQUARE_RIGHT_XML, null)
    expect(out).toContain('<wp:inline')
    expect(out).not.toContain('<wp:anchor')
    expect(out).not.toContain('wp:positionH')
    expect(out).not.toContain('wp:wrapSquare')
    expect(out).toContain('<wp:extent cx="914400"')
  })

  it('changes wrap on an existing anchor', () => {
    const out = applyImageWrap(ANCHOR_SQUARE_RIGHT_XML, 'behind')
    expect(out).toContain('behindDoc="1"')
    expect(out).toContain('<wp:wrapNone/>')
    expect(out).not.toContain('wp:wrapSquare')
  })

  it('zOrder raises relativeHeight so stacked behind anchors keep paint order', () => {
    const out = applyImageWrap(ANCHOR_SQUARE_RIGHT_XML, 'behind', undefined, undefined, 7)
    expect(out).toContain('relativeHeight="251658247"')
    // default stays at the base value
    const plain = applyImageWrap(ANCHOR_SQUARE_RIGHT_XML, 'behind')
    expect(plain).toContain('relativeHeight="251658240"')
  })

  it('applies margin-relative position presets (Word position gallery)', () => {
    const out = applyImageWrap(IMAGE_PARAGRAPH_XML, 'square-right', undefined, {
      h: 'center',
      v: 'bottom',
    })
    expect(out).toContain(
      '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>',
    )
    expect(out).toContain(
      '<wp:positionV relativeFrom="margin"><wp:align>bottom</wp:align></wp:positionV>',
    )
    expect(out).toContain('<wp:wrapSquare wrapText="bothSides"/>')
  })

  it('round-trips a position preset through saveDocx + reparse', async () => {
    const bytes = await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true })
    const parsed = await parseDocx(bytes)
    const saved = await saveDocx(parsed, [
      {
        kind: 'xml',
        xml: applyImageWrap(parsed.blocks[0].originalXml!, 'square-left', undefined, {
          h: 'left',
          v: 'top',
        }),
      },
    ])
    const p2 = await parseDocx(saved)
    expect(p2.blocks[0].imageWrap).toBe('square-left')
    expect(p2.blocks[0].imagePosH).toBe('left')
    expect(p2.blocks[0].imagePosV).toBe('top')
    // numeric posOffset must not leak in from the align-based position
    expect(p2.blocks[0].imageOffsetXEmu).toBeUndefined()
  })

  it('round-trips wrap changes through saveDocx', async () => {
    const bytes = await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true })
    const parsed = await parseDocx(bytes)
    const floated = await saveDocx(parsed, [
      { kind: 'xml', xml: applyImageWrap(parsed.blocks[0].originalXml!, 'square-right') },
    ])
    const p2 = await parseDocx(floated)
    expect(p2.blocks[0].imageWrap).toBe('square-right')
    const backInline = await saveDocx(p2, [
      { kind: 'xml', xml: applyImageWrap(p2.blocks[0].originalXml!, null) },
    ])
    const p3 = await parseDocx(backInline)
    expect(p3.blocks[0].imageWrap).toBeUndefined()
    expect(p3.blocks[0].imageDataUrl).toBeTruthy()
  })

  it('parses imageZOrder from a non-base relativeHeight (round-trip paint order)', async () => {
    // relativeHeight = 251658240 base + 5
    const xml = ANCHOR_SQUARE_RIGHT_XML.replace('relativeHeight="1"', 'relativeHeight="251658245"')
    const doc = await parseDocx(await buildDocx({ bodyXml: xml, withImage: true }))
    expect(doc.blocks[0].imageZOrder).toBe(5)

    // base value (or absent) leaves imageZOrder undefined
    const base = ANCHOR_SQUARE_RIGHT_XML.replace('relativeHeight="1"', 'relativeHeight="251658240"')
    const doc2 = await parseDocx(await buildDocx({ bodyXml: base, withImage: true }))
    expect(doc2.blocks[0].imageZOrder).toBeUndefined()
  })

  it('applyImageZOrder re-encodes relativeHeight and nothing else', () => {
    const out = applyImageZOrder(ANCHOR_SQUARE_RIGHT_XML, 3)
    expect(out).toContain('relativeHeight="251658243"')
    // byte-identical apart from the one attribute value: position basis,
    // wrap element and distances must survive a pure reorder
    expect(out.replace('relativeHeight="251658243"', 'relativeHeight="1"')).toBe(
      ANCHOR_SQUARE_RIGHT_XML,
    )
    // no rank = base level; inline (no anchor) passes through untouched
    expect(applyImageZOrder(ANCHOR_SQUARE_RIGHT_XML)).toContain('relativeHeight="251658240"')
    expect(applyImageZOrder(IMAGE_PARAGRAPH_XML, 5)).toBe(IMAGE_PARAGRAPH_XML)
  })

  it('compresses wild producer relativeHeight values to compact ranks', async () => {
    // LibreOffice-style: arbitrary small relativeHeight (1, 7) decodes to huge
    // negative ranks; parse re-ranks the document's anchors 0..n-1 in paint
    // order so the editor's ±1 reorder steps and CSS bands stay meaningful
    const second = ANCHOR_SQUARE_RIGHT_XML.replace(
      'relativeHeight="1"',
      'relativeHeight="7"',
    ).replace('id="1" name="图片 1"', 'id="2" name="图片 2"')
    const doc = await parseDocx(
      await buildDocx({ bodyXml: ANCHOR_SQUARE_RIGHT_XML + second, withImage: true }),
    )
    expect(doc.blocks[0].imageZOrder).toBeUndefined() // rank 0 = base level
    expect(doc.blocks[1].imageZOrder).toBe(1)
  })

  it('round-trips imageZOrder through applyImageWrap + saveDocx + reparse', async () => {
    const bytes = await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true })
    const parsed = await parseDocx(bytes)
    // float it in front with a stacking rank of 3
    const saved = await saveDocx(parsed, [
      {
        kind: 'xml',
        xml: applyImageWrap(parsed.blocks[0].originalXml!, 'front', undefined, undefined, 3),
      },
    ])
    const p2 = await parseDocx(saved)
    expect(p2.blocks[0].imageWrap).toBe('front')
    expect(p2.blocks[0].imageZOrder).toBe(3)
    // bump the rank; wrap unchanged
    const bumped = await saveDocx(p2, [
      {
        kind: 'xml',
        xml: applyImageWrap(p2.blocks[0].originalXml!, 'front', undefined, undefined, 13),
      },
    ])
    const p3 = await parseDocx(bumped)
    expect(p3.blocks[0].imageZOrder).toBe(13)
  })

  it('embeds new images as anchors when NewImage.wrap is set', async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' })
    const parsed = await parseDocx(bytes)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      {
        kind: 'image',
        image: {
          base64:
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
          mime: 'image/png',
          widthPx: 100,
          heightPx: 80,
          wrap: 'square-left',
        },
      },
    ])
    const p2 = await parseDocx(saved)
    const img = p2.blocks.find((b) => b.type === 'image')!
    expect(img.imageWrap).toBe('square-left')
    expect(img.imageWidthPx).toBe(100)
  })
})

const WRAP_TIGHT_POLYGON =
  '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0">' +
  '<wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/>' +
  '<wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'

describe('tight / through wrap (wrapPolygon fidelity)', () => {
  const tightXml = ANCHOR_SQUARE_RIGHT_XML.replace(
    '<wp:wrapSquare wrapText="bothSides"/>',
    WRAP_TIGHT_POLYGON,
  )
  const throughXml = ANCHOR_SQUARE_RIGHT_XML.replace(
    '<wp:wrapSquare wrapText="bothSides"/>',
    '<wp:wrapThrough wrapText="bothSides"><wp:wrapPolygon><wp:start x="1" y="2"/></wp:wrapPolygon></wp:wrapThrough>',
  )

  it('parses tight / through as their own modes (no square downgrade)', async () => {
    const tight = await parseDocx(await buildDocx({ bodyXml: tightXml, withImage: true }))
    expect(tight.blocks[0].imageWrap).toBe('tight-right')
    const through = await parseDocx(await buildDocx({ bodyXml: throughXml, withImage: true }))
    expect(through.blocks[0].imageWrap).toBe('through-right')
  })

  it('re-applying the same tight wrap keeps the original wrapPolygon bytes', () => {
    const out = applyImageWrap(tightXml, 'tight-right')
    expect(out).toContain(WRAP_TIGHT_POLYGON)
    expect(out).not.toContain('wp:wrapSquare')
  })

  it('free-position drag (posOffset) on a tight image keeps the polygon', () => {
    const out = applyImageWrap(tightXml, 'tight-right', { x: 12345, y: 67890 })
    expect(out).toContain('<wp:posOffset>12345</wp:posOffset>')
    expect(out).toContain(WRAP_TIGHT_POLYGON)
  })

  it('explicitly switching a tight image to square rebuilds as wrapSquare', () => {
    const out = applyImageWrap(tightXml, 'square-left')
    expect(out).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    expect(out).not.toContain('wp:wrapTight')
  })

  it('displaces an allowOverlap="0" anchor colliding with a sibling (tdf#134114)', async () => {
    const anchorPic = (attrs: string, cx: number, cy: number) =>
      '<w:r><w:drawing>' +
      `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" ${attrs}>` +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>180340</wp:posOffset></wp:positionV>' +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      '<wp:wrapSquare wrapText="bothSides"/>' +
      '<wp:docPr id="1" name="p"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
      '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    const body =
      '<w:p>' +
      anchorPic('allowOverlap="1"', 802800, 1198800) +
      anchorPic('allowOverlap="0"', 1432800, 1076400) +
      '<w:r><w:t>No overlap in the frames, please.</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body, withImage: true }))
    const imgs = (doc.blocks[0].runs ?? []).filter((r) => r.image)
    expect(imgs).toHaveLength(2)
    expect(imgs[0].image!.wrap).toBe('square-left')
    // the colliding allowOverlap="0" picture leaves the flow as a front overlay
    // under the collider's box
    expect(imgs[1].image!.wrap).toBe('front')
    expect(imgs[1].image!.offsetYEmu).toBe((126 + 2) * 9525)
  })

  it('tight wrap round-trips through save + reparse', async () => {
    const bytes = await buildDocx({ bodyXml: tightXml, withImage: true })
    const parsed = await parseDocx(bytes)
    const saved = await saveDocx(parsed, [
      {
        kind: 'xml',
        xml: applyImageWrap(parsed.blocks[0].originalXml!, 'tight-right', { x: 100, y: 200 }),
      },
    ])
    const p2 = await parseDocx(saved)
    expect(p2.blocks[0].imageWrap).toBe('tight-left')
    expect(p2.blocks[0].imageOffsetXEmu).toBe(100)
    expect(p2.blocks[0].originalXml).toContain('<wp:wrapPolygon')
  })
})
