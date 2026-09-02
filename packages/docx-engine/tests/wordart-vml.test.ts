import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

/**
 * Legacy VML WordArt (v:textpath, shapetype 136 family) and drawing-canvas
 * groups (v:group editas="canvas"), as written by Word 2003-2010. Both used
 * to degrade to an opaque "Embedded object" chip; they now surface their text.
 */

const V_NS =
  'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'

function wordArtPict(
  shapeAttrs: string,
  inner: string,
  pPr = '<w:pPr><w:jc w:val="center"/></w:pPr>',
): string {
  return (
    `<w:p>${pPr}<w:r><w:pict>` +
    `<v:shape ${V_NS} id="wa1" type="#_x0000_t136" ${shapeAttrs}>${inner}</v:shape>` +
    '</w:pict></w:r></w:p>'
  )
}

describe('VML WordArt (v:textpath) display', () => {
  it('renders the textpath string as a sized, styled text box instead of a chip', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: wordArtPict(
          'style="width:148.5pt;height:28.5pt" fillcolor="#9400ed" strokecolor="#eaeaea" strokeweight="1pt"',
          '<v:textpath style="font-family:&quot;Arial Black&quot;;font-size:20pt" trim="t" fitpath="t" string="My Text Here"/>',
        ),
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Text box')
    expect(block.imageAlign).toBe('center')
    const box = block.textboxes?.[0]
    expect(box?.widthPx).toBe(198) // 148.5pt
    expect(box?.heightPx).toBe(38) // 28.5pt
    expect(box?.readOnly).toBe(true)
    expect(box?.textOutline).toEqual({ colorHex: 'eaeaea', widthPx: 1.33 })
    const run = box?.paras[0]?.runs[0]
    expect(box?.paras[0]?.align).toBe('center')
    expect(run?.text).toBe('My Text Here')
    expect(run?.sizeHalfPoints).toBe(40) // declared 20pt
    expect(run?.fontAscii).toBe('Arial Black')
    expect(run?.color).toBe('9400ed')
  })

  it('falls back to the gradient color2 when the shape has no fillcolor, honors stroked="f"', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: wordArtPict(
          'style="width:187.5pt;height:28.5pt" stroked="f"',
          '<v:fill color2="#aaaaaa" type="gradient"/>' +
            '<v:textpath style="font-family:&quot;Arial Black&quot;;font-size:20pt" string="Мой текст"/>',
        ),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.paras[0]?.runs[0]?.color).toBe('aaaaaa')
    expect(box?.textOutline).toBeUndefined()
  })

  it('keeps anchored DrawingML photos sharing the paragraph with the WordArt shape', async () => {
    const anchoredPic = (blipXml: string, cx: number, cy: number): string =>
      '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ' +
      'relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>' +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      `<pic:pic><pic:blipFill>${blipXml}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      // pic:spPr carries an EMPTY a:blip in its own blipFill (real-world Word
      // output): media resolution must not trip over it
      '<pic:spPr><a:blipFill dpi="0" rotWithShape="0"><a:blip/></a:blipFill>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
      '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    const bodyXml =
      '<w:p><w:r><w:pict>' +
      `<v:shape ${V_NS} id="wa1" type="#_x0000_t136" style="width:148.5pt;height:28.5pt" fillcolor="#9400ed">` +
      '<v:textpath style="font-family:&quot;Arial Black&quot;;font-size:20pt" string="Moon Landing"/></v:shape>' +
      '</w:pict></w:r>' +
      anchoredPic('<a:blip r:embed="rId10"/>', 1418590, 1201420) +
      // second blip carries both r:embed and r:link — the embedded part must resolve
      anchoredPic('<a:blip r:embed="rId11" r:link="rId12"/>', 1448435, 1280160) +
      '</w:p>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml,
        withImage: true,
        extraRels:
          '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
          '<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="http://example.com/moon.jpg" TargetMode="External"/>',
      }),
    )
    const block = doc.blocks[0]
    expect(block.label).toBe('Text box')
    const boxes = block.textboxes ?? []
    expect(boxes.some((b) => b.paras[0]?.runs[0]?.text === 'Moon Landing')).toBe(true)
    const photos = boxes.filter((b) => b.fillImageDataUrl)
    expect(photos).toHaveLength(2)
    for (const photo of photos) {
      expect(photo.fillImageDataUrl).toMatch(/^data:image\/png;base64,/)
    }
    expect(photos.map((p) => p.widthPx)).toEqual([149, 152])
  })

  it('does not lift a picture nested in VML textbox content into a page-level box', async () => {
    const innerDrawing =
      '<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing>'
    const bodyXml =
      `<w:p><w:r><w:pict><v:shape ${V_NS} id="tb1" type="#_x0000_t202" style="width:200pt;height:120pt">` +
      '<v:textbox><w:txbxContent>' +
      `<w:p><w:r><w:t>caption</w:t></w:r></w:p><w:p><w:r>${innerDrawing}</w:r></w:p>` +
      '</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const boxes = doc.blocks[0].textboxes ?? []
    expect(boxes).toHaveLength(1)
    expect(boxes[0].fillImageDataUrl).toBeUndefined()
    expect(boxes[0].paras[0]?.runs[0]?.text).toBe('caption')
  })

  it('page-pins page-anchored wrapNone photos next to WordArt like the drawing branch', async () => {
    const pageAnchoredPic = (x: number, y: number): string =>
      '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ' +
      'relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      `<wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
      '<wp:extent cx="914400" cy="914400"/><wp:wrapNone/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
      '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    const bodyXml =
      '<w:p><w:r><w:t>logo line above</w:t></w:r></w:p>' +
      `<w:p><w:r><w:pict><v:shape ${V_NS} id="wa1" type="#_x0000_t136" style="width:148.5pt;height:28.5pt">` +
      '<v:textpath style="font-family:&quot;Arial Black&quot;" string="Cover Title"/></v:shape></w:pict></w:r>' +
      pageAnchoredPic(914400, 1828800) +
      pageAnchoredPic(2743200, 1828800) +
      '</w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const photos = (doc.blocks[1].textboxes ?? []).filter((b) => b.fillImageDataUrl)
    expect(photos).toHaveLength(2)
    // raw page coordinates, pinned to the page box (not the paragraph origin)
    expect(photos.map((p) => p.pagePinned)).toEqual([true, true])
    expect(photos.map((p) => p.offsetXEmu)).toEqual([914400, 2743200])
    expect(photos.map((p) => p.offsetYEmu)).toEqual([1828800, 1828800])
  })

  it('ignores string-less v:textpath (shapetype template / watermark furniture)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p><w:r><w:pict><v:shapetype ${V_NS} id="_x0000_t136" o:spt="136">` +
          '<v:textpath on="t" fitshape="t"/></v:shapetype></w:pict></w:r></w:p>',
      }),
    )
    expect(doc.blocks[0].textboxes).toBeUndefined()
  })
})

describe('VML drawing canvas (v:group) textboxes', () => {
  const canvasParagraph = (trailing = ''): string =>
    `<w:p><w:r><w:pict><v:group ${V_NS} editas="canvas" ` +
    'style="width:126pt;height:99pt" coordorigin="3834,5690" coordsize="1976,1533">' +
    '<v:shape id="bg" type="#_x0000_t75" style="position:absolute;left:3834;top:5690;width:1976;height:1533" o:preferrelative="f"/>' +
    '<v:shape id="tb" type="#_x0000_t202" style="position:absolute;left:4257;top:5969;width:1130;height:1115">' +
    '<v:textbox><w:txbxContent><w:p><w:r><w:t>Словарь</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
    '</v:shape></v:group></w:pict></w:r>' +
    trailing +
    '</w:p>'

  it('scales group-coordinate children through the canvas extent', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: canvasParagraph() }))
    const block = doc.blocks[0]
    expect(block.label).toBe('Text box')
    // the canvas itself reserves its flow footprint (126×99pt)
    const spacer = block.textboxes?.[0]
    expect(spacer?.widthPx).toBe(168)
    expect(spacer?.heightPx).toBe(132)
    expect(spacer?.paras).toHaveLength(0)
    const box = block.textboxes?.[1]
    // 1130/1976 × 126pt = 72pt = 96px; 1115/1533 × 99pt = 72pt = 96px
    expect(box?.widthPx).toBe(96)
    expect(box?.heightPx).toBe(96)
    // VML strokes default on/black — Word draws the canvas textbox border
    expect(box?.borderColor).toBe('000000')
    expect(box?.paras[0]?.runs[0]?.text).toBe('Словарь')
    // floated at its canvas position: (4257-3834)×sx, (5969-5690)×sy (EMU: ×9525)
    expect(box?.floating).toBe(true)
    expect(box?.offsetXEmu).toBe(Math.round((4257 - 3834) * (168 / 1976) * 9525))
    expect(box?.offsetYEmu).toBe(Math.round((5969 - 5690) * (132 / 1533) * 9525))
  })

  it('keeps paragraph text next to the canvas visible as a display-only line', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: canvasParagraph('<w:r><w:t xml:space="preserve">Hyperlink line</w:t></w:r>'),
      }),
    )
    const block = doc.blocks[0]
    expect(block.label).toBe('Text box')
    expect(block.textboxes).toHaveLength(3)
    const stray = block.textboxes?.[2]
    expect(stray?.readOnly).toBe(true)
    expect(stray?.paras[0]?.runs.map((r) => r.text).join('')).toBe('Hyperlink line')
    expect(block.previewText).toContain('Словарь')
    expect(block.previewText).toContain('Hyperlink line')
  })

  it('keeps the hyperlink target on stray runs next to the canvas', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: canvasParagraph(
          '<w:hyperlink r:id="rId20" w:history="1"><w:r>' +
            '<w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
            '<w:t>Hyperlink</w:t></w:r></w:hyperlink>',
        ),
        extraRels:
          '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://www.google.com" TargetMode="External"/>',
      }),
    )
    const stray = doc.blocks[0].textboxes?.[2]
    const run = stray?.paras[0]?.runs[0]
    expect(run?.text).toBe('Hyperlink')
    expect(run?.link?.href).toBe('http://www.google.com')
  })

  it('ignores a nested textbox jc when the host paragraph has none', async () => {
    const bodyXml =
      `<w:p><w:r><w:pict><v:shape ${V_NS} id="tb3" type="#_x0000_t202" style="width:100pt;height:20pt">` +
      '<v:textbox><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
      '<w:r><w:t>inner</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>' +
      '</w:pict></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    const block = doc.blocks[0]
    expect(block.label).toBe('Text box')
    expect(block.imageAlign).toBeUndefined()
    // the inner paragraph keeps its own alignment inside the box
    expect(block.textboxes?.[0]?.paras[0]?.align).toBe('center')
  })

  it('leaves paragraphs mixing a VML picture with real text on the editable path', async () => {
    const bodyXml =
      `<w:p><w:r><w:pict><v:shape ${V_NS} id="p1" type="#_x0000_t75" style="width:54pt;height:38.25pt">` +
      '<v:imagedata r:id="rId999" o:title=""/></v:shape>' +
      '<v:shape id="tb2" type="#_x0000_t202" style="width:100pt;height:20pt">' +
      '<v:textbox><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>' +
      '</w:pict></w:r><w:r><w:t>editable text</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    // not converted into a protected textbox block: the run text stays editable
    expect(doc.blocks[0].label).not.toBe('Text box')
  })
})

describe('w14:textFill run color approximation', () => {
  const W14_NS = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'

  it('takes the solid fill color', async () => {
    const bodyXml =
      `<w:p><w:r><w:rPr><w14:textFill ${W14_NS}><w14:solidFill><w14:srgbClr w14:val="ED7D31"/></w14:solidFill></w14:textFill></w:rPr>` +
      '<w:t>Solid</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].runs?.[0]?.color).toBe('ED7D31')
  })

  it('averages gradient stops (with tint/shade transforms) into one display color', async () => {
    const bodyXml =
      `<w:p><w:r><w:rPr><w14:textFill ${W14_NS}><w14:gradFill><w14:gsLst>` +
      '<w14:gs w14:pos="0"><w14:srgbClr w14:val="ED7D31"><w14:tint w14:val="70000"/></w14:srgbClr></w14:gs>' +
      '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="ED7D31"><w14:shade w14:val="50000"/></w14:srgbClr></w14:gs>' +
      '</w14:gsLst></w14:gradFill></w14:textFill></w:rPr>' +
      '<w:t>Grad</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    const color = doc.blocks[0].runs?.[0]?.color
    // stop1 = ED7D31 tinted 70%, stop2 = ED7D31 shaded 50%: average is a mid orange
    expect(color).toBe('B47144')
  })

  it('never overrides an explicit w:color', async () => {
    const bodyXml =
      `<w:p><w:r><w:rPr><w:color w:val="112233"/><w14:textFill ${W14_NS}><w14:solidFill><w14:srgbClr w14:val="ED7D31"/></w14:solidFill></w14:textFill></w:rPr>` +
      '<w:t>Explicit</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].runs?.[0]?.color).toBe('112233')
  })
})

describe('VML pictures and geometry in w:pict', () => {
  const txbxShape =
    '<v:shape id="tb" type="#_x0000_t202" style="position:absolute;left:100;top:100;width:400;height:100">' +
    '<v:textbox><w:txbxContent><w:p><w:r><w:t>label</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>'

  it('renders a grouped t75 picture shape as a floating photo box', async () => {
    const bodyXml =
      `<w:p><w:r><w:pict><v:group ${V_NS} style="width:150pt;height:75pt" coordsize="2000,1000">` +
      '<v:shape id="pic" type="#_x0000_t75" style="position:absolute;left:100;top:50;width:200;height:100">' +
      '<v:imagedata r:id="rId10" o:title=""/></v:shape>' +
      txbxShape +
      '</v:group></w:pict></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const boxes = doc.blocks[0].textboxes ?? []
    // canvas spacer, photo, textbox — in document order
    expect(boxes).toHaveLength(3)
    const photo = boxes[1]
    expect(photo?.fillImageDataUrl).toContain('data:image/')
    // 150pt=200px, 2000 units wide → sx 0.1; 75pt=100px, 1000 units → sy 0.1
    expect(photo?.widthPx).toBe(20)
    expect(photo?.heightPx).toBe(10)
    expect(photo?.floating).toBe(true)
    expect(photo?.offsetXEmu).toBe(Math.round(10 * 9525))
    expect(photo?.offsetYEmu).toBe(Math.round(5 * 9525))
  })

  it('converts an omitted-coordinate v:path into normalized pathData', async () => {
    const bodyXml =
      `<w:p><w:r><w:pict><v:group ${V_NS} style="width:150pt;height:75pt" coordsize="2000,1000">` +
      '<v:shape id="bg" style="position:absolute;left:0;top:0;width:2000;height:1000" ' +
      'fillcolor="#5B9BD5" stroked="f" coordsize="21600,21600" ' +
      'path="m,l21600,,21600,21600,,21600nfxe"/>' +
      txbxShape +
      '</v:group></w:pict></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    const geom = (doc.blocks[0].textboxes ?? []).find((b) => b.fill === '5B9BD5')
    expect(geom).toBeDefined()
    expect(geom?.pathData?.path).toBe('M 0 0 L 1 0 L 1 1 L 0 1 Z')
    expect(geom?.widthPx).toBe(200)
    expect(geom?.heightPx).toBe(100)
  })

  it('keeps lowercase #ffffff unstroked placeholders invisible', async () => {
    const bodyXml =
      `<w:p><w:r><w:pict><v:group ${V_NS} style="width:150pt;height:75pt" coordsize="2000,1000">` +
      '<v:rect id="ph" style="position:absolute;left:0;top:0;width:2000;height:1000" ' +
      'fillcolor="#ffffff" stroked="f"/>' +
      txbxShape +
      '</v:group></w:pict></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    const boxes = doc.blocks[0].textboxes ?? []
    // canvas spacer + textbox only — no box for the white placeholder
    expect(boxes).toHaveLength(2)
    expect(boxes.every((b) => b.fill === undefined)).toBe(true)
  })

  it('does not lift a filled shape nested in textbox content onto the page', async () => {
    const bodyXml =
      `<w:p><w:r><w:pict><v:shape ${V_NS} id="tb4" type="#_x0000_t202" style="width:100pt;height:40pt">` +
      '<v:textbox><w:txbxContent><w:p><w:r><w:t>host</w:t></w:r><w:r><w:pict>' +
      '<v:rect id="inner" style="width:50pt;height:10pt" fillcolor="#FF0000" stroked="f"/>' +
      '</w:pict></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    const boxes = doc.blocks[0].textboxes ?? []
    expect(boxes).toHaveLength(1)
    expect(boxes[0]?.fill).toBeUndefined()
  })
})
