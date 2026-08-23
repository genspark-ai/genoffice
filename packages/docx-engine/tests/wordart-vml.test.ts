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
    const box = block.textboxes?.[0]
    // 1130/1976 × 126pt = 72pt = 96px; 1115/1533 × 99pt = 72pt = 96px
    expect(box?.widthPx).toBe(96)
    expect(box?.heightPx).toBe(96)
    // VML strokes default on/black — Word draws the canvas textbox border
    expect(box?.borderColor).toBe('000000')
    expect(box?.paras[0]?.runs[0]?.text).toBe('Словарь')
  })

  it('keeps paragraph text next to the canvas visible as a display-only line', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: canvasParagraph('<w:r><w:t xml:space="preserve">Hyperlink line</w:t></w:r>'),
      }),
    )
    const block = doc.blocks[0]
    expect(block.label).toBe('Text box')
    expect(block.textboxes).toHaveLength(2)
    const stray = block.textboxes?.[1]
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
    const stray = doc.blocks[0].textboxes?.[1]
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
