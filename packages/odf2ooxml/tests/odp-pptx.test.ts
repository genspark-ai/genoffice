import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { convertOdpToPptx } from '../src/index'

const NAMESPACES =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
  ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' +
  ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
  ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
  ' xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"' +
  ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
  ' xmlns:xlink="http://www.w3.org/1999/xlink"'

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

/** 25.4cm × 19.05cm = 10in × 7.5in → 9144000 × 6858000 EMU */
function contentXml(body: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.2">` +
    `<office:body><office:presentation>${body}</office:presentation></office:body></office:document-content>`
  )
}

function stylesXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.2">` +
    '<office:automatic-styles>' +
    '<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="25.4cm" fo:page-height="19.05cm"/></style:page-layout>' +
    '</office:automatic-styles>' +
    '<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>' +
    '</office:document-styles>'
  )
}

function frame(
  x: string,
  y: string,
  width: string,
  height: string,
  inner: string,
  presentationClass = '',
): string {
  const classAttr = presentationClass ? ` presentation:class="${presentationClass}"` : ''
  return `<draw:frame draw:layer="layout"${classAttr} svg:x="${x}" svg:y="${y}" svg:width="${width}" svg:height="${height}">${inner}</draw:frame>`
}

async function buildOdp(body: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation')
  zip.file('content.xml', contentXml(body))
  zip.file('styles.xml', stylesXml())
  return zip.generateAsync({ type: 'uint8array' })
}

async function readPptxEntry(pptx: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(pptx)
  const entry = zip.file(path)
  return entry ? entry.async('string') : null
}

describe('convertOdpToPptx', () => {
  it('makes one slide per draw:page at the declared deck size', async () => {
    const page = (inner: string) => `<draw:page draw:name="p">${inner}</draw:page>`
    const body =
      page(
        frame(
          '2cm',
          '1cm',
          '20cm',
          '2cm',
          '<draw:text-box><text:p>Title on one</text:p></draw:text-box>',
        ),
      ) +
      page(
        frame(
          '2cm',
          '5cm',
          '20cm',
          '4cm',
          '<draw:text-box><text:p>Body of two</text:p></draw:text-box>',
        ),
      )
    const result = await convertOdpToPptx(await buildOdp(body))
    expect(result.slides).toBe(2)
    const presentation = (await readPptxEntry(result.pptx, 'ppt/presentation.xml')) ?? ''
    expect(presentation).toContain('sldSz cx="9144000" cy="6858000"')
    const slide2 = (await readPptxEntry(result.pptx, 'ppt/slides/slide2.xml')) ?? ''
    expect(slide2).toContain('Body of two')
  })

  it('keeps bold runs and text inside positioned textboxes', async () => {
    const body =
      '<draw:page draw:name="p">' +
      frame(
        '1cm',
        '1cm',
        '20cm',
        '10cm',
        '<draw:text-box><text:p>plain <text:span text:style-name="P2">shout</text:span></text:p></draw:text-box>',
      ) +
      '</draw:page>'
    const zip = new JSZip()
    zip.file(
      'content.xml',
      contentXml(body).replace(
        '<office:body>',
        '<office:automatic-styles><style:style style:name="P2" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style></office:automatic-styles><office:body>',
      ),
    )
    zip.file('styles.xml', stylesXml())
    const result = await convertOdpToPptx(await zip.generateAsync({ type: 'uint8array' }))
    const slide1 = (await readPptxEntry(result.pptx, 'ppt/slides/slide1.xml')) ?? ''
    expect(slide1).toContain('plain')
    expect(slide1).toContain('shout')
    expect(slide1).toMatch(/b="1"/)
    // 1cm → 360000 EMU textbox offset
    expect(slide1).toContain('<a:off x="360000" y="360000"/>')
  })

  it('lands framed images at their frame rectangle', async () => {
    const body =
      '<draw:page draw:name="p">' +
      frame('3cm', '3cm', '6cm', '4cm', '<draw:image xlink:href="Pictures/red.png"/>') +
      '</draw:page>'
    const zip = new JSZip()
    zip.file('content.xml', contentXml(body))
    zip.file('styles.xml', stylesXml())
    zip.file('Pictures/red.png', PNG_BASE64, { base64: true })
    const result = await convertOdpToPptx(await zip.generateAsync({ type: 'uint8array' }))
    const slide1 = (await readPptxEntry(result.pptx, 'ppt/slides/slide1.xml')) ?? ''
    expect(slide1).toContain('<p:pic>')
    // 3cm → 1080000 EMU, 6cm → 2160000, 4cm → 1440000
    expect(slide1).toContain('<a:off x="1080000" y="1080000"/>')
    expect(slide1).toContain('<a:ext cx="2160000" cy="1440000"/>')
    const zipOut = await JSZip.loadAsync(result.pptx)
    expect(Object.keys(zipOut.files).some((path) => /^ppt\/media\/.+\.(png|jpg)$/.test(path))).toBe(
      true,
    )
  })

  it('gives title-class frames a default box when coordinates are missing', async () => {
    const body =
      '<draw:page draw:name="p">' +
      '<draw:frame presentation:class="title" svg:width="20cm" svg:height="2cm"><draw:text-box><text:p>No origin</text:p></draw:text-box></draw:frame>' +
      '</draw:page>'
    const result = await convertOdpToPptx(await buildOdp(body))
    const slide1 = (await readPptxEntry(result.pptx, 'ppt/slides/slide1.xml')) ?? ''
    expect(slide1).toContain('No origin')
    // the title fallback box is 8% × 4% of the deck
    expect(slide1).toContain(
      `<a:off x="${Math.round(0.08 * 9144000)}" y="${Math.round(0.04 * 6858000)}"/>`,
    )
  })

  it('counts shapes and notes it cannot carry', async () => {
    const body =
      '<draw:page draw:name="p">' +
      frame('1cm', '1cm', '5cm', '1cm', '<draw:text-box><text:p>kept</text:p></draw:text-box>') +
      '<draw:line svg:x1="0cm" svg:y1="0cm" svg:x2="5cm" svg:y2="5cm"/>' +
      '<presentation:notes><draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="5cm"><draw:text-box><text:p>speaker notes</text:p></draw:text-box></draw:frame></presentation:notes>' +
      '</draw:page>'
    const result = await convertOdpToPptx(await buildOdp(body))
    expect(result.dropped).toContainEqual({ what: 'shapes', count: 1 })
    expect(result.dropped).toContainEqual({ what: 'notes', count: 1 })
    const slide1 = (await readPptxEntry(result.pptx, 'ppt/slides/slide1.xml')) ?? ''
    expect(slide1).toContain('kept')
    expect(slide1).not.toContain('speaker notes')
  })

  it('accepts .otp templates (a -template mimetype, same body shape)', async () => {
    const zip = new JSZip()
    zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation-template')
    zip.file(
      'content.xml',
      contentXml(
        '<draw:page draw:name="p1">' +
          frame(
            '1cm',
            '1cm',
            '10cm',
            '2cm',
            '<draw:text-box><text:p>Template slide</text:p></draw:text-box>',
          ) +
          '</draw:page>',
      ),
    )
    zip.file('styles.xml', stylesXml())
    const result = await convertOdpToPptx(await zip.generateAsync({ type: 'uint8array' }))
    const slide1 = (await readPptxEntry(result.pptx, 'ppt/slides/slide1.xml')) ?? ''
    expect(slide1).toContain('Template slide')
  })

  it('rejects non-ODF bytes and non-presentation bodies', async () => {
    await expect(convertOdpToPptx(new Uint8Array([9, 9, 9]))).rejects.toMatchObject({
      code: 'unreadable',
    })
    const zip = new JSZip()
    zip.file(
      'content.xml',
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>doc</text:p></office:text></office:body></office:document-content>',
    )
    await expect(
      convertOdpToPptx(await zip.generateAsync({ type: 'uint8array' })),
    ).rejects.toMatchObject({
      code: 'not-a-presentation',
    })
  })
})
