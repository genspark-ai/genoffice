/**
 * Reflection as a text effect, a text box's direction, Convert Permanently of
 * Font Substitution, and the SVG a picture's blip names beside its bitmap.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  buildTextboxParagraphXml,
  inlinePictureDrawingXml,
  parseDocx,
  patchShapeStyles,
  replaceFontsInDocx,
  replaceFontsXml,
  saveDocx,
  type Run,
} from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const partXml = async (bytes: Uint8Array, part = 'word/document.xml'): Promise<string> =>
  (await (await JSZip.loadAsync(bytes)).file(part)!.async('string')) ?? ''

describe('w14:reflection', () => {
  const REFLECTION = { startAlpha: 0.5, endPos: 0.55, blurPt: 0.5, distPt: 0 }

  async function saved(runs: Run[], bodyXml = '<w:p><w:r><w:t>plain</w:t></w:r></w:p>') {
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const bytes = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs } },
    ])
    return { xml: await partXml(bytes), runs: (await parseDocx(bytes)).blocks[0].runs ?? [] }
  }

  it('is written on a new run between the shadow and the outline, and read back', async () => {
    const { xml, runs } = await saved([
      {
        text: 'mirrored',
        fxOwn: {
          shadow: { color: '000000', blurPt: 4, distPt: 3, dirDeg: 45 },
          reflection: REFLECTION,
          outline: { color: 'FF0000', widthPt: 1 },
        },
      },
    ])
    expect(xml).toMatch(/<\/w14:shadow><w14:reflection [^>]*\/><w14:textOutline /)
    expect(xml).toContain(
      '<w14:reflection w14:blurRad="6350" w14:stA="50000" w14:stPos="0" w14:endA="300" ' +
        'w14:endPos="55000" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" ' +
        'w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>',
    )
    expect(runs[0].textReflection).toEqual(REFLECTION)
  })

  const FILE_RUN =
    '<w:p><w:r><w:rPr><w14:reflection w14:blurRad="12700" w14:stA="38000" w14:stPos="0" ' +
    'w14:endA="0" w14:endPos="90000" w14:dist="25400" w14:dir="5400000" w14:fadeDir="5400000" ' +
    'w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/></w:rPr>' +
    '<w:t>kept</w:t></w:r></w:p>'

  it('is read from a file and stays byte for byte when the run is saved untouched', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: FILE_RUN }))
    const run = parsed.blocks[0].runs![0]
    expect(run.textReflection).toEqual({ startAlpha: 0.38, endPos: 0.9, blurPt: 1, distPt: 2 })
    const { xml } = await saved([{ ...run, text: 'kept still' }], FILE_RUN)
    expect(xml).toContain('w14:stA="38000"')
    expect(xml).toContain('w14:endA="0"')
  })

  it('is taken off a run from a file when the editor cleared it', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: FILE_RUN }))
    const { textReflection: _was, ...run } = parsed.blocks[0].runs![0]
    const { xml, runs } = await saved([{ ...run, fxOwn: { reflection: null } }], FILE_RUN)
    expect(xml).not.toContain('w14:reflection')
    expect(runs[0].textReflection).toBeUndefined()
  })
})

describe('text box direction', () => {
  const BOX = buildTextboxParagraphXml({ widthEmu: 1800000, heightEmu: 1080000, id: 7 })
  const firstBox = async (paragraphXml: string) => {
    const doc = await parseDocx(await buildDocx({ bodyXml: paragraphXml }))
    return doc.blocks.find((b) => b.textboxes?.length)?.textboxes?.[0]
  }

  it('is written as wps:bodyPr vert, replaced, and removed for horizontal', async () => {
    expect(BOX).toContain('<wps:bodyPr/>')
    const stacked = patchShapeStyles(BOX, [{ vert: 'wordArtVert' }])
    expect(stacked).toContain('<wps:bodyPr vert="wordArtVert"/>')
    expect((await firstBox(stacked))?.vert).toBe('wordArtVert')
    const turned = patchShapeStyles(stacked, [{ vert: 'vert270' }])
    expect(turned.match(/vert="/g)).toHaveLength(1)
    expect((await firstBox(turned))?.vert).toBe('vert270')
    const flat = patchShapeStyles(turned, [{ vert: null }])
    expect(flat).toBe(BOX)
    expect((await firstBox(flat))?.vert).toBeUndefined()
  })

  it('keeps the other attributes of a bodyPr a file wrote', () => {
    const fromFile = BOX.replace(
      '<wps:bodyPr/>',
      '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" anchor="ctr"><a:noAutofit/></wps:bodyPr>',
    )
    const out = patchShapeStyles(fromFile, [{ vert: 'eaVert' }])
    expect(out).toContain(
      '<wps:bodyPr vert="eaVert" rot="0" wrap="square" lIns="91440" anchor="ctr"><a:noAutofit/></wps:bodyPr>',
    )
  })

  it('reads the directions that stack letters or run lines left to right', async () => {
    for (const vert of ['wordArtVert', 'wordArtVertRtl', 'mongolianVert'] as const)
      expect((await firstBox(buildTextboxParagraphXml({ id: 9, vert })))?.vert).toBe(vert)
    const unknown = BOX.replace('<wps:bodyPr/>', '<wps:bodyPr vert="sideways"/>')
    expect((await firstBox(unknown))?.vert).toBeUndefined()
  })
})

describe('Convert Permanently', () => {
  const fonts = new Map([['old sans', 'Arial']])

  it('renames the font in every slot of w:rFonts and leaves other fonts and other attributes alone', () => {
    const xml =
      '<w:rPr><w:rFonts w:ascii="Old Sans" w:hAnsi="old sans" w:eastAsia="Kept" w:cs="Old Sans" w:hint="eastAsia"/></w:rPr>' +
      '<w:t>Old Sans</w:t><w:font w:name="Old Sans"/>'
    expect(replaceFontsXml(xml, fonts)).toBe(
      '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Kept" w:cs="Arial" w:hint="eastAsia"/></w:rPr>' +
        '<w:t>Old Sans</w:t><w:font w:name="Old Sans"/>',
    )
  })

  it('matches a name written with entities and escapes the name it writes', () => {
    const xml = '<w:rFonts w:ascii="A &amp; B"/>'
    expect(replaceFontsXml(xml, new Map([['a & b', 'C & D']]))).toBe(
      '<w:rFonts w:ascii="C &amp; D"/>',
    )
  })

  it('renames the typefaces of a theme', () => {
    const theme =
      '<a:majorFont><a:latin typeface="Old Sans"/><a:ea typeface=""/>' +
      '<a:font script="Jpan" typeface="Old Sans"/></a:majorFont>'
    expect(replaceFontsXml(theme, fonts)).toBe(
      '<a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/>' +
        '<a:font script="Jpan" typeface="Arial"/></a:majorFont>',
    )
  })

  it('rewrites the runs, the styles and the header of a document, and the result opens with the new font', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Old Sans" w:hAnsi="Old Sans"/></w:rPr><w:t>body</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="Fancy"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>',
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="Fancy"><w:name w:val="Fancy"/>' +
        '<w:rPr><w:rFonts w:ascii="Old Sans" w:hAnsi="Old Sans"/></w:rPr></w:style>',
      headerXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Old Sans"/></w:rPr><w:t>head</w:t></w:r></w:p>',
    })
    const out = await replaceFontsInDocx(source, { 'Old Sans': 'Arial' })
    for (const part of ['word/document.xml', 'word/styles.xml', 'word/header1.xml']) {
      const xml = await partXml(out, part)
      expect(xml, part).not.toContain('Old Sans')
      expect(xml, part).toContain('w:ascii="Arial"')
    }
    const parsed = await parseDocx(out)
    expect(parsed.blocks[0].runs?.[0].font ?? parsed.blocks[0].runs?.[0].fontAscii).toBe('Arial')
    expect(parsed.styles.get('Fancy')?.display?.font).toBe('Arial')
  })

  it('hands back the same bytes when the document names none of the fonts', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>plain</w:t></w:r></w:p>' })
    expect(await replaceFontsInDocx(source, { 'Old Sans': 'Arial' })).toBe(source)
    expect(await replaceFontsInDocx(source, {})).toBe(source)
  })
})

describe('the SVG of a picture', () => {
  const SVG =
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M2 2h20v20H2z"/></svg>'
  const SVG_B64 = Buffer.from(SVG).toString('base64')
  const drawing = inlinePictureDrawingXml({
    rId: 'rId10',
    svgRId: 'rId77',
    docPrId: 7,
    widthPx: 96,
    heightPx: 96,
  })
  const withSvg = (bodyXml: string) =>
    buildDocx({
      withImage: true,
      bodyXml,
      extraRels:
        '<Relationship Id="rId77" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/icon.svg"/>',
      binaryParts: [
        {
          path: 'word/media/icon.svg',
          base64: SVG_B64,
          extension: 'svg',
          contentType: 'image/svg+xml',
        },
      ],
    })

  it('is handed to the display beside the bitmap of a picture paragraph', async () => {
    const parsed = await parseDocx(await withSvg(`<w:p><w:r>${drawing}</w:r></w:p>`))
    const block = parsed.blocks.find((b) => b.type === 'image')!
    expect(block.imageDataUrl).toBe(`data:image/png;base64,${TINY_PNG_BASE64}`)
    expect(block.imageVectorUrl).toBe(`data:image/svg+xml;base64,${SVG_B64}`)
  })

  it('is handed to the display for a picture that shares its paragraph with text', async () => {
    const parsed = await parseDocx(
      await withSvg(`<w:p><w:r><w:t>text </w:t></w:r><w:r>${drawing}</w:r></w:p>`),
    )
    const image = parsed.blocks[0].runs?.find((run) => run.image)?.image
    expect(image?.vectorUrl).toBe(`data:image/svg+xml;base64,${SVG_B64}`)
  })

  it('is absent for a picture whose blip names none', async () => {
    const plain = inlinePictureDrawingXml({ rId: 'rId10', docPrId: 7, widthPx: 96, heightPx: 96 })
    const parsed = await parseDocx(
      await buildDocx({ withImage: true, bodyXml: `<w:p><w:r>${plain}</w:r></w:p>` }),
    )
    expect(parsed.blocks.find((b) => b.type === 'image')?.imageVectorUrl).toBeUndefined()
  })

  it('is replaced together with the bitmap when a replacement brings its own SVG', async () => {
    const parsed = await parseDocx(await withSvg(`<w:p><w:r>${drawing}</w:r></w:p>`))
    const block = parsed.blocks.find((b) => b.type === 'image')!
    const red = SVG.replace('<svg ', '<svg fill="#FF0000" ')
    const png =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
    const bytes = await saveDocx(parsed, [
      {
        kind: 'xml',
        xml: block.originalXml!,
        replaceImage: {
          base64: png,
          mime: 'image/png',
          svgBase64: Buffer.from(red).toString('base64'),
        },
      },
    ])
    const reopened = await parseDocx(bytes)
    const again = reopened.blocks.find((b) => b.type === 'image')!
    expect(again.imageDataUrl).toBe(`data:image/png;base64,${png}`)
    expect(again.imageVectorUrl).toBe(
      `data:image/svg+xml;base64,${Buffer.from(red).toString('base64')}`,
    )
    expect((await partXml(bytes)).match(/svgBlip /g)).toHaveLength(1)
  })

  it('comes with a replacement to a picture that had only a bitmap, whatever its blip holds', async () => {
    const plain = inlinePictureDrawingXml({ rId: 'rId10', docPrId: 7, widthPx: 96, heightPx: 96 })
    const blip = '<a:blip r:embed="rId10"/>'
    expect(plain).toContain(blip)
    const svgBase64 = Buffer.from(SVG).toString('base64')
    for (const inside of [
      blip,
      '<a:blip r:embed="rId10"><a:alphaModFix amt="50000"/></a:blip>',
      '<a:blip r:embed="rId10"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"/></a:extLst></a:blip>',
    ]) {
      const parsed = await parseDocx(
        await buildDocx({
          withImage: true,
          bodyXml: `<w:p><w:r>${plain.replace(blip, inside)}</w:r></w:p>`,
        }),
      )
      const block = parsed.blocks.find((b) => b.type === 'image')!
      const bytes = await saveDocx(parsed, [
        {
          kind: 'xml',
          xml: block.originalXml!,
          replaceImage: { base64: TINY_PNG_BASE64, mime: 'image/png', svgBase64 },
        },
      ])
      const xml = await partXml(bytes)
      expect(xml.match(/svgBlip /g), inside).toHaveLength(1)
      expect(xml.match(/<a:extLst>/g), inside).toHaveLength(1)
      // what the blip held stays, before the extension list
      if (inside.includes('alphaModFix'))
        expect(xml).toMatch(/<a:alphaModFix amt="50000"\/><a:extLst>/)
      if (inside.includes('28A0092B'))
        expect(xml).toContain('{28A0092B-C50C-407E-A947-70E740481C1C}')
      const again = (await parseDocx(bytes)).blocks.find((b) => b.type === 'image')!
      expect(again.imageVectorUrl, inside).toBe(`data:image/svg+xml;base64,${svgBase64}`)
    }
  })

  it('is dropped, as before, when the replacement is a bitmap alone', async () => {
    const parsed = await parseDocx(await withSvg(`<w:p><w:r>${drawing}</w:r></w:p>`))
    const block = parsed.blocks.find((b) => b.type === 'image')!
    const bytes = await saveDocx(parsed, [
      {
        kind: 'xml',
        xml: block.originalXml!,
        replaceImage: { base64: TINY_PNG_BASE64, mime: 'image/png' },
      },
    ])
    expect(await partXml(bytes)).not.toContain('svgBlip')
  })
})
