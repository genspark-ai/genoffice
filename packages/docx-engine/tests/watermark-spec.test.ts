import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  parseDocx,
  pictureWatermarkPreviewImage,
  readSectionSettings,
  saveDocx,
} from '../src/index'
import { TINY_PNG_BASE64, buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'

async function headerXml(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const name = Object.keys(zip.files).find((n) => /^word\/header\d*\.xml$/.test(n))!
  return zip.file(name)!.async('string')
}

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>) =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))

describe('watermark spec', () => {
  it('writes the Watermarks-gallery sdt with the requested face, color and orientation', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const saved = await saveDocx(doc, originalOrder(doc), {
      watermark: {
        text: 'CONFIDENTIAL',
        fontFamily: 'Arial',
        colorHex: 'FF0000',
        opacity: 0.3,
        diagonal: false,
        bold: true,
      },
    })
    const xml = await headerXml(saved)
    expect(xml).toContain('<w:docPartGallery w:val="Watermarks"/>')
    expect(xml).toContain('fillcolor="#FF0000"')
    expect(xml).toContain('<v:fill opacity="0.3"/>')
    expect(xml).toContain('font-family:&quot;Arial&quot;')
    expect(xml).toContain('font-weight:bold')
    expect(xml).not.toContain('rotation:315')
    expect(xml.indexOf('<w:sdt>')).toBeLessThan(xml.indexOf('<v:textpath'))

    const reparsed = await parseDocx(saved)
    expect(reparsed.watermarkText).toBe('CONFIDENTIAL')
    const wm = reparsed.headerImages?.find((img) => img.wordArt)
    expect(wm?.wordArt).toMatchObject({
      text: 'CONFIDENTIAL',
      colorHex: 'FF0000',
      fontFamily: 'Arial',
      bold: true,
    })
    expect(wm?.rotationDeg ?? 0).toBe(0)
  })

  it('replaces the sdt-wrapped watermark on re-set and drops it on null', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const first = await saveDocx(doc, originalOrder(doc), { watermark: 'DRAFT' })
    const doc2 = await parseDocx(first)
    const second = await saveDocx(doc2, originalOrder(doc2), { watermark: { text: 'FINAL' } })
    const xml2 = await headerXml(second)
    expect(xml2.match(/<v:textpath[^>]*\bstring=/g)).toHaveLength(1)
    expect(xml2).toContain('string="FINAL"')
    expect(xml2).toContain('rotation:315')

    const doc3 = await parseDocx(second)
    const third = await saveDocx(doc3, originalOrder(doc3), { watermark: null })
    const xml3 = await headerXml(third)
    expect(xml3).not.toContain('<v:textpath')
    expect(xml3).not.toContain('<w:sdt>')
    expect((await parseDocx(third)).watermarkText).toBeNull()
  })

  describe('picture watermark', () => {
    const image = { base64: TINY_PNG_BASE64, mime: 'image/png' as const, widthPx: 96, heightPx: 48 }

    async function headerRels(bytes: Uint8Array): Promise<string> {
      const zip = await JSZip.loadAsync(bytes)
      const name = Object.keys(zip.files).find((n) =>
        /^word\/_rels\/header\d*\.xml\.rels$/.test(n),
      )!
      return zip.file(name)!.async('string')
    }

    it('writes a picture frame in the Watermarks sdt with its media part and header relationship', async () => {
      const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
      const saved = await saveDocx(doc, originalOrder(doc), { watermark: { image } })
      const xml = await headerXml(saved)
      expect(xml).toContain('<w:docPartGallery w:val="Watermarks"/>')
      expect(xml).toContain('<v:shape id="WordPictureWatermark1"')
      expect(xml).toContain('type="#_x0000_t75"')
      expect(xml).toContain(
        '<v:imagedata r:id="rId1" o:title="watermark" gain="19661f" blacklevel="22938f"/>',
      )
      expect(xml).toContain('mso-position-horizontal-relative:margin')
      // A4 with 1in margins: the 2:1 image fits the 451.3pt margin width
      expect(xml).toContain('width:451.3pt;height:225.65pt;')
      expect(/<w:hdr[^>]*xmlns:r=/.test(xml)).toBe(true)
      const rels = await headerRels(saved)
      expect(rels).toMatch(
        /<Relationship Id="rId1" Type="[^"]*\/image" Target="media\/aidocs\d+\.png"\/>/,
      )
      const zip = await JSZip.loadAsync(saved)
      expect(Object.keys(zip.files).some((n) => /^word\/media\/aidocs\d+\.png$/.test(n))).toBe(true)
      expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('Extension="png"')

      const reparsed = await parseDocx(saved)
      expect(reparsed.watermarkText).toBeNull()
      expect(reparsed.watermarkPicture).toEqual({
        rId: 'rId1',
        widthPt: 451.3,
        heightPt: 225.65,
        washout: true,
      })
      const img = reparsed.headerImages?.find((i) => i.floating)
      expect(img?.behind).toBe(true)
      expect(img?.dataUrl.startsWith('data:image/png;base64,')).toBe(true)
      expect(img?.washout?.gain).toBeCloseTo(0.3, 3)
    })

    it('scale and washout:false size the frame from the natural size without fading', async () => {
      const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
      const saved = await saveDocx(doc, originalOrder(doc), {
        watermark: { image, scale: 50, washout: false },
      })
      const xml = await headerXml(saved)
      // 96x48px = 72x36pt at 96dpi, halved
      expect(xml).toContain('width:36pt;height:18pt;')
      expect(xml).not.toContain('gain=')
      expect((await parseDocx(saved)).watermarkPicture?.washout).toBe(false)
    })

    it('replaces text with picture and back, dropping the stale image relationship', async () => {
      const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
      const text = await saveDocx(doc, originalOrder(doc), { watermark: 'DRAFT' })
      const doc2 = await parseDocx(text)
      const pic = await saveDocx(doc2, originalOrder(doc2), { watermark: { image } })
      const xml2 = await headerXml(pic)
      expect(xml2).not.toContain('<v:textpath')
      expect(xml2.match(/<w:docPartGallery w:val="Watermarks"\/>/g)).toHaveLength(1)
      expect(xml2).toContain('<v:imagedata r:id="rId1"')

      const doc3 = await parseDocx(pic)
      expect(doc3.watermarkPicture?.rId).toBe('rId1')
      const back = await saveDocx(doc3, originalOrder(doc3), { watermark: { text: 'FINAL' } })
      const xml3 = await headerXml(back)
      expect(xml3).not.toContain('<v:imagedata')
      expect(xml3).toContain('string="FINAL"')
      expect(await headerRels(back)).not.toContain(
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"',
      )

      const doc4 = await parseDocx(pic)
      const removed = await saveDocx(doc4, originalOrder(doc4), { watermark: null })
      const xml4 = await headerXml(removed)
      expect(xml4).not.toContain('<v:imagedata')
      expect(xml4).not.toContain('<w:sdt>')
      expect((await parseDocx(removed)).watermarkPicture).toBeNull()
    })

    it('auto-fit uses the page box the same save writes, not the original one', async () => {
      const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
      const landscape = {
        ...readSectionSettings(doc),
        pageWidth: 16838,
        pageHeight: 11906,
        orientation: 'landscape' as const,
        marginLeft: 720,
        marginRight: 720,
      }
      const saved = await saveDocx(doc, originalOrder(doc), {
        watermark: { image },
        section: landscape,
      })
      const xml = await headerXml(saved)
      // landscape box 769.9 x 451.3pt: the 2:1 image is bound by the width
      expect(xml).toContain('width:769.9pt;height:384.95pt;')
      const reparsed = await parseDocx(saved)
      expect(readSectionSettings(reparsed).orientation).toBe('landscape')
      expect(reparsed.watermarkPicture).toMatchObject({ widthPt: 769.9, heightPt: 384.95 })
    })

    it('the preview image matches what the saved file parses back as', async () => {
      const preview = pictureWatermarkPreviewImage({ image }, { widthPt: 451.3, heightPt: 645.9 })
      const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
      const saved = await saveDocx(doc, originalOrder(doc), { watermark: { image } })
      const parsed = (await parseDocx(saved)).headerImages!.find((i) => i.watermark)!
      expect(parsed).toMatchObject({
        floating: true,
        behind: true,
        watermark: true,
        posH: 'center',
        posV: 'center',
        posHRel: 'margin',
        posVRel: 'margin',
      })
      expect(preview.widthPx).toBe(parsed.widthPx)
      expect(preview.heightPx).toBe(parsed.heightPx)
      expect(preview.washout?.gain).toBeCloseTo(parsed.washout!.gain, 2)
      expect(preview.dataUrl).toBe(parsed.dataUrl)
      expect(pictureWatermarkPreviewImage({ image, washout: false }, null).washout).toBeUndefined()
    })

    it('a header text edit keeps the picture watermark and its relationship', async () => {
      const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
      const pic = await saveDocx(doc, originalOrder(doc), { watermark: { image } })
      const doc2 = await parseDocx(pic)
      const edited = await saveDocx(doc2, originalOrder(doc2), { header: { text: 'Chapter 1' } })
      const xml = await headerXml(edited)
      expect(xml).toContain('<v:imagedata r:id="rId1"')
      expect(xml).toContain('Chapter 1')
      expect(xml.indexOf('<w:sdt>')).toBeLessThan(xml.indexOf('Chapter 1'))
      const reparsed = await parseDocx(edited)
      expect(reparsed.headerText).toBe('Chapter 1')
      expect(reparsed.watermarkPicture?.rId).toBe('rId1')
    })
  })

  it("a first-page header text edit keeps that part's own watermark", async () => {
    const FIRST_HEADER =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
      ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">' +
      '<w:p><w:r><w:pict><v:shape id="WordPictureWatermark7" type="#_x0000_t75"' +
      ' style="position:absolute;margin-left:0;margin-top:0;width:200pt;height:100pt;z-index:-251658240;' +
      'mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;' +
      'mso-position-vertical-relative:margin"><v:imagedata r:id="rId1" o:title="logo" gain="19661f" blacklevel="22938f"/>' +
      '</v:shape></w:pict></w:r></w:p>' +
      '<w:p><w:r><w:t>Cover</w:t></w:r></w:p></w:hdr>'
    const bytes = await buildDocx({
      bodyXml: BODY,
      withImage: true,
      extraRels:
        '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: FIRST_HEADER,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
        {
          path: 'word/_rels/header1.xml.rels',
          xml:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
            '</Relationships>',
          contentType: 'application/vnd.openxmlformats-package.relationships+xml',
        },
      ],
      sectPrExtra: '<w:headerReference w:type="first" r:id="rId40"/><w:titlePg/>',
    })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(doc, originalOrder(doc), { headerFirst: { text: 'Title page' } })
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/header1.xml')!.async('string')
    expect(xml).toContain('<v:imagedata r:id="rId1"')
    expect(xml).toContain('Title page')
    expect(xml).not.toContain('Cover')
    expect(await zip.file('word/_rels/header1.xml.rels')!.async('string')).toContain('Id="rId1"')
  })
})
