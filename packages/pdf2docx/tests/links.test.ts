/**
 * PDF link annotations → docx hyperlinks (genoffice#1096): URI actions become
 * external relationships, in-document GoTo destinations become anchors on a
 * per-page bookmark.
 */
import { describe, expect, it } from 'vitest'
import { convertPdfToDocx } from '../src'
import { buildSpans } from '../src/analyze/spans'
import type { Word } from '../src/analyze/words'
import type { PdfChar } from '../src/ir'
import { tinyPng } from './helpers/fixtures'
import { loadPdfium } from './helpers/wasm'

async function buildLinkPdf(): Promise<Uint8Array> {
  const { PDFDocument, PDFName, PDFString, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const p1 = doc.addPage([612, 792])
  const p2 = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const uriText = 'Visit the example site'
  const gotoText = 'Jump to the second page'
  p1.drawText('Plain heading text', { x: 72, y: 720, size: 14, font })
  p1.drawText(uriText, { x: 72, y: 680, size: 12, font })
  p1.drawText(gotoText, { x: 72, y: 640, size: 12, font })
  p2.drawText('Second page body', { x: 72, y: 700, size: 12, font })
  const uriW = font.widthOfTextAtSize(uriText, 12)
  const gotoW = font.widthOfTextAtSize(gotoText, 12)
  const uri = doc.context.register(
    doc.context.obj({
      Type: 'Annot',
      Subtype: 'Link',
      Rect: [72, 677, 72 + uriW, 692],
      Border: [0, 0, 0],
      A: { Type: 'Action', S: 'URI', URI: PDFString.of('https://example.com/docs') },
    }),
  )
  const goto = doc.context.register(
    doc.context.obj({
      Type: 'Annot',
      Subtype: 'Link',
      Rect: [72, 637, 72 + gotoW, 652],
      Border: [0, 0, 0],
      Dest: [p2.ref, 'XYZ', null, null, null],
    }),
  )
  p1.node.set(PDFName.of('Annots'), doc.context.obj([uri, goto]))
  return doc.save()
}

async function buildImageTargetPdf(): Promise<Uint8Array> {
  const { PDFDocument, PDFName, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const p1 = doc.addPage([612, 792])
  const p2 = doc.addPage([612, 792])
  const png = await doc.embedPng(tinyPng())
  p1.drawImage(png, { x: 72, y: 400, width: 300, height: 300 })
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const text = 'Back to the cover'
  p2.drawText(text, { x: 72, y: 700, size: 12, font })
  const w = font.widthOfTextAtSize(text, 12)
  const goto = doc.context.register(
    doc.context.obj({
      Type: 'Annot',
      Subtype: 'Link',
      Rect: [72, 697, 72 + w, 712],
      Border: [0, 0, 0],
      Dest: [p1.ref, 'XYZ', null, null, null],
    }),
  )
  p2.node.set(PDFName.of('Annots'), doc.context.obj([goto]))
  return doc.save()
}

async function buildBlankSectionTargetPdf(): Promise<Uint8Array> {
  const { PDFDocument, PDFName, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const p1 = doc.addPage([612, 792])
  const p2 = doc.addPage([792, 612])
  const p3 = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const text = 'See the blank landscape page'
  p1.drawText(text, { x: 72, y: 700, size: 12, font })
  p3.drawText('Third page body', { x: 72, y: 700, size: 12, font })
  const w = font.widthOfTextAtSize(text, 12)
  const goto = doc.context.register(
    doc.context.obj({
      Type: 'Annot',
      Subtype: 'Link',
      Rect: [72, 697, 72 + w, 712],
      Border: [0, 0, 0],
      Dest: [p2.ref, 'XYZ', null, null, null],
    }),
  )
  p1.node.set(PDFName.of('Annots'), doc.context.obj([goto]))
  return doc.save()
}

const charAt = (text: string, x: number, href?: string): PdfChar => ({
  code: text.codePointAt(0)!,
  text,
  box: { x0: x, x1: x + 6, y0: 100, y1: 110 },
  looseBox: { x0: x, x1: x + 6, y0: 100, y1: 110 },
  originX: x,
  originY: 100,
  angle: 0,
  fontSize: 12,
  fontWeight: 400,
  fontFamily: 'Helvetica',
  italic: false,
  color: '000000',
  isGenerated: false,
  isHyphen: false,
  script: 'latin',
  ...(href ? { href } : {}),
})

describe('link annotations', () => {
  it('splits a span where the link target changes', () => {
    const chars = [
      charAt('a', 0),
      charAt('b', 6, 'https://x.test/'),
      charAt('c', 12, 'https://x.test/'),
      charAt('d', 18),
    ]
    const words: Word[] = [{ chars, spaceBefore: false } as Word]
    const spans = buildSpans(words)
    expect(spans.map((s) => [s.text, s.href])).toEqual([
      ['a', undefined],
      ['bc', 'https://x.test/'],
      ['d', undefined],
    ])
  })

  it('emits w:hyperlink for URI links and page anchors for GoTo links', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildLinkPdf(), { pdfium })
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(result.docx)
    const xml = await zip.file('word/document.xml')!.async('string')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')

    const external = /<w:hyperlink r:id="(rId\d+)"[^>]*>([\s\S]*?)<\/w:hyperlink>/.exec(xml)
    expect(external).not.toBeNull()
    expect(external![2]).toContain('Visit the example site')
    expect(rels).toContain(
      `Id="${external![1]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/docs" TargetMode="External"`,
    )

    const internal = /<w:hyperlink w:anchor="(_pdfpage2)"[^>]*>([\s\S]*?)<\/w:hyperlink>/.exec(xml)
    expect(internal).not.toBeNull()
    expect(internal![2]).toContain('Jump to the second page')
    const bookmarkAt = xml.indexOf('w:name="_pdfpage2"')
    expect(bookmarkAt).toBeGreaterThan(-1)
    expect(xml.indexOf('Second page body')).toBeGreaterThan(bookmarkAt)
    expect(xml.indexOf('Plain heading text')).toBeLessThan(bookmarkAt)
    expect(xml.match(/<w:hyperlink\b/g)).toHaveLength(2)
  })

  it('anchors a GoTo target on an image-only first page instead of a later page', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildImageTargetPdf(), { pdfium })
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(result.docx)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:hyperlink w:anchor="_pdfpage1"')
    const bookmarkAt = xml.indexOf('w:name="_pdfpage1"')
    expect(bookmarkAt).toBeGreaterThan(-1)
    expect(bookmarkAt).toBeLessThan(xml.indexOf('<w:drawing'))
    expect(bookmarkAt).toBeLessThan(xml.indexOf('Back to the cover'))
  })

  it('anchors a GoTo target on a blank page that opens its own section', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildBlankSectionTargetPdf(), { pdfium })
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(result.docx)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:hyperlink w:anchor="_pdfpage2"')
    const paras = xml.split('<w:p>').filter((p) => p.includes('w:name="_pdfpage2"'))
    expect(paras).toHaveLength(1)
    expect(paras[0]).toContain('<w:pageBreakBefore/>')
    expect(paras[0]).not.toContain('Third page body')
    const bookmarkAt = xml.indexOf('w:name="_pdfpage2"')
    expect(bookmarkAt).toBeGreaterThan(xml.indexOf('See the blank landscape page'))
    expect(bookmarkAt).toBeLessThan(xml.indexOf('Third page body'))
    const third = xml.split('<w:p>').find((p) => p.includes('Third page body'))!
    expect(third).toContain('<w:pageBreakBefore/>')
  })
})
