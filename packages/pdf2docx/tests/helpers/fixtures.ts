/**
 * Fixture PDFs generated on the fly:
 * - Latin/image/scanned fixtures via pdf-lib (deterministic, always available)
 * - CJK fixtures via PDFium itself (pdf-lib's built-in fonts have no CJK; we
 *   embed a system CJK-capable font instead — tests skip when none exists)
 */
import { readFileSync } from 'node:fs'
import type { PdfiumModule } from '../../src'
import { encodeRgbaPng } from '../../src/extract'

// ── pdf-lib fixtures ──

export async function buildLatinPdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const bold = await doc.embedFont(StandardFonts.HelveticaBold)

  page.drawText('Sample Document', { x: 216, y: 720, size: 18, font: bold })

  const para1 = [
    'The quick brown fox jumps over the',
    'lazy dog while the cat watches from',
    'a warm windowsill.',
  ]
  para1.forEach((text, i) => {
    page.drawText(text, { x: 72, y: 660 - i * 14, size: 12, font })
  })
  const para2 = ['A second paragraph sits below the', 'first one after a wider gap.']
  para2.forEach((text, i) => {
    page.drawText(text, { x: 72, y: 600 - i * 14, size: 12, font })
  })
  return doc.save()
}

/** small red square as PNG bytes (dogfoods our own encoder through a real decoder) */
export function tinyPng(): Uint8Array {
  const px = new Uint8Array(8 * 8 * 4)
  for (let i = 0; i < 64; i++) px.set([220, 30, 30, 255], i * 4)
  return encodeRgbaPng(px, 8, 8)
}

export async function buildImagePdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('Figure below:', { x: 72, y: 700, size: 12, font })
  const image = await doc.embedPng(tinyPng())
  page.drawImage(image, { x: 236, y: 500, width: 140, height: 140 })
  return doc.save()
}

/** 8×8 solid red baseline JPEG (components 3, single DCTDecode filter) */
const RED_JPEG_B64 =
  '/9j/4AAQSkZJRgABAQAASABIAAD/4QBMRXhpZgAATU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAAACKADAAQAAAABAAAACAAAAAD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZjwCyBOmACZjs+EJ+/8AAEQgACAAIAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/bAEMAAQEBAQEBAgEBAgMCAgIDBAMDAwMEBQQEBAQEBQYFBQUFBQUGBgYGBgYGBgcHBwcHBwgICAgICQkJCQkJCQkJCf/bAEMBAQEBAgICBAICBAkGBQYJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCf/dAAQAAf/aAAwDAQACEQMRAD8A/Leiiiv8/wA/6+D/2Q=='

export function tinyJpeg(): Uint8Array {
  return Uint8Array.from(Buffer.from(RED_JPEG_B64, 'base64'))
}

/** plain opaque JPEG page (embedJpg → single DCTDecode filter, no mask) */
export async function buildJpegPdf(): Promise<Uint8Array> {
  const { PDFDocument } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const image = await doc.embedJpg(tinyJpeg())
  page.drawImage(image, { x: 236, y: 500, width: 140, height: 140 })
  return doc.save()
}

/** full-page wallpaper JPEG under several text lines (P9 B) */
export async function buildBgImagePdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const image = await doc.embedJpg(tinyJpeg())
  page.drawImage(image, { x: 0, y: 0, width: 612, height: 792 })
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('Title over the wallpaper', { x: 72, y: 700, size: 18, font })
  page.drawText('Body line one on the background.', { x: 72, y: 660, size: 12, font })
  page.drawText('Body line two keeps flowing.', { x: 72, y: 640, size: 12, font })
  return doc.save()
}

/**
 * Buried slide-template background (P16 B): junk text and a wallpaper image
 * hidden under a full-page white wash, with the real content drawn on top.
 * z-order: white base → junk text → wallpaper jpeg → white wash → real text.
 */
export async function buildBuriedTemplatePdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts, rgb } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawRectangle({ x: 0, y: 0, width: 612, height: 792, color: rgb(1, 1, 1) })
  page.drawText('template junk furniture line', { x: 72, y: 400, size: 12, font })
  const image = await doc.embedJpg(tinyJpeg())
  page.drawImage(image, { x: 0, y: 0, width: 612, height: 792 })
  page.drawRectangle({ x: 0, y: 0, width: 612, height: 792, color: rgb(1, 1, 1), opacity: 0.85 })
  page.drawText('Real title above the wash', { x: 72, y: 700, size: 18, font })
  page.drawText('Real body keeps flowing here.', { x: 72, y: 660, size: 12, font })
  return doc.save()
}

/**
 * Full-bleed photo mid-stack with a large-but-not-full-bleed content card
 * over it (P16 B guard): the card (~92% × 90% of the page) must NOT extend
 * the background stack — its text stays in the flow.
 */
export async function buildPhotoCardPdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts, rgb } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([960, 540])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('buried under the photo', { x: 100, y: 300, size: 12, font })
  const image = await doc.embedJpg(tinyJpeg())
  page.drawImage(image, { x: 0, y: 0, width: 960, height: 540 })
  page.drawRectangle({ x: 30, y: 54, width: 890, height: 486, color: rgb(1, 1, 1) })
  page.drawText('Card title stays in the flow', { x: 100, y: 400, size: 24, font })
  page.drawText('Card body line one keeps flowing.', { x: 100, y: 340, size: 14, font })
  return doc.save()
}

/**
 * A DCT (JPEG) image whose transparency lives in a separate grayscale SMask —
 * left half transparent, right half opaque (P9 A). pdf-lib's embedJpg cannot
 * attach masks, so the image XObject is built by hand.
 */
export async function buildMaskedJpegPdf(): Promise<Uint8Array> {
  const {
    PDFDocument,
    concatTransformationMatrix,
    drawObject,
    popGraphicsState,
    pushGraphicsState,
  } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const ctx = doc.context
  const mask = new Uint8Array(8 * 8)
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) mask[y * 8 + x] = x < 4 ? 0 : 255
  }
  const maskRef = ctx.register(
    ctx.stream(mask, {
      Type: 'XObject',
      Subtype: 'Image',
      Width: 8,
      Height: 8,
      ColorSpace: 'DeviceGray',
      BitsPerComponent: 8,
    }),
  )
  const imgRef = ctx.register(
    ctx.stream(tinyJpeg(), {
      Type: 'XObject',
      Subtype: 'Image',
      Width: 8,
      Height: 8,
      ColorSpace: 'DeviceRGB',
      BitsPerComponent: 8,
      Filter: 'DCTDecode',
      SMask: maskRef,
    }),
  )
  const name = page.node.newXObject('Image', imgRef)
  page.pushOperators(
    pushGraphicsState(),
    concatTransformationMatrix(140, 0, 0, 140, 236, 500),
    drawObject(name),
    popGraphicsState(),
  )
  return doc.save()
}

/**
 * Bordered 2-row × 3-col table drawn with pdf-lib lines + a shaded header cell.
 * The vertical border at x=328 exists only in the top row, so the bottom row
 * merges columns 1-2 into one cell. Grid: x 72|200|328|456, y 620|660|700.
 */
export async function buildTablePdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts, rgb } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)

  page.drawText('Before the table.', { x: 72, y: 730, size: 12, font })

  // header cell shading first, borders on top
  page.drawRectangle({ x: 72, y: 660, width: 128, height: 40, color: rgb(1, 0.8, 0) })

  const line = (x1: number, y1: number, x2: number, y2: number) =>
    page.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: 1 })
  for (const y of [700, 660, 620]) line(72, y, 456, y)
  for (const x of [72, 200, 456]) line(x, 620, x, 700)
  line(328, 660, 328, 700) // only the top row → bottom row merges columns 1-2

  const cell = (text: string, x: number, y: number) => page.drawText(text, { x, y, size: 11, font })
  cell('Name', 80, 674)
  cell('Qty', 208, 674)
  cell('Price', 336, 674)
  cell('Total', 80, 634)
  cell('wide merged cell', 208, 634)
  return doc.save()
}
/**
 * P4 style-mapping page: a plain line, a fill-highlighted line, an underlined
 * line and a struck-through line (shape widths follow the real text metrics
 * so the "rule extends past its text" separator guard stays quiet).
 */
export async function buildStyledPdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts, rgb } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const size = 12

  page.drawText('Plain opening line here.', { x: 72, y: 700, size, font })

  const hl = 'Yellow highlighted words'
  const hlWidth = font.widthOfTextAtSize(hl, size)
  page.drawRectangle({ x: 70, y: 657, width: hlWidth + 4, height: 16, color: rgb(1, 1, 0) })
  page.drawText(hl, { x: 72, y: 660, size, font })

  const ul = 'Underlined words here'
  const ulWidth = font.widthOfTextAtSize(ul, size)
  page.drawLine({ start: { x: 72, y: 638 }, end: { x: 72 + ulWidth, y: 638 }, thickness: 1 })
  page.drawText(ul, { x: 72, y: 640, size, font })

  const st = 'Struck through words'
  const stWidth = font.widthOfTextAtSize(st, size)
  page.drawLine({ start: { x: 72, y: 614 }, end: { x: 72 + stWidth, y: 614 }, thickness: 1 })
  page.drawText(st, { x: 72, y: 610, size, font })

  return doc.save()
}

/** P4 list page: an intro line, three bullet items, then a numbered sequence. */
export async function buildListPdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)

  page.drawText('Shopping list intro:', { x: 72, y: 710, size: 12, font })
  const bullets = ['apples and pears', 'bread with butter', 'milk in bottles']
  bullets.forEach((t, i) => {
    page.drawText(`• ${t}`, { x: 72, y: 686 - i * 14, size: 11, font })
  })
  const steps = ['first step here', 'second step done', 'third step ends']
  steps.forEach((t, i) => {
    page.drawText(`${i + 1}. ${t}`, { x: 72, y: 630 - i * 14, size: 11, font })
  })
  return doc.save()
}

/**
 * P4 vector-illustration page: a text line, then a dense cluster of stroked
 * ellipses (bezier curves — exactly what shape normalization ignores).
 */
export async function buildVectorArtPdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts, rgb } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)

  page.drawText('Chart on this page:', { x: 72, y: 720, size: 12, font })
  for (let i = 0; i < 8; i++) {
    page.drawEllipse({
      x: 150 + i * 12,
      y: 500 + (i % 3) * 20,
      xScale: 30,
      yScale: 18,
      borderColor: rgb(0.2, 0.4, 0.8),
      borderWidth: 1.5,
    })
  }
  page.drawText('And a closing paragraph well below the chart.', { x: 72, y: 380, size: 12, font })
  return doc.save()
}

/**
 * Borderless three-line (booktabs) table page: a paragraph, then 3 aligned
 * text rows (2 columns at x 72 / 300) framed by 3 horizontal rules only —
 * no vertical strokes, so the lattice pass can never claim it.
 */
export async function buildStreamTablePdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)

  page.drawText('Quarterly results follow:', { x: 72, y: 720, size: 12, font })

  const rows: Array<[string, string, number]> = [
    ['Region', 'Revenue', 674],
    ['North', '1200', 656],
    ['South', '890', 638],
  ]
  for (const [a, b, y] of rows) {
    page.drawText(a, { x: 72, y, size: 11, font })
    page.drawText(b, { x: 300, y, size: 11, font })
  }
  const rule = (y: number) =>
    page.drawLine({ start: { x: 68, y }, end: { x: 360, y }, thickness: 1 })
  rule(688) // top rule
  rule(668) // below the header
  rule(630) // bottom rule
  return doc.save()
}

/** Two balanced text columns (x 72 / 320), six lines each, column-major order. */
export async function buildTwoColumnPdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)

  const left = [
    'Left column begins here with a story',
    'about a fox that crossed the meadow',
    'every morning before the sun rose,',
    'looking for the shortest way to the',
    'river where the water ran clear and',
    'cold over smooth grey stones below.',
  ]
  const right = [
    'Right column starts with the baker',
    'setting out loaves at dawn, steam',
    'rising from the crusts while early',
    'customers queued along the narrow',
    'street, coins ready, waiting for a',
    'door that opened exactly at seven.',
  ]
  left.forEach((t, i) => page.drawText(t, { x: 72, y: 700 - i * 14, size: 10, font }))
  right.forEach((t, i) => page.drawText(t, { x: 320, y: 700 - i * 14, size: 10, font }))
  return doc.save()
}

/** A low title with big whitespace above and below (before_space chain case). */
export async function buildSpacedPdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const bold = await doc.embedFont(StandardFonts.HelveticaBold)

  page.drawText('Spaced Title', { x: 230, y: 640, size: 18, font: bold })
  const body = [
    'The body paragraph starts far below the title,',
    'after a deliberately large stretch of empty page',
    'that must come back as paragraph spacing.',
  ]
  body.forEach((t, i) => page.drawText(t, { x: 72, y: 500 - i * 14, size: 12, font }))
  return doc.save()
}

export async function buildScannedPdf(): Promise<Uint8Array> {
  const { PDFDocument } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const image = await doc.embedPng(tinyPng())
  page.drawImage(image, { x: 0, y: 0, width: 612, height: 792 })
  return doc.save()
}

// ── PDFium-generated CJK fixtures ──

const FPDF_FONT_TRUETYPE = 2

/** single-face sfnt fonts with CJK coverage (no .ttc — FPDFText_LoadFont takes raw sfnt) */
const CJK_FONT_PATHS = [
  '/System/Library/Fonts/Supplemental/Arial Unicode.ttf',
  'C:\\Windows\\Fonts\\arialuni.ttf',
  '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
]

let cjkFontCache: Buffer | null | undefined

export function cjkFontBytes(): Buffer | null {
  if (cjkFontCache === undefined) {
    cjkFontCache = null
    for (const p of CJK_FONT_PATHS) {
      try {
        if (p.endsWith('.ttc')) continue
        cjkFontCache = readFileSync(p)
        break
      } catch {
        /* try next */
      }
    }
  }
  return cjkFontCache
}

/** functions only the fixture builder needs (kept out of the package's interface) */
interface FixturePdfium extends PdfiumModule {
  _FPDF_CreateNewDocument(): number
  _FPDFPage_New(doc: number, index: number, width: number, height: number): number
  _FPDFText_LoadFont(doc: number, data: number, size: number, fontType: number, cid: number): number
  _FPDFPageObj_CreateTextObj(doc: number, font: number, size: number): number
  _FPDFText_SetText(obj: number, text: number): number
  _FPDFPageObj_Transform(
    obj: number,
    a: number,
    b: number,
    c: number,
    d: number,
    e: number,
    f: number,
  ): void
  _FPDFPage_InsertObject(page: number, obj: number): void
  _FPDFPage_GenerateContent(page: number): number
  _FPDFFont_Close(font: number): void
  _PDFiumExt_OpenFileWriter(): number
  _PDFiumExt_SaveAsCopy(doc: number, writer: number): number
  _PDFiumExt_GetFileWriterSize(writer: number): number
  _PDFiumExt_GetFileWriterData(writer: number, buf: number, size: number): number
  _PDFiumExt_CloseFileWriter(writer: number): void
}

export interface FixtureLine {
  text: string
  x: number
  y: number
  size?: number
}

/** Build a one-page PDF with the given text lines using an embedded CJK font. */
export function buildCjkPdf(
  pdfium: PdfiumModule,
  fontBytes: Buffer,
  lines: FixtureLine[],
): Uint8Array {
  const m = pdfium as FixturePdfium
  const doc = m._FPDF_CreateNewDocument()
  if (!doc) throw new Error('FPDF_CreateNewDocument failed')
  try {
    const page = m._FPDFPage_New(doc, 0, 612, 792)
    const fontPtr = m._malloc(fontBytes.length)
    m.HEAPU8.set(fontBytes, fontPtr)
    const font = m._FPDFText_LoadFont(doc, fontPtr, fontBytes.length, FPDF_FONT_TRUETYPE, 1)
    m._free(fontPtr)
    if (!font) throw new Error('FPDFText_LoadFont failed')

    for (const line of lines) {
      const obj = m._FPDFPageObj_CreateTextObj(doc, font, line.size ?? 12)
      const bytes = Buffer.from(`${line.text}\0`, 'utf16le')
      const textPtr = m._malloc(bytes.length)
      m.HEAPU8.set(bytes, textPtr)
      const ok = m._FPDFText_SetText(obj, textPtr)
      m._free(textPtr)
      if (!ok) throw new Error('FPDFText_SetText failed')
      m._FPDFPageObj_Transform(obj, 1, 0, 0, 1, line.x, line.y)
      m._FPDFPage_InsertObject(page, obj)
    }
    if (!m._FPDFPage_GenerateContent(page)) throw new Error('FPDFPage_GenerateContent failed')
    m._FPDF_ClosePage(page)

    const writer = m._PDFiumExt_OpenFileWriter()
    try {
      if (!m._PDFiumExt_SaveAsCopy(doc, writer)) throw new Error('SaveAsCopy failed')
      const size = m._PDFiumExt_GetFileWriterSize(writer)
      const buf = m._malloc(size)
      m._PDFiumExt_GetFileWriterData(writer, buf, size)
      const out = Uint8Array.from(m.HEAPU8.subarray(buf, buf + size))
      m._free(buf)
      return out
    } finally {
      m._PDFiumExt_CloseFileWriter(writer)
    }
  } finally {
    m._FPDF_CloseDocument(doc)
  }
}

/** wide-format scan: 25in-wide page (past Word's 22in ceiling), one page-covering image */
export async function buildWideScannedPdf(): Promise<Uint8Array> {
  const { PDFDocument } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([1800, 1012.5])
  const image = await doc.embedPng(tinyPng())
  page.drawImage(image, { x: 0, y: 0, width: 1800, height: 1012.5 })
  return doc.save()
}

/** full-page WHITE wash under real content + a stroked box: the wash must not
 * activate the background stack and swallow the vectors under/after it */
export async function buildWhiteWashPdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts, rgb } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  page.drawRectangle({ x: 0, y: 0, width: 612, height: 792, color: rgb(1, 1, 1) })
  page.drawRectangle({
    x: 72,
    y: 500,
    width: 200,
    height: 100,
    borderColor: rgb(0.1, 0.1, 0.1),
    borderWidth: 1,
  })
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('Content over a plain white wash.', { x: 72, y: 650, size: 12, font })
  return doc.save()
}

/** AcroForm page: two checkbox WIDGETS (one checked) beside text labels */
export async function buildCheckboxFormPdf(): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('Check one of the following options today.', { x: 72, y: 700, size: 12, font })
  page.drawText('Individual', { x: 92, y: 660, size: 12, font })
  page.drawText('Corporation', { x: 204, y: 660, size: 12, font })
  const form = doc.getForm()
  const checked = form.createCheckBox('opt.individual')
  checked.addToPage(page, { x: 74, y: 657, width: 12, height: 12 })
  checked.check()
  const unchecked = form.createCheckBox('opt.corporation')
  unchecked.addToPage(page, { x: 186, y: 657, width: 12, height: 12 })
  return doc.save()
}
