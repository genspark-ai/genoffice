/**
 * P27 silent-empty coverage: pages that USED to convert to nothing without a
 * single warning — flipped media boxes, searchable-scan OCR overlays, page-
 * covering transparent watermark images, and the content-lost pipeline guard.
 */
import { describe, expect, it } from 'vitest'
import { extractIrDocument } from '../src/pipeline'
import { extractPage, withPdfDocument } from '../src/extract'
import type { IrPage } from '../src/ir'
import { loadPdfium } from './helpers/wasm'

// ── raw PDF fixture builder (hand-assembled xref so exotic structures like
// flipped MediaBoxes and Tr-3 text are expressible) ──

function rawPdf(objects: string[]): Uint8Array {
  const header = '%PDF-1.4\n'
  let body = ''
  const offsets: number[] = []
  for (const obj of objects) {
    offsets.push(header.length + body.length)
    body += obj
  }
  const xrefPos = header.length + body.length
  let xref = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const off of offsets) xref += `${String(off).padStart(10, '0')} 00000 n \n`
  const trailer = `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefPos}\n%%EOF\n`
  return new TextEncoder().encode(header + body + xref + trailer)
}

function contentObj(num: number, stream: string): string {
  return `${num} 0 obj\n<< /Length ${stream.length} >>\nstream\n${stream}\nendstream\nendobj\n`
}

/** a page of Helvetica text; mediaBox is written verbatim (may be flipped) */
function textPdf(
  mediaBox: string,
  content: string,
  extraResources = '',
  extraObjects: string[] = [],
): Uint8Array {
  return rawPdf([
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    `3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [${mediaBox}] ` +
      `/Resources << /Font << /F1 4 0 R >> ${extraResources} >> /Contents 5 0 R >>\nendobj\n`,
    '4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n',
    contentObj(5, content),
    ...extraObjects,
  ])
}

/** 1×1 opaque gray JPEG-free raw image XObject */
function rawImageObj(num: number, extra = ''): string {
  return (
    `${num} 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 ` +
    `/ColorSpace /DeviceRGB /BitsPerComponent 8 ${extra}/Length 3 >>\nstream\n\x80\x80\x80\nendstream\nendobj\n`
  )
}

/** 1×1 fully-transparent SMask (alpha 0) */
function smaskObj(num: number): string {
  return (
    `${num} 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 ` +
    `/ColorSpace /DeviceGray /BitsPerComponent 8 /Length 1 >>\nstream\n\x00\nendstream\nendobj\n`
  )
}

const BODY_LINES = [
  'The quick brown fox jumps over the lazy dog near the river bank.',
  'A second line of body text keeps the page clearly non-trivial here.',
  'Third line so the guard thresholds are comfortably exceeded today.',
]

function bodyContent(prefix = ''): string {
  let s = 'BT /F1 12 Tf\n'
  BODY_LINES.forEach((line, i) => {
    s += `1 0 0 1 72 ${700 - i * 20} Tm ${prefix}(${line}) Tj\n`
  })
  return s + 'ET\n'
}

function irText(pages: IrPage[]): string {
  let out = ''
  const walk = (b: IrPage['blocks'][number]): void => {
    if (b.kind === 'text') {
      for (const l of b.lines) for (const s of l.spans) out += s.text
    } else if (b.kind === 'table') {
      for (const r of b.rows) for (const c of r) for (const tb of c.blocks) walk(tb)
    }
  }
  for (const p of pages) for (const b of p.blocks) walk(b)
  return out
}

describe('P27 A1: flipped media boxes', () => {
  it('recovers text when the MediaBox top is negative (top < bottom)', async () => {
    const pdfium = await loadPdfium()
    // chars live at negative y in author space; the origin is the min corner
    let content = 'BT /F1 12 Tf\n'
    BODY_LINES.forEach((line, i) => {
      content += `1 0 0 1 72 ${-(92 + i * 20)} Tm (${line}) Tj\n`
    })
    content += 'ET\n'
    const pdf = textPdf('0 0 612 -792', content)
    const { irPages, pageResults } = extractIrDocument(pdf, { pdfium })
    expect(irText(irPages)).toContain('quick brown fox')
    expect(pageResults[0]!.status).toBe('ok')
  })

  it('recovers text when bottom/top are swapped (MediaBox [0 792 612 0])', async () => {
    const pdfium = await loadPdfium()
    const pdf = textPdf('0 792 612 0', bodyContent())
    const { irPages } = extractIrDocument(pdf, { pdfium })
    expect(irText(irPages)).toContain('quick brown fox')
  })
})

describe('P27 A2: searchable-scan OCR overlay', () => {
  it('flips the invisible OCR layer visible over a page-covering image', async () => {
    const pdfium = await loadPdfium()
    const content =
      'q 612 0 0 792 0 0 cm /Im1 Do Q\n' +
      // render mode 3 = invisible: the classic OCR text layer
      bodyContent('3 Tr ')
    const pdf = textPdf('0 0 612 792', content, '/XObject << /Im1 6 0 R >> ', [rawImageObj(6)])
    const { irPages, warnings } = extractIrDocument(pdf, { pdfium })
    expect(irText(irPages)).toContain('quick brown fox')
    expect(warnings.some((w) => w.includes('OCR text layer recovered'))).toBe(true)
    // the scan bitmap must NOT ride along (it would paint the words twice)
    expect(irPages[0]!.blocks.some((b) => b.kind === 'image')).toBe(false)
  })

  it('leaves small invisible runs alone (hidden formatting marks, P11 C)', async () => {
    const pdfium = await loadPdfium()
    const content =
      'q 612 0 0 792 0 0 cm /Im1 Do Q\n' +
      bodyContent() +
      'BT /F1 12 Tf 3 Tr 1 0 0 1 72 100 Tm (hidden) Tj ET\n'
    const pdf = textPdf('0 0 612 792', content, '/XObject << /Im1 6 0 R >> ', [rawImageObj(6)])
    const { warnings } = extractIrDocument(pdf, { pdfium })
    expect(warnings.some((w) => w.includes('OCR text layer recovered'))).toBe(false)
  })
})

describe('P27 A3: page-covering transparent watermark', () => {
  it('does not bury body text under a mostly-transparent top image', async () => {
    const pdfium = await loadPdfium()
    // text first, then a full-page transparent watermark image ON TOP
    const content = bodyContent() + 'q 612 0 0 792 0 0 cm /Im1 Do Q\n'
    const pdf = textPdf('0 0 612 792', content, '/XObject << /Im1 6 0 R >> ', [
      rawImageObj(6, '/SMask 7 0 R '),
      smaskObj(7),
    ])
    const { irPages, pageResults } = extractIrDocument(pdf, { pdfium })
    expect(irText(irPages)).toContain('quick brown fox')
    expect(pageResults[0]!.status).toBe('ok')
  })
})

describe('P27 A4: content-lost guard', () => {
  it('degrades + warns instead of shipping a silently empty page', async () => {
    const pdfium = await loadPdfium()
    // every char parked far off-canvas: extraction keeps none of the body
    let content = 'BT /F1 12 Tf\n'
    BODY_LINES.forEach((line, i) => {
      content += `1 0 0 1 -2000 ${700 - i * 20} Tm (${line}) Tj\n`
    })
    content += 'ET\n'
    const pdf = textPdf('0 0 612 792', content)
    const { warnings, pageResults } = extractIrDocument(pdf, { pdfium })
    expect(pageResults[0]!.status).toBe('degraded')
    expect(pageResults[0]!.reason).toBe('content-lost')
    expect(warnings.some((w) => w.includes('content could not be recovered'))).toBe(true)
  })

  it('stays quiet on genuinely blank pages', async () => {
    const pdfium = await loadPdfium()
    const pdf = textPdf('0 0 612 792', '')
    const { warnings, pageResults } = extractIrDocument(pdf, { pdfium })
    expect(pageResults[0]!.status).toBe('ok')
    expect(warnings).toEqual([])
  })
})

describe('P27: zero-page documents', () => {
  it('rejects a document with no readable pages instead of emitting empty output', async () => {
    const pdfium = await loadPdfium()
    const pdf = rawPdf([
      '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
      '2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n',
    ])
    expect(() => extractIrDocument(pdf, { pdfium })).toThrowError(
      expect.objectContaining({ name: 'PdfLoadError', code: 'corrupt' }),
    )
  })
})

describe('P27: extractPage flag plumbing', () => {
  it('marks ocrTextRecovered on the extracted page', async () => {
    const pdfium = await loadPdfium()
    const content = 'q 612 0 0 792 0 0 cm /Im1 Do Q\n' + bodyContent('3 Tr ')
    const pdf = textPdf('0 0 612 792', content, '/XObject << /Im1 6 0 R >> ', [rawImageObj(6)])
    const flag = withPdfDocument(pdfium, pdf, (doc) => extractPage(pdfium, doc, 0).ocrTextRecovered)
    expect(flag).toBe(true)
  })
})

describe('P27 T5: /Rotate page normalization', () => {
  it('extracts upright display-space text from a /Rotate 90 page', async () => {
    // landscape content in a portrait MediaBox, displayed via /Rotate 90:
    // text is drawn at +90° in user space so it reads upright when rotated
    const objects = [
      '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
      '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
      '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Rotate 90 ' +
        '/Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>\nendobj\n',
      '4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n',
    ]
    let content = 'BT /F1 12 Tf\n'
    BODY_LINES.forEach((line, i) => {
      // Tm = rotate +90° (cos=0 sin=1), advancing "lines" along +x
      content += `0 1 -1 0 ${100 + i * 20} 100 Tm (${line}) Tj\n`
    })
    content += 'ET\n'
    const pdf = rawPdf([...objects, contentObj(5, content)])
    const pdfium = await loadPdfium()
    const page = withPdfDocument(pdfium, pdf, (doc: number) => extractPage(pdfium, doc, 0))
    // display dims: 792×612 (rotated)
    expect(Math.round(page.widthPt)).toBe(792)
    expect(Math.round(page.heightPt)).toBe(612)
    expect(page.degraded).toBe(false)
    const body = page.chars.filter((c) => !c.isGenerated && c.code > 0x20)
    expect(body.length).toBeGreaterThan(100)
    // normalized: glyphs read as horizontal in display space
    const angled = body.filter((c) => Math.abs(c.angle) > 0.26)
    expect(angled.length / body.length).toBeLessThan(0.05)
    // and the text flows left→right inside the display page box
    const first = body[0]!
    expect(first.box.x0).toBeGreaterThanOrEqual(0)
    expect(first.box.x1).toBeLessThanOrEqual(792)
  })
})
