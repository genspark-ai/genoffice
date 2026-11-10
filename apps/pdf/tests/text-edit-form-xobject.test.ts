import { PDFDocument, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { SHARED_TEMPLATE_REASON } from '../src/main/form-text'
import { applyTextEdits, validateTextEdits } from '../src/main/text-edit'
import type { TextEditInput } from '../src/shared/ipc'

interface Item {
  str: string
  x: number
  y: number
}

/** First-page text items with their baseline origins, read back by pdf.js */
async function pageItems(bytes: Uint8Array, pageIndex = 0): Promise<Item[]> {
  const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await getDocument({ data: bytes.slice(), useSystemFonts: true }).promise
  try {
    const page = await doc.getPage(pageIndex + 1)
    const content = await page.getTextContent()
    return content.items
      .filter((i) => 'str' in i && i.str.trim())
      .map((i) => ({
        str: (i as { str: string }).str,
        x: (i as { transform: number[] }).transform[4]!,
        y: (i as { transform: number[] }).transform[5]!,
      }))
  } finally {
    await doc.loadingTask.destroy()
  }
}

const SIZE = 14
const FORM_X = 60
const FORM_Y = 300
const SCALE = 0.5

/** A page whose text lives only inside a Form XObject (an embedded page), painted
    with a translate+scale matrix like an FPDI/TCPDF template */
async function makeTemplateDoc(opts: { nested?: boolean; copies?: number; pages?: number } = {}) {
  const src = await PDFDocument.create()
  const srcPage = src.addPage([595, 842])
  const font = await src.embedFont(StandardFonts.Helvetica)
  srcPage.drawText('SECTOR', { x: 50, y: 700, size: SIZE, font })
  srcPage.drawText('CORTE DE CHAPA', { x: 200, y: 700, size: SIZE, font })
  srcPage.drawText('OPERARIO', { x: 50, y: 650, size: SIZE, font })
  let inner = await PDFDocument.load(await src.save({ useObjectStreams: false }))
  if (opts.nested) {
    const mid = await PDFDocument.create()
    const midPage = mid.addPage([595, 842])
    const emb = await mid.embedPage(inner.getPage(0))
    midPage.drawPage(emb, { x: 0, y: 0 })
    inner = await PDFDocument.load(await mid.save({ useObjectStreams: false }))
  }
  const doc = await PDFDocument.create()
  const emb = await doc.embedPage(inner.getPage(0))
  for (let p = 0; p < (opts.pages ?? 1); p++) {
    const page = doc.addPage([595, 842])
    for (let c = 0; c < (opts.copies ?? 1); c++) {
      page.drawPage(emb, { x: FORM_X, y: FORM_Y - c * 250, xScale: SCALE, yScale: SCALE })
    }
  }
  return doc.save({ useObjectStreams: false })
}

const px = (x: number) => FORM_X + x * SCALE
const py = (y: number) => FORM_Y + y * SCALE

const chapaEdit = (newText: string, extra: Partial<TextEditInput> = {}): TextEditInput => ({
  pageIndex: 0,
  rect: [px(198), py(696), px(200) + 110 * SCALE, py(700 + SIZE + 2)],
  oldText: 'CORTE DE CHAPA',
  newText,
  fontSize: SIZE * SCALE,
  ...extra,
})

describe('text inside Form XObjects', () => {
  it('validates and replaces a run drawn by a template', async () => {
    const bytes = await makeTemplateDoc()
    const before = await pageItems(bytes)
    expect(before.map((i) => i.str)).toEqual(['SECTOR', 'CORTE DE CHAPA', 'OPERARIO'])

    const [v] = await validateTextEdits(bytes, [chapaEdit('CORTE DE CHAPA')])
    expect(v!.reason).toBeNull()

    const { bytes: out, skipped } = await applyTextEdits(bytes, [chapaEdit('CORTE DE ACERO')])
    expect(skipped).toEqual([])
    const after = await pageItems(out)
    expect(after.map((i) => i.str)).toEqual(['SECTOR', 'OPERARIO', 'CORTE DE ACERO'])
    // Untouched template text keeps its place; the rewrite sits where the old run was
    const sector = after.find((i) => i.str === 'SECTOR')!
    expect(sector.x).toBeCloseTo(px(50), 0)
    expect(sector.y).toBeCloseTo(py(700), 0)
    const acero = after.find((i) => i.str === 'CORTE DE ACERO')!
    expect(acero.x).toBeCloseTo(px(200), 0)
    expect(acero.y).toBeCloseTo(py(700), 0)
  })

  it('deletes and moves template text', async () => {
    const bytes = await makeTemplateDoc()
    const gone = await applyTextEdits(bytes, [chapaEdit('')])
    expect(gone.skipped).toEqual([])
    expect((await pageItems(gone.bytes)).map((i) => i.str)).toEqual(['SECTOR', 'OPERARIO'])

    const moved = await applyTextEdits(bytes, [
      chapaEdit('CORTE DE CHAPA', { translate: [0, -40] }),
    ])
    expect(moved.skipped).toEqual([])
    const chapa = (await pageItems(moved.bytes)).find((i) => i.str === 'CORTE DE CHAPA')!
    expect(chapa.x).toBeCloseTo(px(200), 0)
    expect(chapa.y).toBeCloseTo(py(700) - 40, 0)
  })

  it('keeps stream order when several template runs are edited at once', async () => {
    const bytes = await makeTemplateDoc()
    const sector: TextEditInput = {
      pageIndex: 0,
      rect: [px(48), py(696), px(50) + 50 * SCALE, py(700 + SIZE + 2)],
      oldText: 'SECTOR',
      newText: 'AREA',
      fontSize: SIZE * SCALE,
    }
    const { bytes: out, skipped } = await applyTextEdits(bytes, [
      sector,
      chapaEdit('CORTE DE ACERO'),
    ])
    expect(skipped).toEqual([])
    expect((await pageItems(out)).map((i) => i.str)).toEqual(['OPERARIO', 'AREA', 'CORTE DE ACERO'])
  })

  it('edits text inside a nested template', async () => {
    const bytes = await makeTemplateDoc({ nested: true })
    const { bytes: out, skipped } = await applyTextEdits(bytes, [chapaEdit('CORTE DE ACERO')])
    expect(skipped).toEqual([])
    const after = await pageItems(out)
    expect(after.map((i) => i.str).sort()).toEqual(['CORTE DE ACERO', 'OPERARIO', 'SECTOR'])
    expect(after.filter((i) => i.str === 'CORTE DE CHAPA')).toEqual([])
    const operario = after.find((i) => i.str === 'OPERARIO')!
    expect(operario.x).toBeCloseTo(px(50), 0)
    expect(operario.y).toBeCloseTo(py(650), 0)
  })

  it('refuses to edit a template painted on more than one page', async () => {
    const bytes = await makeTemplateDoc({ pages: 2 })
    const [v] = await validateTextEdits(bytes, [chapaEdit('CORTE DE CHAPA')])
    expect(v!.reason).toBe(SHARED_TEMPLATE_REASON)
    const res = await applyTextEdits(bytes, [chapaEdit('CORTE DE ACERO')])
    expect(res.skipped.map((s) => s.reason)).toEqual([SHARED_TEMPLATE_REASON])
    expect(res.bytes).toBe(bytes)
    expect((await pageItems(bytes, 1)).map((i) => i.str)).toContain('CORTE DE CHAPA')
  })

  it('refuses to edit a template painted twice on the same page', async () => {
    const bytes = await makeTemplateDoc({ copies: 2 })
    const [v] = await validateTextEdits(bytes, [chapaEdit('CORTE DE CHAPA')])
    expect(v!.reason).toBe(SHARED_TEMPLATE_REASON)
  })
})
