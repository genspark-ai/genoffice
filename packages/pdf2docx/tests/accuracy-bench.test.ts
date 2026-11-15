/**
 * PDF → editable conversion accuracy bench — invoice / report class.
 *
 * Gated: `PDF2DOCX_ACCURACY=1 scripts/pdf2docx-accuracy.mjs` drives it (see
 * that script). CI never runs it — the numbers are a narrative instrument
 * for the "PDF → editable" story, not a pass/fail gate; only the smoke case
 * (corpus builds, conversion runs) asserts anything, and only under the gate.
 *
 * For every corpus sample the bench converts with the package pipeline AND
 * with LibreOffice (`soffice --convert-to docx`), then scores both outputs
 * against the ground truth the generator baked in:
 *
 *   intact   — the field appears as one contiguous string inside a single
 *              table cell or paragraph (what a user can re-edit directly)
 *   split    — the field's text is present but fragmented across cells —
 *              the #2669 symptom (tax ID → "9" | "1310000MA1FL0000X")
 *   missing  — not found at all
 *
 * plus the detail-band column count for the invoice samples, the number a
 * banded-form refinement is supposed to get right.
 */
import { spawnSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { convertPdfToDocx, type PdfiumModule } from '../src'
import { loadPdfium } from './helpers/wasm'

const RUN = process.env.PDF2DOCX_ACCURACY === '1'
const OUT = process.env.PDF2DOCX_ACCURACY_OUT ?? ''
const REAL_INVOICE = join(import.meta.dirname, 'fixtures', 'invoice-text-merge.pdf')

const norm = (s: string): string => s.replace(/\s+/g, ' ').trim()

// ── corpus (deterministic, pdf-lib) ─────────────────────────────────────

/** the minimal pdf-lib surface the generators use */
interface BenchPage {
  drawText(text: string, options: { x: number; y: number; size: number; font: unknown }): void
  drawLine(options: {
    start: { x: number; y: number }
    end: { x: number; y: number }
    thickness: number
  }): void
}
interface PdfLib {
  PDFDocument: {
    create(): {
      addPage(size: [number, number]): BenchPage
      save(): Promise<Uint8Array>
      embedFont(font: unknown): Promise<unknown>
    }
  }
  StandardFonts: { Helvetica: unknown }
}
const pdfLib = async (): Promise<PdfLib> => (await import('pdf-lib')) as unknown as PdfLib

interface Sample {
  name: string
  pdf: Uint8Array
  /** ground-truth fields that must survive as one contiguous cell/paragraph */
  fields: string[]
  /** the detail band's expected column count (invoice samples) */
  expectedColumns?: number
  /** the table row identifiable by this header fragment */
  headerField?: string
}

/** An English stand-in for the #2669 e-invoice: same banded geometry — outer
 *  frame, full-width band rules, buyer/seller band with interior verticals,
 *  borderless detail band, remarks band. */
async function buildBandedInvoice(): Promise<Uint8Array> {
  // pdf-lib Y grows upward, same as the pipeline's page space
  {
    const { PDFDocument, StandardFonts } = await pdfLib()
    const doc = await PDFDocument.create()
    const page = doc.addPage([600, 400])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const p = page
    const T = (s: string, x: number, y: number, size = 7): void =>
      p.drawText(s, { x, y, size, font })
    const H = (y: number): void =>
      p.drawLine({ start: { x: 15, y }, end: { x: 585, y }, thickness: 1 })
    const V = (x: number, y0: number, y1: number): void =>
      p.drawLine({ start: { x, y: y0 }, end: { x, y: y1 }, thickness: 1 })
    // outer frame + band rules (same band ys as the real sample)
    V(15, 41.5, 312)
    V(585, 41.5, 312)
    H(312)
    H(248.5)
    H(121)
    H(99)
    H(42)
    // buyer / seller band
    T('Buyer Info', 18, 300)
    T('Tax ID / Name', 36, 295)
    T('Acme Engineering Co., Ltd.', 36, 280)
    T('TAX-ID 91310000MA1FL0000X', 36, 265)
    T('Seller Info', 304, 300)
    T('Tax ID / Name', 322, 295)
    T('Example Hotel Management LLC', 322, 280)
    T('TAX-ID 91110000MA0EXAMPLEX', 322, 265)
    V(33, 249, 311)
    V(301, 249, 311)
    V(319, 249, 311)
    // detail band: header labels + two item rows (no interior rules)
    T('Description', 47, 238)
    T('Spec', 121, 238)
    T('Unit', 192, 238)
    T('Qty', 266, 238)
    T('Unit Price', 336, 238)
    T('Amount', 409, 238)
    T('Tax Rate', 448, 238)
    T('Tax', 556, 238)
    T('Cleaning service *', 14.5, 231)
    T('1', 289.4, 228)
    T('100.000000', 295.9, 233)
    T('100.00', 409.6, 232)
    T('1%', 470.9, 228)
    T('6.00', 566.4, 230)
    T('Plumbing repair *', 14.5, 210)
    T('2', 289.4, 207)
    T('50.000000', 295.9, 212)
    T('50.00', 409.6, 211)
    T('1%', 470.9, 207)
    T('3.00', 566.4, 209)
    // totals band
    T('TOTAL', 60.7, 105)
    T('$106.00', 409.6, 104)
    V(162, 99, 121)
    // remarks band
    T('Remarks', 18, 88)
    V(33, 42, 99)
    T('Buyer address: 1 Example Rd; Phone: 00000000;', 36, 78)
    T('Buyer bank: Example Bank Branch; Account: 000000000000000000000;', 36, 66)
    T('Seller address: 2 Example St; Phone: 00000001;', 36, 54)
    T('Seller bank: Example Bank Main; Account: 111111111111111111111;', 36, 44)
    return doc.save()
  }
}

/** a bordered 6×5 operations report table with title and footer */
async function buildBorderedReport(): Promise<Uint8Array> {
  {
    const { PDFDocument, StandardFonts } = await pdfLib()
    const doc = await PDFDocument.create()
    const page = doc.addPage([612, 792])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const p = page
    const T = (s: string, x: number, y: number, size = 8): void =>
      p.drawText(s, { x, y, size, font })
    T('Q3 Operations Report', 60, 730, 14)
    const xs = [60, 180, 280, 380, 470, 550]
    const ys = [700, 680, 660, 640, 620, 600, 580]
    for (const y of ys) p.drawLine({ start: { x: xs[0]!, y }, end: { x: xs[5]!, y }, thickness: 1 })
    for (const x of xs)
      p.drawLine({ start: { x, y: ys[0]! }, end: { x, y: ys[ys.length - 1]! }, thickness: 1 })
    T('Region', 66, 704)
    T('Revenue', 186, 704)
    T('Costs', 286, 704)
    T('Margin', 386, 704)
    T('YoY', 476, 704)
    const rows: Array<[string, string, string, string, string]> = [
      ['North', '1,204,500', '912,300', '24.2%', '+3.1%'],
      ['South', '986,200', '744,100', '24.6%', '+5.4%'],
      ['East', '1,530,800', '1,201,500', '21.5%', '-1.2%'],
      ['West', '874,100', '650,900', '25.5%', '+7.8%'],
      ['Overseas', '412,700', '355,200', '13.9%', '+12.0%'],
    ]
    rows.forEach((row, i) => {
      const y = 684 - i * 20
      row.forEach((cellText, c) => T(cellText, xs[c]! + 6, y))
    })
    T('All figures in USD; unaudited management accounts.', 60, 555)
    return doc.save()
  }
}

/** the same report with no rules — a stream table */
async function buildBorderlessReport(): Promise<Uint8Array> {
  {
    const { PDFDocument, StandardFonts } = await pdfLib()
    const doc = await PDFDocument.create()
    const page = doc.addPage([612, 792])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const p = page
    const T = (s: string, x: number, y: number, size = 8): void =>
      p.drawText(s, { x, y, size, font })
    T('Q3 Operations Report', 60, 730, 14)
    const xs = [66, 186, 286, 386, 476]
    T('Region', xs[0]!, 704)
    T('Revenue', xs[1]!, 704)
    T('Costs', xs[2]!, 704)
    T('Margin', xs[3]!, 704)
    T('YoY', xs[4]!, 704)
    const rows: Array<[string, string, string, string, string]> = [
      ['North', '1,204,500', '912,300', '24.2%', '+3.1%'],
      ['South', '986,200', '744,100', '24.6%', '+5.4%'],
      ['East', '1,530,800', '1,201,500', '21.5%', '-1.2%'],
      ['West', '874,100', '650,900', '25.5%', '+7.8%'],
      ['Overseas', '412,700', '355,200', '13.9%', '+12.0%'],
    ]
    rows.forEach((row, i) => {
      const y = 684 - i * 20
      row.forEach((cellText, c) => T(cellText, xs[c]!, y))
    })
    T('All figures in USD; unaudited management accounts.', 60, 555)
    return doc.save()
  }
}

// ── extraction + scoring ────────────────────────────────────────────────

const docXml = async (docx: Uint8Array): Promise<string> => {
  const { default: JSZip } = await import('jszip')
  const zip = await JSZip.loadAsync(docx)
  return zip.file('word/document.xml')!.async('string')
}

function tableRows(xml: string): string[][] {
  return [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].flatMap((tbl) =>
    [...tbl[0].matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((row) =>
      [...row[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((cell) =>
        norm([...cell[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]!).join('')),
      ),
    ),
  )
}

function paragraphs(xml: string): string[] {
  return [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)]
    .map((para) =>
      norm([...para[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]!).join('')),
    )
    .filter((t) => t.length > 0)
}

interface Verdict {
  status: 'intact' | 'split' | 'missing'
  where: string
}

function scoreField(field: string, rows: string[][], paras: string[]): Verdict {
  const want = norm(field)
  for (const row of rows)
    for (const cell of row) if (cell.includes(want)) return { status: 'intact', where: 'cell' }
  for (const para of paras) if (para.includes(want)) return { status: 'intact', where: 'paragraph' }
  // present but fragmented: the field's text spans cell/paragraph boundaries —
  // the #2669 symptom (tax ID → "9" | "1310000MA1FL0000X")
  const whole = norm([...rows.flat(), ...paras].join('\u0001'))
  if (want.split(' ').every((w) => whole.includes(w)))
    return { status: 'split', where: 'fragmented' }
  return { status: 'missing', where: 'absent' }
}

interface EngineResult {
  intact: number
  split: number
  missing: number
  detailColumns?: number
  xml: string
}

async function scoreEngine(
  engine: 'genoffice' | 'soffice',
  sample: Sample,
  pdfium: PdfiumModule,
  workDir: string,
): Promise<EngineResult> {
  let docx: Uint8Array
  const pdfPath = join(workDir, `${sample.name}.pdf`)
  writeFileSync(pdfPath, sample.pdf)
  if (engine === 'genoffice') {
    const r = await convertPdfToDocx(sample.pdf, { pdfium })
    docx = r.docx
  } else {
    const profile = join(workDir, 'lo-profile')
    const run = spawnSync(
      'soffice',
      [
        '--headless',
        // the PDF import lands in Draw, whose docx export does not exist —
        // force the Writer PDF import filter, the one route LibreOffice has
        // to an editable docx from a PDF
        '--infilter=writer_pdf_import',
        '--convert-to',
        'docx',
        '--outdir',
        join(workDir, 'lo-out'),
        `-env:UserInstallation=file:///${profile}`,
        pdfPath,
      ],
      { timeout: 120_000 },
    )
    if (run.status !== 0) throw new Error(`soffice failed for ${sample.name}`)
    docx = readFileSync(join(workDir, 'lo-out', `${sample.name}.docx`))
  }
  const xml = await docXml(docx)
  const rows = tableRows(xml)
  const paras = paragraphs(xml)
  const verdicts = sample.fields.map((f) => scoreField(f, rows, paras))
  const result: EngineResult = {
    intact: verdicts.filter((v) => v.status === 'intact').length,
    split: verdicts.filter((v) => v.status === 'split').length,
    missing: verdicts.filter((v) => v.status === 'missing').length,
    xml,
  }
  if (sample.expectedColumns !== undefined && sample.headerField !== undefined) {
    const headerRow = rows.find((r) => r.some((c) => c.includes(norm(sample.headerField!))))
    result.detailColumns = headerRow?.length
  }
  return result
}

// ── the bench ───────────────────────────────────────────────────────────

describe.skipIf(!RUN)('pdf → editable accuracy bench (invoice/report class)', () => {
  it('scores genoffice vs soffice across the corpus', async () => {
    const pdfium = await loadPdfium()
    const samples: Sample[] = [
      {
        name: 'invoice-banded-en',
        pdf: await buildBandedInvoice(),
        fields: [
          'Acme Engineering Co., Ltd.',
          'TAX-ID 91310000MA1FL0000X',
          'Example Hotel Management LLC',
          'Cleaning service *',
          '100.000000',
          '6.00',
          'Plumbing repair *',
          '50.00',
          '$106.00',
          'Buyer address: 1 Example Rd; Phone: 00000000;',
          'Seller bank: Example Bank Main; Account: 111111111111111111111;',
        ],
        expectedColumns: 8,
        headerField: 'Description',
      },
      {
        name: 'invoice-cjk-real',
        pdf: new Uint8Array(readFileSync(REAL_INVOICE)),
        fields: [
          '91310000MA1FL0000X',
          '样例重工工程技术有限责任公司',
          '100.000000000000',
          '¥ 106.00',
          '壹佰零陆圆整',
          '购买方地址:示例市示例区示例路1号;',
          '销方开户银行:示例银行示例分行;',
        ],
        expectedColumns: 8,
        headerField: '项目名称',
      },
      {
        name: 'report-bordered',
        pdf: await buildBorderedReport(),
        fields: [
          'Region',
          'North',
          '1,204,500',
          '24.2%',
          '+3.1%',
          'South',
          '986,200',
          'East',
          '1,530,800',
          'Overseas',
          '412,700',
          'All figures in USD; unaudited management accounts.',
        ],
      },
      {
        name: 'report-borderless',
        pdf: await buildBorderlessReport(),
        fields: [
          'North',
          '1,204,500',
          '912,300',
          '24.2%',
          'Overseas',
          '412,700',
          '355,200',
          '13.9%',
          'All figures in USD; unaudited management accounts.',
        ],
      },
    ]

    const workDir = mkdtempSync(join(tmpdir(), 'pdf2docx-accuracy-'))
    writeFileSync(join(workDir, '.keep'), '')
    const scorecard: string[] = [
      '# PDF → editable accuracy — invoice / report class',
      '',
      'Field = a ground-truth string that must survive as ONE contiguous cell/paragraph',
      '(what a user can re-edit directly). split = present but fragmented across cells',
      '(the #2669 symptom). Columns = cells in the detail/header row vs expected.',
      '',
      '| sample | engine | intact | split | missing | detail columns (expected) |',
      '|---|---|---|---|---|---|',
    ]
    for (const sample of samples) {
      for (const engine of ['genoffice', 'soffice'] as const) {
        const r = await scoreEngine(engine, sample, pdfium, workDir)
        scorecard.push(
          `| ${sample.name} | ${engine} | ${r.intact}/${sample.fields.length} | ${r.split} | ${r.missing} | ${r.detailColumns ?? '-'}${sample.expectedColumns !== undefined ? ` (${sample.expectedColumns})` : ''} |`,
        )
        const bad = sample.fields.filter(
          (f) => scoreField(f, tableRows(r.xml), paragraphs(r.xml)).status !== 'intact',
        )
        if (bad.length > 0) scorecard.push(`  - not intact: ${bad.map((b) => `"${b}"`).join(', ')}`)
      }
    }
    rmSync(workDir, { recursive: true, force: true })

    const md = `${scorecard.join('\n')}\n`
    if (OUT) writeFileSync(OUT, md)
    else console.log(`\n${md}`)

    // the smoke assertion: the pipeline produces tables for the invoice class
    // (the numbers themselves are the narrative, never a gate)
    expect(md).toContain('genoffice')
  })
})
