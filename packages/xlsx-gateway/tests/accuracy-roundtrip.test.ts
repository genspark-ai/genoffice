/**
 * XLSX round-trip fidelity bench — styled workbook class.
 *
 * Gated: `XLSX_ROUNDTRIP_ACCURACY=1 scripts/roundtrip-accuracy.mjs xlsx`
 * drives it (see that script). CI never runs it — the numbers are a narrative
 * instrument for the "open what Excel wrote, edit, save, nothing lost" story,
 * not a pass/fail gate; the smoke case (corpus builds, round-trip runs) asserts
 * only under the gate.
 *
 * The fixture workbook is saved with the package's real streamed save
 * (`applyCellEditsToXlsx`: one value edit plus one style-only edit that turns
 * bold OFF on a probe cell carrying double underline + fill + border + a
 * custom format — the composition the style pipeline must survive) AND with
 * LibreOffice (`soffice --convert-to xlsx`, its own open + re-save). Both
 * outputs are scored against the ground truth the fixture baked in:
 *
 *   intact   — the attribute is present with its original value and still
 *              attached to the original cell
 *   degraded — present but changed (double underline → single, format code
 *              rewritten, …)
 *   lost     — gone
 *
 * One small styles resolver scores both engines' outputs, so LibreOffice's
 * restructured styles.xml reads the same way the genoffice output does.
 */
import { spawnSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'

const RUN = process.env.XLSX_ROUNDTRIP_ACCURACY === '1'
const OUT = process.env.XLSX_ROUNDTRIP_ACCURACY_OUT ?? ''

// ── fixture (deterministic, JSZip) ──────────────────────────────────────

const XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/**
 * Two sheets. "Invoice" carries the formatting matrix (the probe cell B8
 * carries bold + double underline + fill + double bottom border + a custom
 * format at once — the style-only edit then asks the composition to keep all
 * of them minus bold); "Data" carries the table part and takes the value edit.
 */
async function buildFixture(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XMLDECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `${XMLDECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `${XMLDECL}<workbook xmlns="${MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Invoice" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="TOTALS">Invoice!$B$8</definedName></definedNames></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${XMLDECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
  )
  // fonts: 0 default, 1 bold 14 + color, 2 bold + double underline (the
  // probe), 3 superscript + red, 4 italic. fills: 0 none, 1 FFF2CC, 2
  // D9E2F3. borders: 0 none, 1 thin box, 2 double bottom. numFmts: 164
  // date, 165 currency, 166 percent, 167 custom "Total" accounting.
  zip.file(
    'xl/styles.xml',
    `${XMLDECL}<styleSheet xmlns="${MAIN}"><numFmts count="4"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/><numFmt numFmtId="165" formatCode="$#,##0.00"/><numFmt numFmtId="166" formatCode="0.00%"/><numFmt numFmtId="167" formatCode="_-&quot;Total&quot; * #,##0.00_-"/></numFmts><fonts count="5"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><color rgb="FF1F4E79"/><name val="Calibri"/></font><font><b/><u val="double"/><sz val="11"/><name val="Calibri"/></font><font><vertAlign val="superscript"/><color rgb="FFC00000"/><sz val="11"/><name val="Calibri"/></font><font><i/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="double"><color auto="1"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="8"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="167" fontId="2" fillId="2" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  )
  zip.file(
    'xl/sharedStrings.xml',
    `${XMLDECL}<sst xmlns="${MAIN}" count="3" uniqueCount="3"><si><t>Q3 Invoice</t></si><si><t>Line total</t></si><si><t>Acme Engineering Co., Ltd.</t></si></sst>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${XMLDECL}<worksheet xmlns="${MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:B8"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="24" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" s="1" t="s"><v>0</v></c></row><row r="2"><c r="A2" s="2"><v>46258</v></c><c r="B2" s="3"><v>1234.5</v></c></row><row r="3"><c r="A3" s="4"><v>0.125</v></c></row><row r="4"><c r="A4" s="5"><v>31</v></c><c r="B4"><f>Data!B2*2</f><v>10</v></c></row><row r="5"><c r="A5" t="inlineStr"><is><r><rPr><b/><color rgb="FFC00000"/></rPr><t>Total </t></r><r><t>100</t></r></is></c></row><row r="6"><c r="A6" s="7" t="s"><v>2</v></c></row><row r="7"><c r="A7" t="s"><v>1</v></c></row><row r="8"><c r="B8" s="6"><f>SUM(Data!B2:B3)</f><v>0</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A7:B7"/></mergeCells></worksheet>`,
  )
  zip.file(
    'xl/worksheets/sheet2.xml',
    `${XMLDECL}<worksheet xmlns="${MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:B3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>North</t></is></c><c r="B2"><v>120</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>South</t></is></c><c r="B3"><v>340</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`,
  )
  zip.file(
    'xl/worksheets/_rels/sheet2.xml.rels',
    `${XMLDECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`,
  )
  zip.file(
    'xl/tables/table1.xml',
    `${XMLDECL}<table xmlns="${MAIN}" id="1" name="Regions" displayName="Regions" ref="A1:B3" totalsRowShown="0"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Amount"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`,
  )
  return zip.generateAsync({ type: 'uint8array' })
}

// ── scoring (one resolver for both engines) ─────────────────────────────

interface CellStyle {
  fontXml: string
  fillXml: string
  borderXml: string
  numFmtCode: string | undefined
}

interface Workbook {
  sheetXml(name: string): Promise<string>
  stylesXml(): Promise<string>
  sharedXml(): Promise<string>
  workbookXml(): Promise<string>
  /** undefined when the part does not exist in this output */
  part(path: string): Promise<string | undefined>
}

async function readWorkbook(bytes: Uint8Array): Promise<Workbook> {
  const zip = await JSZip.loadAsync(bytes)
  const read = async (path: string): Promise<string | undefined> =>
    (await zip.file(path)?.async('string')) ?? undefined
  const workbookXml = (await read('xl/workbook.xml')) ?? ''
  const rels = (await read('xl/_rels/workbook.xml.rels')) ?? ''
  const targets = new Map<string, string>()
  // targets are relative to the rels file's part directory (xl/_rels/ → xl/)
  const baseDir = 'xl/'
  for (const m of rels.matchAll(/Id="(rId\d+)"[^>]*Target="([^"]+)"/g)) {
    const raw = m[2]!
    targets.set(m[1]!, raw.startsWith('/') ? raw.slice(1) : baseDir + raw.replace(/^xl\//, ''))
  }
  const sheetPaths = new Map<string, string>()
  for (const m of workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"/g)) {
    const target = targets.get(m[2]!)
    if (target) sheetPaths.set(m[1]!, target)
  }
  return {
    sheetXml: async (name) => {
      const path = sheetPaths.get(name)
      return path ? ((await read(path)) ?? '') : ''
    },
    stylesXml: async () => (await read('xl/styles.xml')) ?? '',
    sharedXml: async () => (await read('xl/sharedStrings.xml')) ?? '',
    workbookXml: async () => workbookXml,
    part: read,
  }
}

interface ResolvedStyles {
  fonts: string[]
  fills: string[]
  borders: string[]
  numFmt(numFmtId: number): string | undefined
  cellStyle(sheetXml: string, ref: string): CellStyle | undefined
}

function parseStyles(stylesXml: string): ResolvedStyles {
  // the plural wrapper (<fonts count="5">) lists singular items (<font>…)
  const list = (tag: string): string[] => {
    const wrap = new RegExp(`<${tag}s\\b[^>]*>([\\s\\S]*?)</${tag}s>`)
    const body = wrap.exec(stylesXml)?.[1]
    if (body === undefined) return []
    return body.match(new RegExp(`<${tag}\\b[\\s\\S]*?</${tag}>|<${tag}\\b[^>]*/>`, 'g')) ?? []
  }
  const fonts = list('font')
  const fills = list('fill')
  const borders = list('border')
  const numFmts = new Map<number, string>()
  for (const m of stylesXml.matchAll(/<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/g)) {
    numFmts.set(Number(m[1]), m[2]!)
  }
  // the xf list lives under <cellXfs>, not <xfs>
  const xfBody = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? ''
  const xfList = xfBody.match(/<xf\b[\s\S]*?<\/xf>|<xf\b[^>]*\/>/g)?.slice(0, 256) ?? []
  const xfs = xfList.map((xf) => ({
    fontId: Number(/fontId="(\d+)"/.exec(xf)?.[1] ?? 0),
    fillId: Number(/fillId="(\d+)"/.exec(xf)?.[1] ?? 0),
    borderId: Number(/borderId="(\d+)"/.exec(xf)?.[1] ?? 0),
    numFmtId: Number(/numFmtId="(\d+)"/.exec(xf)?.[1] ?? 0),
  }))
  const cellBlock = (xml: string, ref: string): string | undefined => {
    // the <c> element whose r= attribute is exactly this ref
    const re = new RegExp(`<c r="${ref}"[ >][\\s\\S]*?</c>|<c r="${ref}"[^>]*/>`)
    return re.exec(xml)?.[0]
  }
  return {
    fonts,
    fills,
    borders,
    numFmt: (id) => numFmts.get(id),
    cellStyle: (sheetXml, ref) => {
      const cell = cellBlock(sheetXml, ref)
      if (cell === undefined) return undefined
      const s = Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)
      const xf = xfs[s]
      if (!xf) return undefined
      const numFmtCode = xfs[s]!.numFmtId >= 164 ? numFmts.get(xf.numFmtId) : undefined
      // a builtin numFmtId (14 = yyyy-mm-date, 44 = accounting) still counts
      return {
        fontXml: fonts[xf.fontId] ?? '',
        fillXml: fills[xf.fillId] ?? '',
        borderXml: borders[xf.borderId] ?? '',
        numFmtCode,
      }
    },
  }
}

type Status = 'intact' | 'degraded' | 'lost'
interface Verdict {
  status: Status
  note?: string
}

const ok = (): Verdict => ({ status: 'intact' })
const lost = (note: string): Verdict => ({ status: 'lost', note })
const degraded = (note: string): Verdict => ({ status: 'degraded', note })

/** every ground-truth attribute, scored on one output workbook */
async function scoreWorkbook(
  wb: Workbook,
  opts: { expectProbeBoldRemoved: boolean; expectEditedValue: boolean },
): Promise<Record<string, Verdict>> {
  const styles = parseStyles(await wb.stylesXml())
  const invoice = await wb.sheetXml('Invoice')
  const data = await wb.sheetXml('Data')
  const workbook = await wb.workbookXml()
  const shared = await wb.sharedXml()
  const v: Record<string, Verdict> = {}

  const cell = (sheetXml: string, ref: string): string | undefined =>
    new RegExp(`<c r="${ref}"[ >][\\s\\S]*?</c>|<c r="${ref}"[^>]*/>`).exec(sheetXml)?.[0]

  const cellText = (cellXml: string): string => {
    const inline = [...cellXml.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((m) => m[1]!).join('')
    if (inline) return inline
    const idx = /t="s"[^>]*>\s*<v>(\d+)<\/v>/.exec(cellXml)?.[1]
    if (idx === undefined) return ''
    const si = new RegExp(`<si[ >][\\s\\S]*?</si>`, 'g')
    const items = [...shared.matchAll(si)]
    return (items[Number(idx)]?.[0] ?? '').replace(/<[^>]+>/g, '')
  }

  // 1. the title cell keeps its text
  v['cell text: title'] = cellText(cell(invoice, 'A1') ?? '').includes('Q3 Invoice')
    ? ok()
    : lost('A1 text')
  // 2. the title keeps bold + size + fill + thin box border
  const title = styles.cellStyle(invoice, 'A1')
  v['title style: bold 14 + fill + thin box'] = title
    ? /<b\/>/.test(title.fontXml) &&
      /sz val="14"/.test(title.fontXml) &&
      /FFFFF2CC/i.test(title.fillXml) &&
      /style="thin"/.test(title.borderXml)
      ? ok()
      : degraded('A1 style parts')
    : lost('A1 style')

  // number formats ride the right cells
  const fmtOf = async (ref: string): Promise<string | undefined> =>
    styles.cellStyle(invoice, ref)?.numFmtCode
  const flat = (code: string | undefined): string => (code ?? '').replace(/\\/g, '').toUpperCase()
  const date = flat(await fmtOf('A2'))
  v['numfmt: date on A2'] =
    date.includes('YYYY') && date.includes('MM') && date.includes('DD')
      ? ok()
      : date
        ? degraded(`formatCode "${date}"`)
        : lost('A2 numfmt')
  const currency = flat(await fmtOf('B2'))
  v['numfmt: currency on B2'] =
    currency.includes('$') && currency.includes('#,##0.00')
      ? ok()
      : currency
        ? degraded(`formatCode "${currency}"`)
        : lost('B2 numfmt')
  const percent = flat(await fmtOf('A3'))
  v['numfmt: percent on A3'] = percent.includes('0.00%')
    ? ok()
    : percent
      ? degraded(`formatCode "${percent}"`)
      : lost('A3 numfmt')

  // formulas
  const b4 = cell(invoice, 'B4') ?? ''
  v['formula: cross-sheet survives'] = /Data!B2\*2|Data!B2\s*\*\s*2/.test(b4)
    ? ok()
    : b4.includes('<f>') && b4.includes('Data!')
      ? degraded(/<f>([\s\S]*?)<\/f>/.exec(b4)?.[1] ?? '')
      : lost('B4 formula')
  const b8 = cell(invoice, 'B8') ?? ''
  v['formula: SUM range survives'] = /SUM\(Data!B2:B3\)/i.test(b8.replace(/\s+/g, ''))
    ? ok()
    : b8.includes('<f>')
      ? degraded(/<f>([\s\S]*?)<\/f>/.exec(b8)?.[1] ?? '')
      : lost('B8 formula')

  // rich text: two runs, the first bold
  const a5 = cell(invoice, 'A5') ?? ''
  const runsXml =
    /<is>([\s\S]*?)<\/is>/.exec(a5)?.[1] ??
    (() => {
      const idx = /t="s"[^>]*>\s*<v>(\d+)<\/v>/.exec(a5)?.[1]
      if (idx === undefined) return ''
      const items = [...shared.matchAll(/<si[ >][\s\S]*?<\/si>/g)]
      return /<r>[\s\S]*?<\/r>/.test(items[Number(idx)]?.[0] ?? '')
        ? (items[Number(idx)]?.[0] ?? '').replace(/^<si[^>]*>/, '').replace(/<\/si>$/, '')
        : ''
    })()
  const runCount = (runsXml.match(/<r>/g) ?? []).length
  const firstRunBold = runsXml.includes('<b/>')
  v['rich text: 2 runs, first bold'] =
    runCount === 2 && firstRunBold
      ? ok()
      : runCount === 1
        ? degraded('runs merged into one')
        : runCount > 0 && !firstRunBold
          ? degraded('runs survive, formatting does not')
          : lost('A5 runs')

  // the probe cell: bold + double underline + fill + double bottom border +
  // custom "Total" format. The genoffice run asks bold OFF; everything else
  // must survive the composition. LibreOffice only round-trips the file, so
  // its probe keeps bold too.
  const probe = styles.cellStyle(invoice, 'B8')
  if (!probe) {
    v['probe: double underline'] = lost('B8 style')
    v['probe: fill'] = lost('B8 style')
    v['probe: double border'] = lost('B8 style')
    v['probe: custom format'] = lost('B8 style')
    v['probe: bold state'] = lost('B8 style')
  } else {
    v['probe: double underline'] = probe.fontXml.includes('<u val="double"/>')
      ? ok()
      : probe.fontXml.includes('<u')
        ? degraded('underline present, variant gone')
        : lost('underline')
    v['probe: fill'] = /FFD9E2F3/i.test(probe.fillXml)
      ? ok()
      : probe.fillXml.includes('patternType="solid"')
        ? degraded('solid fill, color changed')
        : lost('fill')
    v['probe: double border'] = /style="double"/.test(probe.borderXml)
      ? ok()
      : probe.borderXml.includes('style=')
        ? degraded('border present, weight changed')
        : lost('border')
    const custom = flat(probe.numFmtCode)
    v['probe: custom format'] =
      custom.includes('#,##0.00') && custom.includes('TOTAL')
        ? ok()
        : custom
          ? degraded(`formatCode "${custom}"`)
          : lost('custom numfmt')
    const boldOn = probe.fontXml.includes('<b/>')
    v['probe: bold state'] =
      opts.expectProbeBoldRemoved && boldOn
        ? degraded('edit asked bold off, bold still on')
        : !opts.expectProbeBoldRemoved && boldOn
          ? ok()
          : opts.expectProbeBoldRemoved
            ? ok()
            : degraded('bold lost without an edit asking for it')
  }

  // the superscript red font
  const a4 = styles.cellStyle(invoice, 'A4')
  v['font: superscript + color'] =
    a4 && /vertAlign val="superscript"/.test(a4.fontXml) && /FFC00000/i.test(a4.fontXml)
      ? ok()
      : a4 && a4.fontXml.includes('vertAlign')
        ? degraded('vertAlign present, color changed')
        : lost('A4 font')

  // the italic font
  const a6 = styles.cellStyle(invoice, 'A6')
  v['font: italic'] = a6 && /<i\b[^>]*\/>/.test(a6.fontXml) ? ok() : lost('A6 font')

  // structure
  v['merged range A7:B7'] = /mergeCell ref="A7:B7"/.test(invoice)
    ? ok()
    : invoice.includes('mergeCell')
      ? degraded(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/.exec(invoice)?.[1] ?? '')
      : lost('merge')
  v['defined name TOTALS'] = /name="TOTALS"/.test(workbook) ? ok() : lost('defined name')
  const table = await wb.part('xl/tables/table1.xml')
  v['table part: name + columns'] =
    table &&
    /name="Regions"/.test(table) &&
    /name="Region"/.test(table) &&
    /name="Amount"/.test(table)
      ? ok()
      : table
        ? degraded('table part present, definition changed')
        : lost('table part')
  const sheetOrder = [...workbook.matchAll(/<sheet[^>]*name="([^"]+)"/g)].map((m) => m[1]!)
  v['sheet order: Invoice, Data'] =
    sheetOrder.join(',') === 'Invoice,Data' ? ok() : degraded(sheetOrder.join(','))
  v['frozen top row'] =
    /ySplit="1"/.test(invoice) && /state="frozen"/.test(invoice) ? ok() : lost('freeze pane')
  v['column width survives'] = /width="24"/.test(invoice) ? ok() : lost('col width')
  v['table header text'] = cellText(cell(data, 'A1') ?? '').includes('Region')
    ? ok()
    : lost('Data!A1')

  // the value edit landed (genoffice only)
  if (opts.expectEditedValue) {
    const b2 = cell(data, 'B2') ?? ''
    v['edit landed: Data!B2 = 5'] = /<v>5<\/v>/.test(b2) ? ok() : lost(`B2 = ${cellText(b2)}`)
  }

  return v
}

// ── the bench ───────────────────────────────────────────────────────────

describe.skipIf(!RUN)('xlsx round-trip accuracy bench (styled workbook class)', () => {
  it('scores genoffice streamed save vs soffice round-trip', async () => {
    const fixture = await buildFixture()
    const workDir = mkdtempSync(join(tmpdir(), 'xlsx-roundtrip-accuracy-'))
    const fixturePath = join(workDir, 'styled.xlsx')
    writeFileSync(fixturePath, fixture)

    // the app's save: one value edit on Data!B2 and one style-only edit that
    // turns bold OFF on the probe cell B8 (the #2912-class composition)
    const saved = await applyCellEditsToXlsx(Buffer.from(fixture), [
      { sheetName: 'Data', row: 1, column: 1, writeValue: true, cell: { value: 5 } },
      {
        sheetName: 'Invoice',
        row: 7,
        column: 1,
        writeValue: false,
        cell: { value: null },
        style: { bold: false },
      },
    ])

    const genoffice = await readWorkbook(saved.buffer)
    const genofficeVerdicts = await scoreWorkbook(genoffice, {
      expectProbeBoldRemoved: true,
      expectEditedValue: true,
    })

    const loOut = join(workDir, 'lo-out')
    const run = spawnSync(
      'soffice',
      [
        '--headless',
        '--convert-to',
        'xlsx',
        '--outdir',
        loOut,
        `-env:UserInstallation=file:///${join(workDir, 'lo-profile')}`,
        fixturePath,
      ],
      { timeout: 120_000 },
    )
    if (run.status !== 0) throw new Error('soffice failed')
    const soffice = await readWorkbook(new Uint8Array(readFileSync(join(loOut, 'styled.xlsx'))))
    const sofficeVerdicts = await scoreWorkbook(soffice, {
      expectProbeBoldRemoved: false,
      expectEditedValue: false,
    })

    const attributes = Object.keys(genofficeVerdicts).filter(
      (k) => k !== 'edit landed: Data!B2 = 5',
    )
    const scorecard: string[] = [
      '# XLSX round-trip accuracy — styled workbook class',
      '',
      'Attribute = a formatting/structure fact the fixture baked in. intact = present',
      'with its original value on the original cell. The genoffice run saves with two',
      'edits (Data!B2 = 5, and bold OFF on the probe cell that also carries double',
      'underline + fill + double border + a custom format); the LibreOffice run is its',
      'own open + re-save of the untouched fixture.',
      '',
      '| attribute | genoffice | soffice |',
      '|---|---|---|',
    ]
    let genofficeIntact = 0
    let sofficeIntact = 0
    for (const key of attributes) {
      const g = genofficeVerdicts[key]!
      const s = sofficeVerdicts[key]!
      if (g.status === 'intact') genofficeIntact += 1
      if (s.status === 'intact') sofficeIntact += 1
      const note = (x: Verdict) =>
        x.status === 'intact' ? '✓' : `${x.status}${x.note ? `: ${x.note}` : ''}`
      scorecard.push(`| ${key} | ${note(g)} | ${note(s)} |`)
    }
    scorecard.push(
      '',
      `genoffice intact ${genofficeIntact}/${attributes.length}; LibreOffice intact ${sofficeIntact}/${attributes.length}.`,
    )
    rmSync(workDir, { recursive: true, force: true })

    const md = `${scorecard.join('\n')}\n`
    if (OUT) writeFileSync(OUT, md)
    else console.log(`\n${md}`)

    // smoke-level pins only: the corpus ran, and the composition guarantees the
    // style pipeline exists to keep are real (the numbers above are narrative)
    expect(genofficeVerdicts['probe: double underline']!.status).toBe('intact')
    expect(genofficeVerdicts['probe: fill']!.status).toBe('intact')
    expect(genofficeVerdicts['edit landed: Data!B2 = 5']!.status).toBe('intact')
  })
})
