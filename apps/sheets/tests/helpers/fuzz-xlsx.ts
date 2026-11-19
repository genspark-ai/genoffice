/** Seeded generator of varied-but-valid XLSX packages for the round-trip
 *  fidelity suite — the sheets sibling of docs' helpers/fuzz-docx.ts. Same
 *  mulberry32 seeding as the hostile-input suite (same seed → same bytes, a
 *  failing case reproduces exactly via buildFuzzXlsx(seed)); the knobs cover
 *  the save pipeline's part families: shared and inline strings, numbers,
 *  booleans, errors, formulas with cached values, a stylesheet with several
 *  cell formats, merges, frozen panes, an autofilter, conditional formats,
 *  data validation, hyperlinks, threaded-style legacy comments (comments part
 *  + VML drawing + sheet rels), tables (table part + sheet rels +
 *  tableParts), a defined name and up to three sheets with CJK and latin
 *  names. Parts a seed omits stay omitted, so saves also run against
 *  packages that never had a stylesheet or a shared-strings part.
 *
 *  Every part family must fire at least once across the CI seed range; the
 *  suite pins that with a no-dead-knobs test. */
import JSZip from 'jszip'

export function rng(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a |= 0
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const CJK_NAMES = [
  '\u5b63\u5ea6\u56de\u987e',
  '\u8425\u6536\u660e\u7ec6',
  '\u9644\u5f55 \u7532',
  '\u98ce\u9669\u767b\u8bb0',
]
const LATIN_NAMES = ['Summary', 'data 2024', 'Q3', 'notes']
const TEXTS = [
  '\u5b63\u5ea6\u56de\u987e',
  'Quarterly review',
  '\u98ce\u9669\u767b\u8bb0',
  'milestone 12',
  'draft v2',
  '\u9644\u5f55 \u7532',
]
const FORMULAS = ['SUM(A1:A3)', 'B1*2', 'A2&amp;" x"', 'IF(A1>0,1,0)'] as const

interface SheetPlan {
  readonly name: string
  readonly sheetId: number
  readonly rows: number
  readonly cols: number
  merge?: string
  pane?: boolean
  autoFilter?: string
  conditionalFormatting?: boolean
  dataValidation?: boolean
  hyperlink?: boolean
  comment?: string
  table?: { readonly ref: string; readonly colNames: readonly string[] }
}

/** Build one valid workbook: the whole seed's knob draws happen here so the
 *  suite only sees bytes. */
export async function buildFuzzXlsx(seed: number): Promise<Uint8Array> {
  const r = rng(seed * 7919 + 17)
  const pick = <T>(list: readonly T[]): T => list[Math.floor(r() * list.length)]!
  const zip = new JSZip()

  const sheetCount = 1 + Math.floor(r() * 3)
  const plans: SheetPlan[] = Array.from({ length: sheetCount }, (_v, i) => {
    const rows = 1 + Math.floor(r() * 12)
    const cols = 1 + Math.floor(r() * 6)
    const firstCell = 'A1'
    const lastCell = `${String.fromCharCode(65 + cols - 1)}${rows}`
    const plan: SheetPlan = {
      name:
        r() < 0.5
          ? CJK_NAMES[(seed + i) % CJK_NAMES.length]!
          : LATIN_NAMES[(seed * 3 + i) % LATIN_NAMES.length]!,
      sheetId: i + 1,
      rows,
      cols,
    }
    if (r() < 0.4) {
      const mergeCols = Math.max(2, Math.min(cols, 2 + Math.floor(r() * 2)))
      plan.merge = `A1:${String.fromCharCode(64 + mergeCols)}${Math.min(rows, 2)}`
    }
    if (r() < 0.4) plan.pane = true
    if (r() < 0.3) plan.autoFilter = `A1:${lastCell}`
    if (r() < 0.4) plan.conditionalFormatting = true
    if (r() < 0.3) plan.dataValidation = true
    if (r() < 0.3) plan.hyperlink = true
    if (r() < 0.3) plan.comment = pick(TEXTS)
    // tables own their range: an autofilter over the same range would be
    // ambiguous, so a table suppresses the plain autofilter knob
    if (r() < 0.3 && !plan.autoFilter) {
      const tableCols = Math.max(1, Math.min(cols, 3))
      plan.table = {
        ref: `A1:${String.fromCharCode(64 + tableCols)}${Math.max(rows, 2)}`,
        colNames: Array.from({ length: tableCols }, (_x, c) => `Col${c + 1}`),
      }
    }
    void firstCell
    return plan
  })

  const withStyles = r() < 0.85
  const styleCount = withStyles ? 3 : 1
  const withShared = r() < 0.8
  const shared: string[] = []
  if (withShared) {
    const count = 3 + Math.floor(r() * 5)
    for (let i = 0; i < count; i++) shared.push(pick(TEXTS))
  }
  const withDefinedName = r() < 0.7

  // ── [Content_Types].xml ────────────────────────────────────────────
  const overrides: string[] = [
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
  ]
  for (const plan of plans) {
    overrides.push(
      `<Override PartName="/xl/worksheets/sheet${plan.sheetId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    if (plan.comment) {
      overrides.push(
        `<Override PartName="/xl/comments${plan.sheetId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`,
      )
    }
    if (plan.table) {
      overrides.push(
        `<Override PartName="/xl/tables/table${plan.sheetId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`,
      )
    }
  }
  if (plans.some((p) => p.comment)) {
    zip.file(
      '[Content_Types].xml',
      `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        overrides.join('') +
        '</Types>',
    )
  } else {
    zip.file(
      '[Content_Types].xml',
      `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        overrides.join('') +
        '</Types>',
    )
  }

  // ── package + workbook rels ────────────────────────────────────────
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="${RELS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="xl/workbook.xml"/>` +
      '</Relationships>',
  )

  const definedName = withDefinedName
    ? `<definedNames><definedName name="Scope_${seed}">'${plans[0]!.name.replace(/&/g, '&amp;')}'!$A$1:$B$2</definedName></definedNames>`
    : ''
  zip.file(
    'xl/workbook.xml',
    `${XML_DECL}<workbook xmlns="${MAIN}" xmlns:r="${OFFICE_REL}">` +
      definedName +
      '<sheets>' +
      plans
        .map(
          (plan) =>
            `<sheet name="${plan.name.replace(/&/g, '&amp;')}" sheetId="${plan.sheetId}" r:id="rId${plan.sheetId}"/>`,
        )
        .join('') +
      '</sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${XML_DECL}<Relationships xmlns="${RELS}">` +
      plans
        .map(
          (plan) =>
            `<Relationship Id="rId${plan.sheetId}" Type="${OFFICE_REL}/worksheet" Target="worksheets/sheet${plan.sheetId}.xml"/>`,
        )
        .join('') +
      '</Relationships>',
  )

  // ── stylesheet: three usable cell formats + the default ───────────
  if (withStyles) {
    zip.file(
      'xl/styles.xml',
      `${XML_DECL}<styleSheet xmlns="${MAIN}">` +
        '<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy&quot;-&quot;mm&quot;-&quot;dd"/><numFmt numFmtId="165" formatCode="0.0%"/></numFmts>' +
        '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><name val="Arial"/></font><font><i/><sz val="10"/><name val="Calibri"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill></fills>' +
        '<borders count="2"><border/><border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right></border></borders>' +
        '<cellStyleXfs count="1"><xf/></cellStyleXfs>' +
        '<cellXfs count="4"><xf xfId="0"/><xf xfId="0" fontId="1" applyFont="1"/><xf xfId="0" fontId="2" fillId="2" numFmtId="164" applyNumberFormat="1" applyFill="1" applyFont="1"/><xf xfId="0" numFmtId="165" applyNumberFormat="1"/></cellXfs>' +
        '</styleSheet>',
    )
  }

  if (withShared) {
    zip.file(
      'xl/sharedStrings.xml',
      `${XML_DECL}<sst xmlns="${MAIN}" count="${shared.length}" uniqueCount="${shared.length}">` +
        shared.map((text) => `<si><t>${text}</t></si>`).join('') +
        '</sst>',
    )
  }

  // ── sheets ─────────────────────────────────────────────────────────
  for (const plan of plans) {
    const cells: string[] = []
    let formulaCell: string | null = null
    for (let row = 1; row <= plan.rows; row++) {
      const rowCells: string[] = []
      for (let col = 0; col < plan.cols; col++) {
        const ref = `${String.fromCharCode(65 + col)}${row}`
        const style =
          withStyles && r() < 0.25 ? ` s="${1 + Math.floor(r() * (styleCount - 1))}"` : ''
        const kind = r()
        if (kind < 0.25 && shared.length > 0) {
          rowCells.push(`<c r="${ref}"${style} t="s"><v>${Math.floor(r() * shared.length)}</v></c>`)
        } else if (kind < 0.4) {
          rowCells.push(`<c r="${ref}"${style} t="inlineStr"><is><t>${pick(TEXTS)}</t></is></c>`)
        } else if (kind < 0.65) {
          rowCells.push(`<c r="${ref}"${style}><v>${Math.floor(r() * 100000) / 100}</v></c>`)
        } else if (kind < 0.75 && !formulaCell) {
          rowCells.push(
            `<c r="${ref}"${style}><f>${pick(FORMULAS)}</f><v>${Math.floor(r() * 500)}</v></c>`,
          )
          formulaCell = ref
        } else if (kind < 0.85) {
          rowCells.push(`<c r="${ref}"${style} t="b"><v>${r() < 0.5 ? 1 : 0}</v></c>`)
        } else if (kind < 0.92) {
          rowCells.push(`<c r="${ref}"${style} t="e"><v>#N/A</v></c>`)
        } else {
          rowCells.push(`<c r="${ref}"${style}/>`)
        }
      }
      const rowHt =
        withStyles && r() < 0.3 ? ` ht="${(14 + r() * 20).toFixed(2)}" customHeight="1"` : ''
      cells.push(`<row r="${row}"${rowHt}>${rowCells.join('')}</row>`)
    }

    const lastCol = String.fromCharCode(64 + plan.cols)
    const lastRow = plan.rows
    const sheetViews = plan.pane
      ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
      : '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    const colsXml =
      withStyles && r() < 0.5
        ? `<cols><col min="1" max="${plan.cols}" width="${(8 + r() * 10).toFixed(2)}" customWidth="1"/></cols>`
        : ''
    const autoFilterXml = plan.autoFilter ? `<autoFilter ref="${plan.autoFilter}"/>` : ''
    const mergeXml = plan.merge
      ? `<mergeCells count="1"><mergeCell ref="${plan.merge}"/></mergeCells>`
      : ''
    const cfXml = plan.conditionalFormatting
      ? `<conditionalFormatting sqref="A1:${lastCol}${lastRow}">` +
        (r() < 0.5
          ? '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule>'
          : '<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>') +
        '</conditionalFormatting>'
      : ''
    const dvXml = plan.dataValidation
      ? `<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="1" sqref="A1:A5"><formula1>0</formula1><formula2>100</formula2></dataValidation></dataValidations>`
      : ''
    // sheet rels collect the optional parts; ids stay stable so the sheet XML
    // below can reference them even when a family is skipped
    const sheetRels: string[] = []
    let relId = 1
    const hyperlinkRelId = plan.hyperlink ? `rId${relId++}` : null
    const commentRelId = plan.comment ? `rId${relId++}` : null
    const vmlRelId = plan.comment ? `rId${relId++}` : null
    const tableRelId = plan.table ? `rId${relId}` : null
    if (hyperlinkRelId) {
      sheetRels.push(
        `<Relationship Id="${hyperlinkRelId}" Type="${OFFICE_REL}/hyperlink" Target="https://example.com/genoffice" TargetMode="External"/>`,
      )
    }
    if (commentRelId) {
      sheetRels.push(
        `<Relationship Id="${commentRelId}" Type="${OFFICE_REL}/comments" Target="../comments${plan.sheetId}.xml"/>`,
      )
    }
    if (vmlRelId) {
      sheetRels.push(
        `<Relationship Id="${vmlRelId}" Type="${OFFICE_REL}/vmlDrawing" Target="../drawings/vmlDrawing${plan.sheetId}.vml"/>`,
      )
    }
    if (tableRelId) {
      sheetRels.push(
        `<Relationship Id="${tableRelId}" Type="${OFFICE_REL}/table" Target="../tables/table${plan.sheetId}.xml"/>`,
      )
    }
    const hyperlinksXml =
      plan.hyperlink && hyperlinkRelId
        ? `<hyperlinks><hyperlink ref="A1" r:id="${hyperlinkRelId}"/></hyperlinks>`
        : ''
    const legacyDrawingXml = plan.comment && vmlRelId ? `<legacyDrawing r:id="${vmlRelId}"/>` : ''
    const tablePartsXml =
      plan.table && tableRelId
        ? `<tableParts count="1"><tablePart r:id="${tableRelId}"/></tableParts>`
        : ''

    zip.file(
      `xl/worksheets/sheet${plan.sheetId}.xml`,
      `${XML_DECL}<worksheet xmlns="${MAIN}" xmlns:r="${OFFICE_REL}">` +
        sheetViews +
        colsXml +
        `<sheetData>${cells.join('')}</sheetData>` +
        autoFilterXml +
        mergeXml +
        cfXml +
        dvXml +
        hyperlinksXml +
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
        legacyDrawingXml +
        tablePartsXml +
        '</worksheet>',
    )
    if (sheetRels.length) {
      zip.file(
        `xl/worksheets/_rels/sheet${plan.sheetId}.xml.rels`,
        `${XML_DECL}<Relationships xmlns="${RELS}">${sheetRels.join('')}</Relationships>`,
      )
    }

    if (plan.comment && commentRelId && vmlRelId) {
      zip.file(
        `xl/comments${plan.sheetId}.xml`,
        `${XML_DECL}<comments xmlns="${MAIN}">` +
          '<authors><author>fuzz</author></authors>' +
          `<commentList><comment ref="A1" authorId="0"><r><t>${plan.comment}</t></r></comment></commentList>` +
          '</comments>',
      )
      zip.file(
        `xl/drawings/vmlDrawing${plan.sheetId}.vml`,
        '<xml xmlns:v="urn:schemas-microsoft-com:vml"' +
          ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
          ' xmlns:x="urn:schemas-microsoft-com:office:excel">' +
          '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
          '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
          '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
          `<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
          '<v:fill color="ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
          '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
          '</v:shape></xml>',
      )
    }

    if (plan.table && tableRelId) {
      zip.file(
        `xl/tables/table${plan.sheetId}.xml`,
        `${XML_DECL}<table xmlns="${MAIN}" id="${plan.sheetId}" name="Table${plan.sheetId}" displayName="Table${plan.sheetId}" ref="${plan.table.ref}">` +
          `<autoFilter ref="${plan.table.ref}"/>` +
          `<tableColumns count="${plan.table.colNames.length}">` +
          plan.table.colNames
            .map((name, i) => `<tableColumn id="${i + 1}" name="${name}"/>`)
            .join('') +
          '</tableColumns>' +
          '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>' +
          '</table>',
      )
    }
  }

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
