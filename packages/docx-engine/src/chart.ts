import JSZip from 'jszip'
import type { ChartDisplay, ChartSeries, NewChart } from './types'
import {
  attrsOf,
  childrenOf,
  escapeXmlText,
  findChild,
  findChildren,
  nameOf,
  textOf,
  xmlParser,
  type XNode,
} from './xml-utils'

const CHART_KINDS: Record<string, ChartDisplay['kind']> = {
  'c:barChart': 'bar',
  'c:bar3DChart': 'bar',
  'c:lineChart': 'line',
  'c:line3DChart': 'line',
  'c:pieChart': 'pie',
  'c:pie3DChart': 'pie',
  'c:doughnutChart': 'pie',
  'c:areaChart': 'area',
  'c:area3DChart': 'area',
}

/**
 * Read the display model of a chart part (word/charts/chartN.xml). Only the
 * caches Word writes next to the data references (c:strCache / c:numCache)
 * are read — the embedded workbook is never opened. Returns null when the
 * part has no series with cached values (e.g. scatter charts).
 */
export function parseChartPartXml(xml: string, partPath: string): ChartDisplay | null {
  const parsed = xmlParser.parse(xml) as XNode[]
  if (xml.includes('<cx:chartSpace')) return parseChartexPartXml(parsed, partPath)
  const space = parsed.find((n) => nameOf(n) === 'c:chartSpace')
  const chart = space ? findChild(space, 'c:chart') : undefined
  const plotArea = chart ? findChild(chart, 'c:plotArea') : undefined
  if (!chart || !plotArea) return null

  // first plotted chart type wins (combo charts render their primary series)
  const plot = childrenOf(plotArea).find((c) => (nameOf(c) ?? '').endsWith('Chart'))
  if (!plot) return null
  const kind = CHART_KINDS[nameOf(plot) ?? ''] ?? 'other'
  const horizontal = kind === 'bar' && attrsOf(findChild(plot, 'c:barDir') ?? {})['val'] === 'bar'

  let categories: string[] = []
  const series: ChartSeries[] = []
  for (const ser of findChildren(plot, 'c:ser')) {
    const val = findChild(ser, 'c:val')
    const rawValues = val ? cachePoints(val) : []
    const values = rawValues.map((v) => {
      if (v === null || v.trim() === '') return null
      const n = Number(v)
      return Number.isFinite(n) ? n : null
    })
    if (values.length === 0) continue
    const cat = findChild(ser, 'c:cat')
    if (cat && categories.length === 0) {
      categories = cachePoints(cat).map((v) => v ?? '')
      // date-formatted numeric caches hold Excel serials; display them as dates
      const fmt = catFormatCode(cat)
      if (fmt && /[yd]/i.test(fmt)) {
        categories = categories.map((v) => serialDateText(v) ?? v)
      }
    }
    const name = seriesName(ser)
    series.push({ ...(name !== undefined ? { name } : {}), values })
  }
  if (series.length === 0) return null

  const title = chartTitle(chart)
  return {
    partPath,
    kind,
    ...(horizontal ? { horizontal } : {}),
    ...(title !== undefined ? { title } : {}),
    categories,
    series,
  }
}

/** cx layoutId → closest classic display kind (degrade: shapes differ, data/labels/title survive) */
const CHARTEX_KINDS: Record<string, ChartDisplay['kind']> = {
  clusteredColumn: 'bar',
  boxWhisker: 'bar',
  waterfall: 'bar',
  funnel: 'bar',
  paretoLine: 'line',
  sunburst: 'pie',
  treemap: 'pie',
}

/**
 * Chartex (cx: 2014 chart extension: sunburst/treemap/boxWhisker/waterfall…)
 * degrade: read cx:chartData dimensions + series names + title into the
 * classic ChartDisplay model so the existing chart pipeline renders the data
 * with the nearest classic shape (title and cached texts/numbers survive).
 */
function parseChartexPartXml(parsed: XNode[], partPath: string): ChartDisplay | null {
  const space = parsed.find((n) => nameOf(n) === 'cx:chartSpace')
  if (!space) return null
  const chartData = findChild(space, 'cx:chartData')
  // data id → {cats, vals}: strDim/numDim carry cx:lvl point caches (the first
  // strDim level holds the leaf labels of hierarchical charts)
  const dataById = new Map<string, { cats: string[]; vals: (number | null)[] }>()
  for (const data of chartData ? findChildren(chartData, 'cx:data') : []) {
    const id = attrsOf(data)['id'] ?? ''
    const entry: { cats: string[]; vals: (number | null)[] } = { cats: [], vals: [] }
    const ptsOf = (dim: XNode): (string | null)[] => {
      const lvl = findChild(dim, 'cx:lvl')
      if (!lvl) return []
      const out: (string | null)[] = []
      for (const pt of findChildren(lvl, 'cx:pt')) {
        const idx = parseInt(attrsOf(pt)['idx'] ?? '', 10)
        if (Number.isFinite(idx) && idx >= 0) out[idx] = textOf(pt)
      }
      return out
    }
    for (const dim of childrenOf(data)) {
      const dimName = nameOf(dim)
      if (dimName === 'cx:strDim') {
        entry.cats = ptsOf(dim).map((v) => v ?? '')
      } else if (dimName === 'cx:numDim') {
        entry.vals = ptsOf(dim).map((v) => {
          if (v === null || v.trim() === '') return null
          const n = Number(v)
          return Number.isFinite(n) ? n : null
        })
      }
    }
    dataById.set(id, entry)
  }
  const chart = findChild(space, 'cx:chart')
  const plotRegion = chart
    ? findChild(findChild(chart, 'cx:plotArea') ?? {}, 'cx:plotAreaRegion')
    : undefined
  let kind: ChartDisplay['kind'] = 'other'
  let categories: string[] = []
  const series: ChartSeries[] = []
  for (const ser of plotRegion ? findChildren(plotRegion, 'cx:series') : []) {
    const layout = attrsOf(ser)['layoutId'] ?? ''
    if (kind === 'other' && CHARTEX_KINDS[layout]) kind = CHARTEX_KINDS[layout]
    const dataId = attrsOf(findChild(ser, 'cx:dataId') ?? {})['val'] ?? ''
    const data = dataById.get(dataId)
    if (!data || data.vals.length === 0) continue
    if (categories.length === 0) categories = data.cats
    const name = textOf(
      findChild(findChild(findChild(ser, 'cx:tx') ?? {}, 'cx:txData') ?? {}, 'cx:v') ?? {},
    )
    series.push({ ...(name ? { name } : {}), values: data.vals })
  }
  if (series.length === 0) return null
  // cx:title holds a full a:t rich body like classic charts
  const title = chart ? chartexTitle(chart) : undefined
  return {
    partPath,
    kind,
    ...(title !== undefined ? { title } : {}),
    categories,
    series,
  }
}

function chartexTitle(chart: XNode): string | undefined {
  const title = findChild(chart, 'cx:title')
  if (!title) return undefined
  const texts: string[] = []
  const walk = (node: XNode) => {
    for (const child of childrenOf(node)) {
      if (nameOf(child) === 'a:t') texts.push(textOf(child))
      else walk(child)
    }
  }
  walk(title)
  const joined = texts.join('')
  return joined !== '' ? joined : undefined
}

/** c:formatCode of a category cache (numCache/numLit), if any */
function catFormatCode(container: XNode): string | undefined {
  const ref = findChild(container, 'c:numRef')
  const cache = ref ? findChild(ref, 'c:numCache') : findChild(container, 'c:numLit')
  const code = cache ? findChild(cache, 'c:formatCode') : undefined
  return code ? textOf(code) : undefined
}

/** Excel date serial → "m/d/yyyy" display (Word/LO render category dates, not serials) */
function serialDateText(v: string | null): string | null {
  const n = Number(v)
  if (!Number.isFinite(n) || n <= 0 || n > 80000) return null
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000)
  return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`
}

/** cached point texts of a c:cat / c:val / c:tx container, in idx order */
function cachePoints(container: XNode): (string | null)[] {
  const ref = findChild(container, 'c:strRef') ?? findChild(container, 'c:numRef')
  const cache = ref
    ? (findChild(ref, 'c:strCache') ?? findChild(ref, 'c:numCache'))
    : (findChild(container, 'c:strLit') ?? findChild(container, 'c:numLit'))
  if (!cache) return []
  const count = parseInt(attrsOf(findChild(cache, 'c:ptCount') ?? {})['val'] ?? '', 10)
  const points: (string | null)[] = []
  for (const pt of findChildren(cache, 'c:pt')) {
    const idx = parseInt(attrsOf(pt)['idx'] ?? '', 10)
    if (!Number.isFinite(idx) || idx < 0) continue
    points[idx] = textOf(findChild(pt, 'c:v') ?? {})
  }
  const length = Number.isFinite(count) ? Math.max(count, points.length) : points.length
  const out: (string | null)[] = []
  for (let i = 0; i < length; i++) out.push(points[i] ?? null)
  return out
}

function seriesName(ser: XNode): string | undefined {
  const tx = findChild(ser, 'c:tx')
  if (!tx) return undefined
  const literal = findChild(tx, 'c:v')
  if (literal) return textOf(literal)
  const cached = cachePoints(tx)
  return cached[0] ?? undefined
}

function chartTitle(chart: XNode): string | undefined {
  const title = findChild(chart, 'c:title')
  if (!title) return undefined
  const texts: string[] = []
  const walk = (node: XNode) => {
    for (const child of childrenOf(node)) {
      if (nameOf(child) === 'a:t') texts.push(textOf(child))
      else walk(child)
    }
  }
  walk(title)
  let joined = texts.join('')
  if (joined) return joined
  // strRef titles carry the cached text in c:strCache c:v, not a:t
  const walkV = (node: XNode) => {
    for (const child of childrenOf(node)) {
      if (nameOf(child) === 'c:v') texts.push(textOf(child))
      else walkV(child)
    }
  }
  walkV(title)
  joined = texts.join('')
  if (joined) return joined
  // text-less c:title = auto title; Word renders the "Chart Title" placeholder
  // unless the auto title was explicitly deleted (CT_Boolean: a val-less
  // element and val="true" both mean true)
  const del = findChild(chart, 'c:autoTitleDeleted')
  const delVal = del ? attrsOf(del)['val'] : undefined
  const deleted = del !== undefined && (delVal === undefined || delVal === '1' || delVal === 'true')
  return deleted ? undefined : 'Chart Title'
}

const _XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
export const CHART_WORKBOOK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'

const colLetter = (i: number) => String.fromCharCode(66 + i) // B, C, D...

function strCacheXml(values: string[], f: string): string {
  return (
    `<c:strRef><c:f>${escapeXmlText(f)}</c:f><c:strCache><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXmlText(v)}</c:v></c:pt>`).join('') +
    '</c:strCache></c:strRef>'
  )
}

function numCacheXml(values: (number | null)[], f: string): string {
  return (
    `<c:numRef><c:f>${escapeXmlText(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode>` +
    `<c:ptCount val="${values.length}"/>` +
    values.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('') +
    '</c:numCache></c:numRef>'
  )
}

/**
 * Build a complete chart part (word/charts/chartN.xml) from data. The chart
 * references an embedded workbook via c:externalData so Word's "Edit Data" works.
 * Pass `externalDataRId` to wire the c:externalData relationship; omit it when
 * the workbook will be added in a separate step.
 */
export function buildChartPartXml(chart: NewChart, externalDataRId?: string): string {
  const rows = chart.categories.length
  const sers = chart.series
    .map((ser, i) => {
      const col = colLetter(i)
      return (
        `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
        `<c:tx>${strCacheXml([ser.name], `Sheet1!$${col}$1`)}</c:tx>` +
        `<c:cat>${strCacheXml(chart.categories, `Sheet1!$A$2:$A$${rows + 1}`)}</c:cat>` +
        `<c:val>${numCacheXml(ser.values.slice(0, rows), `Sheet1!$${col}$2:$${col}$${rows + 1}`)}</c:val>` +
        '</c:ser>'
      )
    })
    .join('')

  let plot: string
  if (chart.kind === 'pie') {
    plot = `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`
  } else {
    const axes =
      '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>' +
      '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>'
    const inner =
      chart.kind === 'bar'
        ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}` +
          '<c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>'
        : `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>` +
          '<c:axId val="111111111"/><c:axId val="222222222"/></c:lineChart>'
    plot = inner + axes
  }

  const title = chart.title
    ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
      `<a:t>${escapeXmlText(chart.title)}</a:t></a:r></a:p></c:rich></c:tx>` +
      '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
    : ''

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>` +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    (externalDataRId
      ? `<c:externalData r:id="${externalDataRId}"><c:autoUpdate val="0"/></c:externalData>`
      : '') +
    '</c:chartSpace>'
  )
}

export interface ChartSeriesPatch {
  name?: string
  /** aligned with ChartSeries.values; null = keep the original value */
  values?: (number | null)[]
}

export interface ChartPatch {
  title?: string
  /** aligned with ChartDisplay.categories; null = keep */
  categories?: (string | null)[]
  /** aligned with ChartDisplay.series; null = keep that series untouched */
  series?: (ChartSeriesPatch | null)[]
}

/**
 * Patch cached texts/numbers of a chart part while keeping the structure —
 * data references (c:f), styling, layout — byte-identical. Anything the
 * patch cannot anchor (missing title, missing cache point) is left as-is.
 * The embedded workbook is intentionally not touched: Word renders from
 * these caches, but "Edit Data" will show the original sheet numbers.
 */
export function patchChartPartXml(xml: string, patch: ChartPatch): string {
  // a text-less title (auto title / strRef with no rich body) has no a:t to
  // rewrite: give it one first, then patch it like any other title
  if (patch.title !== undefined) {
    const t = tagRange(xml, 'c:title')
    if (
      t &&
      innerTextRanges(xml, 'a:t', t.start, t.end).length === 0 &&
      innerTextRanges(xml, 'c:v', t.start, t.end).length === 0
    ) {
      const runXml = `<a:r><a:t>${escapeXmlText(patch.title)}</a:t></a:r>`
      const tx = tagRange(xml, 'c:tx', t.start, t.end)
      const p = tx ? tagRange(xml, 'a:p', tx.start, tx.end) : null
      if (p) {
        // Word auto titles carry an empty c:tx/c:rich paragraph (only
        // a:endParaRPr); CT_Title allows one c:tx, so inject the run there,
        // before a:endParaRPr per schema order
        const endPr = xml.indexOf('<a:endParaRPr', p.start)
        const at = endPr !== -1 && endPr < p.end ? endPr : p.end - '</a:p>'.length
        xml = xml.slice(0, at) + runXml + xml.slice(at)
      } else {
        const rich = `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>${runXml}</a:p></c:rich></c:tx>`
        if (tx) {
          // cache-less strRef tx (no a:p, no c:v): nothing to inject into —
          // replace the whole c:tx with a rich body (CT_Tx is a choice)
          xml = xml.slice(0, tx.start) + rich + xml.slice(tx.end)
        } else {
          const openEnd = xml.indexOf('>', t.start) + 1
          xml = xml.slice(0, openEnd) + rich + xml.slice(openEnd)
        }
      }
    }
  }

  const edits: Array<{ start: number; end: number; text: string }> = []

  if (patch.title !== undefined) {
    const title = tagRange(xml, 'c:title')
    if (title) {
      let texts = innerTextRanges(xml, 'a:t', title.start, title.end)
      // strRef titles keep the text in the c:strCache c:v cache instead
      if (texts.length === 0) texts = innerTextRanges(xml, 'c:v', title.start, title.end)
      // whole title into the first run, remaining runs blanked
      texts.forEach((range, i) => {
        edits.push({ ...range, text: i === 0 ? patch.title! : '' })
      })
    }
  }

  const serRanges = tagRanges(xml, 'c:ser')
  serRanges.forEach((ser, i) => {
    const serPatch = patch.series?.[i]
    if (serPatch?.name !== undefined) {
      const tx = tagRange(xml, 'c:tx', ser.start, ser.end)
      if (tx) {
        const texts = innerTextRanges(xml, 'c:v', tx.start, tx.end)
        if (texts.length > 0) edits.push({ ...texts[0], text: serPatch.name })
      }
    }
    if (serPatch?.values) {
      const val = tagRange(xml, 'c:val', ser.start, ser.end)
      if (val)
        pushPointEdits(
          xml,
          val,
          serPatch.values.map((v) => (v === null ? null : String(v))),
          edits,
        )
    }
    // categories are cached per series; every copy must agree
    if (patch.categories) {
      const cat = tagRange(xml, 'c:cat', ser.start, ser.end)
      if (cat) pushPointEdits(xml, cat, patch.categories, edits)
    }
  })

  if (edits.length === 0) return xml
  let out = xml
  for (const edit of edits.sort((a, b) => b.start - a.start)) {
    out = out.slice(0, edit.start) + escapeXmlText(edit.text) + out.slice(edit.end)
  }
  return out
}

/** queue edits for the c:v of each indexed cache point present in [range] */
function pushPointEdits(
  xml: string,
  range: { start: number; end: number },
  texts: (string | null)[],
  edits: Array<{ start: number; end: number; text: string }>,
): void {
  const re = /<c:pt idx="(\d+)"[^>]*>([\s\S]*?)<\/c:pt>/g
  re.lastIndex = range.start
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null && m.index < range.end) {
    const idx = parseInt(m[1], 10)
    const text = texts[idx]
    if (text === null || text === undefined) continue
    const inner = innerTextRanges(xml, 'c:v', m.index, m.index + m[0].length)
    if (inner.length > 0) edits.push({ ...inner[0], text })
  }
}

/** first depth-0 range of `tag` in [from, to); charts never nest these tags */
function tagRange(
  xml: string,
  tag: string,
  from = 0,
  to = xml.length,
): { start: number; end: number } | null {
  const openPrefix = '<' + tag
  let i = from
  while (i < to) {
    const o = xml.indexOf(openPrefix, i)
    if (o === -1 || o >= to) return null
    const after = xml.charAt(o + openPrefix.length)
    if (after !== '>' && after !== ' ' && after !== '/') {
      i = o + openPrefix.length // prefix of a longer tag (c:tx vs c:txPr)
      continue
    }
    const close = xml.indexOf('</' + tag + '>', o)
    if (close === -1 || close >= to) return null
    return { start: o, end: close + tag.length + 3 }
  }
  return null
}

/** all depth-0 ranges of `tag` (used for c:ser, which never nests) */
function tagRanges(xml: string, tag: string): Array<{ start: number; end: number }> {
  const out: Array<{ start: number; end: number }> = []
  let from = 0
  let range: { start: number; end: number } | null
  while ((range = tagRange(xml, tag, from)) !== null) {
    out.push(range)
    from = range.end
  }
  return out
}

/** inner-text ranges (between open and close tag) of every `tag` in [from, to) */
function innerTextRanges(
  xml: string,
  tag: string,
  from: number,
  to: number,
): Array<{ start: number; end: number }> {
  const out: Array<{ start: number; end: number }> = []
  const re = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`, 'g')
  re.lastIndex = from
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null && m.index < to) {
    const openEnd = m.index + m[0].indexOf('>') + 1
    out.push({ start: openEnd, end: openEnd + m[1].length })
  }
  return out
}

// ---- embedded xlsx workbook ----

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

/** Excel column letter: A, B, C... */
const xlsxColLetter = (i: number) => String.fromCharCode(65 + i)

/**
 * Build a minimal but valid xlsx file containing one Sheet1 with the chart
 * data (header row + data rows). Returns base64-encoded bytes.
 *
 * Layout:
 *   A1        | B1 (ser 0 name) | C1 (ser 1 name) ...
 *   A2 (cat0) | B2 (val 0,0)    | C2 (val 1,0)   ...
 *   A3 (cat1) | B3 (val 0,1)    | C3 (val 1,1)   ...
 */
export async function buildChartWorkbookXlsxBase64(
  categories: string[],
  series: Array<{ name: string; values: (number | null)[] }>,
): Promise<string> {
  const rows = categories.length
  const serCount = series.length

  // shared strings: all text cells collected in order
  const strings: string[] = []
  const si = (s: string): number => {
    const idx = strings.indexOf(s)
    if (idx !== -1) return idx
    strings.push(s)
    return strings.length - 1
  }

  // Build sheetData rows
  const headerCells: string[] = []
  // A1: empty label cell
  headerCells.push(`<c r="A1" t="s"><v>${si('')}</v></c>`)
  for (let j = 0; j < serCount; j++) {
    headerCells.push(`<c r="${xlsxColLetter(j + 1)}1" t="s"><v>${si(series[j].name)}</v></c>`)
  }
  const dataRows: string[] = []
  for (let i = 0; i < rows; i++) {
    const rowNum = i + 2
    const cells: string[] = []
    cells.push(`<c r="A${rowNum}" t="s"><v>${si(categories[i])}</v></c>`)
    for (let j = 0; j < serCount; j++) {
      const val = series[j].values[i]
      if (val !== null && val !== undefined) {
        cells.push(`<c r="${xlsxColLetter(j + 1)}${rowNum}"><v>${val}</v></c>`)
      }
    }
    dataRows.push(`<row r="${rowNum}">${cells.join('')}</row>`)
  }

  const sheetXml =
    XML_DECL +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData>' +
    `<row r="1">${headerCells.join('')}</row>` +
    dataRows.join('') +
    '</sheetData>' +
    '</worksheet>'

  const sharedStringsXml =
    XML_DECL +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
    strings.map((s) => `<si><t>${escapeXmlText(s)}</t></si>`).join('') +
    '</sst>'

  const workbookXml =
    XML_DECL +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
    '</workbook>'

  const workbookRels =
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
    '</Relationships>'

  const topRels =
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>'

  const contentTypes =
    XML_DECL +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '</Types>'

  const zip = new JSZip()
  zip.file('[Content_Types].xml', contentTypes)
  zip.file('_rels/.rels', topRels)
  zip.file('xl/workbook.xml', workbookXml)
  zip.file('xl/_rels/workbook.xml.rels', workbookRels)
  zip.file('xl/worksheets/sheet1.xml', sheetXml)
  zip.file('xl/sharedStrings.xml', sharedStringsXml)

  return zipToBase64(zip)
}

async function zipToBase64(zip: JSZip): Promise<string> {
  const bytes = await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  let binary = ''
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i])
  return btoa(binary)
}

/**
 * Patch the Sheet1 sheetData inside an embedded xlsx file (base64).
 * Rewrites category column A and each series column with the provided values.
 * Returns the updated base64, or null on failure.
 */
export async function patchChartWorkbookXlsxBase64(
  base64: string,
  categories: string[],
  series: Array<{ name: string; values: (number | null)[] }>,
): Promise<string | null> {
  try {
    const binaryStr = atob(base64)
    const bytes = new Uint8Array(binaryStr.length)
    for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i)

    const zip = await JSZip.loadAsync(bytes)
    const sheetFile = zip.file('xl/worksheets/sheet1.xml')
    if (!sheetFile) return null

    // Replace only Sheet1's sheetData in place, using inline strings so
    // sharedStrings.xml (and every other part: styles, extra sheets, defined
    // names) survives untouched.
    const inlineStr = (ref: string, text: string) =>
      `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXmlText(text)}</t></is></c>`
    const headerCells = [inlineStr('A1', '')]
    for (let j = 0; j < series.length; j++) {
      headerCells.push(inlineStr(`${xlsxColLetter(j + 1)}1`, series[j].name))
    }
    const dataRows: string[] = []
    for (let i = 0; i < categories.length; i++) {
      const rowNum = i + 2
      const cells = [inlineStr(`A${rowNum}`, categories[i])]
      for (let j = 0; j < series.length; j++) {
        const val = series[j].values[i]
        if (val !== null && val !== undefined) {
          cells.push(`<c r="${xlsxColLetter(j + 1)}${rowNum}"><v>${val}</v></c>`)
        }
      }
      dataRows.push(`<row r="${rowNum}">${cells.join('')}</row>`)
    }
    const newSheetData = `<sheetData><row r="1">${headerCells.join('')}</row>${dataRows.join('')}</sheetData>`

    const sheetXml = await sheetFile.async('string')
    if (!/<sheetData\/>|<sheetData[\s>]/.test(sheetXml)) return null
    let updatedSheet = sheetXml.replace(
      /<sheetData\/>|<sheetData[^>]*>[\s\S]*?<\/sheetData>/,
      newSheetData,
    )
    const lastRef = `${xlsxColLetter(series.length)}${categories.length + 1}`
    updatedSheet = updatedSheet.replace(/<dimension[^>]*\/>/, `<dimension ref="A1:${lastRef}"/>`)

    zip.file('xl/worksheets/sheet1.xml', updatedSheet)
    return await zipToBase64(zip)
  } catch {
    return null
  }
}
