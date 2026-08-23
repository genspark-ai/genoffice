/**
 * chartEx parsing (2014 cx namespace: ppt/charts/chartExN.xml → ChartModel).
 *
 * Modern chart types PowerPoint stores outside the classic c: namespace.
 * Supported layoutIds: funnel (categories + values → centered bars) and
 * sunburst (multi-level strDim hierarchy + leaf sizes → rings). Other
 * layoutIds (waterfall, treemap, …) return null and fall back to the chip.
 */
import { XMLParser } from 'fast-xml-parser'
import { type Theme } from './theme'
import { resolveColorNode } from './color'
import type { ChartModel } from './chart'

const cxParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  isArray: (name) => ['cx:data', 'cx:lvl', 'cx:pt', 'cx:series', 'cx:dataPt'].includes(name),
})

/** One cx:lvl: ptCount + sparse idx→text points → dense array ('' for gaps). */
function readLvl(lvl: any): string[] {
  const n = parseInt(lvl?.['@_ptCount'], 10) || 0
  const out: string[] = Array.from({ length: n }, () => '')
  const ptsRaw = lvl?.['cx:pt']
  const pts: any[] = Array.isArray(ptsRaw) ? ptsRaw : ptsRaw ? [ptsRaw] : []
  for (const pt of pts) {
    const i = parseInt(pt?.['@_idx'], 10)
    if (!Number.isNaN(i) && i < n) out[i] = typeof pt === 'object' ? String(pt['#text'] ?? '') : ''
  }
  return out
}

/** All a:t text values in a parsed subtree (handles xml:space="preserve" object form). */
function collectAT(node: unknown): string[] {
  const out: string[] = []
  const walk = (n: unknown) => {
    if (n == null || typeof n !== 'object') return
    for (const [k, v] of Object.entries(n as Record<string, unknown>)) {
      if (k.startsWith('@_')) continue
      if (k === 'a:t') {
        for (const t of Array.isArray(v) ? v : [v]) {
          if (typeof t === 'string') out.push(t)
          else if (t && typeof t === 'object')
            out.push(String((t as Record<string, unknown>)['#text'] ?? ''))
        }
      } else {
        for (const c of Array.isArray(v) ? v : [v]) walk(c)
      }
    }
  }
  walk(node)
  return out
}

export function parseChartExXml(xml: string, theme?: Theme): ChartModel | null {
  let doc: any
  try {
    doc = cxParser.parse(xml)
  } catch {
    return null
  }
  const space = doc['cx:chartSpace']
  const region = space?.['cx:chart']?.['cx:plotArea']?.['cx:plotAreaRegion']
  const seriesRaw = region?.['cx:series']
  const ser = Array.isArray(seriesRaw) ? seriesRaw[0] : seriesRaw
  const layoutId = ser?.['@_layoutId']
  if (layoutId !== 'funnel' && layoutId !== 'sunburst') return null

  const dataId = ser?.['cx:dataId']?.['@_val'] ?? '0'
  const datas: any[] = space?.['cx:chartData']?.['cx:data'] ?? []
  const data = datas.find((d) => String(d?.['@_id']) === String(dataId)) ?? datas[0]
  if (!data) return null
  const strLvlsRaw = data['cx:strDim']?.['cx:lvl'] ?? []
  const strLvls: string[][] = (Array.isArray(strLvlsRaw) ? strLvlsRaw : [strLvlsRaw]).map(readLvl)
  const numLvlRaw = data['cx:numDim']?.['cx:lvl']
  const numLvl = Array.isArray(numLvlRaw) ? numLvlRaw[0] : numLvlRaw
  const values = readLvl(numLvl).map((v) => {
    const n = parseFloat(v)
    return Number.isFinite(n) ? n : null
  })
  if (!values.some((v) => v != null)) return null

  const model: ChartModel = {
    kind: layoutId,
    categories: strLvls[0] ?? [],
    series: [{ values }],
  }
  // <cx:title/> with no rich text still renders as PowerPoint's placeholder title
  const title = space?.['cx:chart']?.['cx:title']
  if (title !== undefined) {
    model.title = collectAT(title?.['cx:tx']).join('') || 'Chart Title'
  }
  if (layoutId === 'sunburst') {
    model.sunburst = { levels: strLvls, sizes: values }
    const dPtsRaw = ser?.['cx:dataPt'] ?? []
    const dPts: any[] = Array.isArray(dPtsRaw) ? dPtsRaw : [dPtsRaw]
    const pointColors: Array<string | undefined> = []
    for (const dPt of dPts) {
      const i = parseInt(dPt?.['@_idx'], 10)
      const c = resolveColorNode(dPt?.['cx:spPr']?.['a:solidFill'], theme)
      if (!Number.isNaN(i) && c != null) pointColors[i] = c
    }
    if (pointColors.length) model.sunburst.pointColors = pointColors
  }
  if (layoutId === 'funnel') {
    const gap = parseFloat(
      space?.['cx:chart']?.['cx:plotArea']?.['cx:axis']?.['cx:catScaling']?.['@_gapWidth'],
    )
    if (Number.isFinite(gap) && gap >= 0) model.gapWidthPct = gap * 100
  }
  if (theme) {
    const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
      .map((k) => theme.colors?.[k])
      .filter((c): c is string => !!c)
    if (accents.length === 6) model.themePalette = accents
  }
  return model
}
