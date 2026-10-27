/// Slicer / timeline OOXML persistence (MS-XLSX x14 slicers, x15 timelines).
///
/// Write side: session slicer/timeline panels are materialized as
///   xl/slicerCaches/slicerCacheN.xml   (pivot binding + member selection)
///   xl/slicers/slicerN.xml             (the control: name/caption/geometry)
///   xl/timelineCaches/timelineCacheN.xml + xl/timelines/timelineN.xml
/// wired up by
///   worksheet rels (…/2007/relationships/slicer|timeline) + worksheet
///     extLst x14:slicerList / x15:slicerList ext entries,
///   workbook rels (…/slicerCache|timelineCache) + workbook extLst
///     x14:slicerCaches / x15:slicerCaches ext entries,
///   [Content_Types].xml overrides per part.
///
/// Read side mirrors the writer exactly (the reader keys on container element
/// names and relationship types, not on ext URIs, so a URI mismatch between
/// writers degrades to "control not restored" instead of corrupt state).
/// Excel-side fidelity of the ext URIs is best-effort: a real-Excel round
/// trip cannot be verified in this environment (declared in the PR body).
///
/// Pivot binding: slicerCacheDefinition carries the spec'd tabular@pivotCacheId
/// (resolved through workbook pivotCaches) and a pivotTables/pivotTable entry
/// whose tabId is the hosting worksheet's sheetId and whose name is the pivot
/// table's own name (MS-XLSX) — Excel does not bind the slicer to its pivot
/// without them. The reader prefers pivotCacheId and falls back to the tabId
/// rule over sorted pivot parts for foreign files.

import {
  allocatePartPath,
  appendRelationship,
  registerContentTypeOverride,
  relsPathFor,
  relativeTarget,
  resolveRelTarget,
  type MutablePackage,
} from './xlsx-drawing-add'
import { parseRelationships, removePartOverride, removeRelationshipById } from './xlsx-sheets'

export class SlicerAddError extends Error {}

// ─── Constants ─────────────────────────────────────────────────────────────

const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const X15_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export const SLICER_REL_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
export const TIMELINE_REL_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/timeline'
export const SLICER_CACHE_REL_TYPE =
  'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
export const TIMELINE_CACHE_REL_TYPE =
  'http://schemas.microsoft.com/office/2007/relationships/timelineCache'
/// ECMA-376: the worksheet → pivotTable part relationship, used to find which
/// sheet hosts a pivot (the slicerCache binding's tabId is that sheet's sheetId)
const PIVOT_TABLE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'

const SLICER_CONTENT_TYPE = 'application/vnd.ms-excel.slicer+xml'
const SLICER_CACHE_CONTENT_TYPE = 'application/vnd.ms-excel.slicerCache+xml'
const TIMELINE_CONTENT_TYPE = 'application/vnd.ms-excel.timeline+xml'
const TIMELINE_CACHE_CONTENT_TYPE = 'application/vnd.ms-excel.timelineCache+xml'

/// Ext URIs per MS-XLSX; the reader below does not depend on them (it matches
/// container names + relationship types), so a wrong URI can only affect
/// Excel, never our own round trip. Timelines do NOT share the slicer
/// containers: the worksheet carries x15:timelineRefs and the workbook
/// x15:timelineCacheRefs, each with its own URI.
const WORKSHEET_SLICER_LIST_URI = '{A8765BA9-456A-4dab-B4F3-ACF1056F45AA}'
const WORKSHEET_TIMELINE_LIST_URI = '{7E03D99C-DC04-49d9-9315-930204A7B6E9}'
const WORKBOOK_SLICER_CACHES_URI = '{BBE1A952-AA13-448e-AADC-164F8A28A991}'
const WORKBOOK_TIMELINE_CACHES_URI = '{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

// ─── Small XML helpers ─────────────────────────────────────────────────────

function escapeAttribute(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function decodeAttribute(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/// Same tolerance as xlsx-sheets readAttribute: double-quoted attributes.
function attr(element: string, name: string): string | undefined {
  return new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(element)?.[1]
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/// Excel name sanitation: no punctuation beyond underscore, cannot start with
/// a digit, bounded length (ECMA-376 name rules as Excel enforces them).
function sanitizeNameBase(source: string): string {
  const cleaned = source
    .replace(/[^A-Za-z0-9_]/g, '_')
    .replace(/^[0-9_]+/, 'S$&')
    .slice(0, 200)
  return cleaned === '' ? 'Field' : cleaned
}

function uniqueName(base: string, taken: Set<string>): string {
  if (!taken.has(base)) {
    taken.add(base)
    return base
  }
  for (let index = 2; ; index += 1) {
    const candidate = `${base}_${index}`
    if (!taken.has(candidate)) {
      taken.add(candidate)
      return candidate
    }
  }
}

// ─── Write-side model ──────────────────────────────────────────────────────

/// One selectable member: `x` is the sharedItems index of the member in the
/// bound pivot cache field (what slicerCache items/@x stores).
export interface SlicerItemWrite {
  readonly x: number
  readonly selected: boolean
}

export interface SlicerControlWrite {
  /// Worksheet part that hosts the control (its rels + extLst get wired).
  readonly worksheetPath: string
  /// Bound pivotTable part path; supplies cacheId + tabId.
  readonly pivotPath: string
  /// slicerCacheDefinition/@sourceName — the bound cache field's name.
  readonly sourceName: string
  /// Control caption shown in the header.
  readonly caption: string
  /// slicer@name — stable file identity from a previous save (restore).
  readonly name?: string | undefined
  /// Existing part paths from a previous save; present = update in place.
  readonly partPath?: string | undefined
  readonly cachePath?: string | undefined
  /// slicerCacheDefinition/@name from a previous save.
  readonly cacheName?: string | undefined
  /// Every member of the field with its selection state.
  readonly items: readonly SlicerItemWrite[]
}

export interface TimelineControlWrite {
  readonly worksheetPath: string
  readonly pivotPath: string
  readonly sourceName: string
  readonly caption: string
  readonly name?: string | undefined
  readonly partPath?: string | undefined
  readonly cachePath?: string | undefined
  readonly cacheName?: string | undefined
  /// Selected inclusive month range (year*12+month-1); null = unfiltered.
  readonly range: { readonly startMonth: number; readonly endMonth: number } | null
}

export interface SlicerRemovalWrite {
  readonly kind: 'slicer' | 'timeline'
  readonly partPath: string
  readonly cachePath: string
}

export interface SlicerStateWrite {
  readonly slicers: readonly SlicerControlWrite[]
  readonly timelines: readonly TimelineControlWrite[]
  readonly removals: readonly SlicerRemovalWrite[]
}

// ─── Part XML builders ─────────────────────────────────────────────────────

/// SlicerCache/timelineCache pivot binding per MS-XLSX: pivotTables/pivotTable
/// carries the hosting worksheet's sheetId as tabId and the pivot table's own
/// name. The hosting worksheet is the one whose rels reference the pivot part.
async function pivotBinding(
  pkg: MutablePackage,
  workbookXml: string,
  pivotPath: string,
): Promise<{ tabId: number; name: string }> {
  if (!(await pkg.has(pivotPath))) {
    throw new SlicerAddError(`Slicer pivot part ${pivotPath} is missing from the workbook.`)
  }
  const name = attr(await pkg.readText(pivotPath), 'name')
  if (name === undefined) {
    throw new SlicerAddError(`Pivot part ${pivotPath} has no pivot table name.`)
  }
  let hostPath: string | null = null
  for (const path of await pkg.paths()) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(path)) continue
    const relsPath = relsPathFor(path)
    if (!(await pkg.has(relsPath))) continue
    if (findRelationshipTo(await pkg.readText(relsPath), PIVOT_TABLE_REL_TYPE, path, pivotPath)) {
      hostPath = path
      break
    }
  }
  if (hostPath === null) {
    throw new SlicerAddError(`Pivot part ${pivotPath} is not hosted by any worksheet.`)
  }
  const relsXml = (await pkg.has('xl/_rels/workbook.xml.rels'))
    ? await pkg.readText('xl/_rels/workbook.xml.rels')
    : ''
  const rels = parseRelationships(relsXml)
  for (const sheet of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const sheetId = attr(sheet[0], 'sheetId')
    const relId = attr(sheet[0], 'r:id')
    if (sheetId === undefined || relId === undefined) continue
    const entry = rels.find((rel) => rel.id === relId)
    if (entry === undefined || resolveRelTarget('xl/workbook.xml', entry.target) !== hostPath) {
      continue
    }
    const parsed = Number(sheetId)
    if (!Number.isInteger(parsed) || parsed < 1) {
      throw new SlicerAddError(`Sheet ${hostPath} carries no usable sheetId.`)
    }
    return { tabId: parsed, name: decodeAttribute(name) }
  }
  throw new SlicerAddError(`The workbook does not list a sheet for ${hostPath}.`)
}

async function pivotCacheIdOf(pkg: MutablePackage, pivotPath: string): Promise<number> {
  const cacheId = attr(await pkg.readText(pivotPath), 'cacheId')
  const parsed = cacheId === undefined ? Number.NaN : Number(cacheId)
  if (!Number.isInteger(parsed) || parsed < 0) {
    throw new SlicerAddError(`Pivot part ${pivotPath} has no cacheId.`)
  }
  return parsed
}

function monthIso(key: number, lastDay: boolean): string {
  const year = Math.floor(key / 12)
  const month = (key % 12) + 1
  const day = lastDay ? new Date(Date.UTC(year, month, 0)).getUTCDate() : 1
  return (
    `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-` +
    `${String(day).padStart(2, '0')}T00:00:00`
  )
}

/// Geometry attributes worth keeping when rewriting an Excel-authored part
/// (our session model has no geometry; the file's survives the rewrite).
function carryOverGeometry(existingXml: string, tag: 'slicer' | 'timeline'): string {
  const element = new RegExp(`<${tag}(?:\\s[^>]*)?/>`).exec(existingXml)?.[0]
  if (element === undefined) return ''
  let carried = ''
  for (const name of ['rowHeight', 'columnCount', 'style', 'showCaption', 'startItem', 'level']) {
    const value = attr(element, name)
    if (value !== undefined) carried += ` ${name}="${escapeAttribute(value)}"`
  }
  return carried
}

function buildSlicerPartXml(name: string, cacheName: string, caption: string, carried: string) {
  // rowHeight is required for Excel to lay the control out; a fresh control
  // gets Excel's own default (241300 EMU ≈ a quarter inch), an updated one
  // keeps whatever the file carried.
  const geometry = carried.includes('rowHeight') ? carried : `${carried} rowHeight="241300"`
  return (
    `${XML_DECL}<slicers xmlns="${X14_NS}">` +
    `<slicer name="${escapeAttribute(name)}" cache="${escapeAttribute(cacheName)}"` +
    ` caption="${escapeAttribute(caption)}"${geometry}/>` +
    '</slicers>'
  )
}

function buildTimelinePartXml(name: string, cacheName: string, caption: string, carried: string) {
  // level/selectionLevel are required; our session model is month-granular
  // (level 2 in the year/quarter/month/day ladder)
  const geometry = carried.includes('level') ? carried : `${carried} level="2" selectionLevel="2"`
  return (
    `${XML_DECL}<timelines xmlns="${X15_NS}">` +
    `<timeline name="${escapeAttribute(name)}" cache="${escapeAttribute(cacheName)}"` +
    ` caption="${escapeAttribute(caption)}"${geometry}/>` +
    '</timelines>'
  )
}

function buildSlicerCacheXml(
  cacheName: string,
  sourceName: string,
  binding: { readonly tabId: number; readonly name: string },
  pivotCacheId: number,
  items: readonly SlicerItemWrite[],
): string {
  const itemsXml = items
    .map((item) => `<i x="${item.x}"${item.selected ? ' s="1"' : ''}/>`)
    .join('')
  return (
    `${XML_DECL}<slicerCacheDefinition xmlns="${X14_NS}" xmlns:r="${REL_NS}"` +
    ` name="${escapeAttribute(cacheName)}" sourceName="${escapeAttribute(sourceName)}">` +
    `<pivotTables><pivotTable tabId="${binding.tabId}" name="${escapeAttribute(binding.name)}"/></pivotTables>` +
    '<data>' +
    `<tabular pivotCacheId="${pivotCacheId}" crossFilter="showItemsWithDataAtTop">` +
    `<items count="${items.length}">${itemsXml}</items>` +
    '</tabular></data></slicerCacheDefinition>'
  )
}

function buildTimelineCacheXml(
  cacheName: string,
  sourceName: string,
  binding: { readonly tabId: number; readonly name: string },
  pivotCacheId: number,
  range: { readonly startMonth: number; readonly endMonth: number } | null,
): string {
  // The selection window lives on the cache; our month-granularity model
  // serializes as inclusive first/last-of-month ISO datetimes, carried in the
  // spec's <bounds> element (the reader also accepts <selection>, the shape
  // this writer shipped before the state attributes existed). The state
  // attributes are what Excel requires to open the cache: refresh versions,
  // the bound pivot cache, and how the filter applies.
  const inner =
    range === null
      ? ''
      : `<selection startDate="${monthIso(range.startMonth, false)}"` +
        ` endDate="${monthIso(range.endMonth, true)}"/>` +
        `<bounds startDate="${monthIso(range.startMonth, false)}"` +
        ` endDate="${monthIso(range.endMonth, true)}"/>`
  return (
    `${XML_DECL}<timelineCacheDefinition xmlns="${X15_NS}" xmlns:r="${REL_NS}"` +
    ` name="${escapeAttribute(cacheName)}" sourceName="${escapeAttribute(sourceName)}">` +
    `<pivotTables><pivotTable tabId="${binding.tabId}" name="${escapeAttribute(binding.name)}"/></pivotTables>` +
    `<state minimalRefreshVersion="1" lastRefreshVersion="1" pivotCacheId="${pivotCacheId}"` +
    ` filterType="${range === null ? 'none' : 'dateBetween'}">${inner}</state>` +
    '</timelineCacheDefinition>'
  )
}

// ─── extLst surgery (worksheet and workbook share the container shape) ─────

interface ExtWiring {
  /// ext uri attribute value.
  readonly uri: string
  /// x14 / x15 namespace prefix and URI.
  readonly prefix: 'x14' | 'x15'
  readonly namespace: string
  /// slicerList (worksheet) / slicerCaches (workbook).
  readonly container: string
  /// slicer / slicerCache.
  readonly item: string
}

const WORKSHEET_SLICER_WIRING: ExtWiring = {
  uri: WORKSHEET_SLICER_LIST_URI,
  prefix: 'x14',
  namespace: X14_NS,
  container: 'slicerList',
  item: 'slicer',
}
const WORKSHEET_TIMELINE_WIRING: ExtWiring = {
  uri: WORKSHEET_TIMELINE_LIST_URI,
  prefix: 'x15',
  namespace: X15_NS,
  container: 'timelineRefs',
  item: 'timelineRef',
}
const WORKBOOK_SLICER_WIRING: ExtWiring = {
  uri: WORKBOOK_SLICER_CACHES_URI,
  prefix: 'x14',
  namespace: X14_NS,
  container: 'slicerCaches',
  item: 'slicerCache',
}
const WORKBOOK_TIMELINE_WIRING: ExtWiring = {
  uri: WORKBOOK_TIMELINE_CACHES_URI,
  prefix: 'x15',
  namespace: X15_NS,
  container: 'timelineCacheRefs',
  item: 'timelineCacheRef',
}

function extOpenTag(xml: string, wiring: ExtWiring): { start: number; end: number } | null {
  const match = new RegExp(`<ext\\b[^>]*\\buri="${escapeRegExp(wiring.uri)}"[^>]*>`, 'i').exec(xml)
  if (!match) return null
  return { start: match.index, end: match.index + match[0].length }
}

/// Adds `<{prefix}:{item} r:id="…"/>` inside the wiring's ext (creating the
/// ext, its container, and the host extLst as needed). For worksheets the
/// `</sheetData>` anchor disambiguates the worksheet-level extLst from cell
/// extLst elements that sit inside sheetData.
function addExtRef(
  xml: string,
  wiring: ExtWiring,
  relId: string,
  hostClose: '</worksheet>' | '</workbook>',
): string {
  const entry = `<${wiring.prefix}:${wiring.item} r:id="${relId}"/>`
  const container = `${wiring.prefix}:${wiring.container}`
  const ext = extOpenTag(xml, wiring)
  if (ext) {
    const closeAt = xml.indexOf('</ext>', ext.end)
    if (closeAt === -1) throw new SlicerAddError('A slicer ext element is malformed.')
    const block = xml.slice(ext.start, closeAt)
    if (block.includes(`r:id="${relId}"`)) return xml
    const containerClose = `</${container}>`
    const containerAt = block.lastIndexOf(containerClose)
    const patched =
      containerAt === -1
        ? block + `<${container}>${entry}</${container}>`
        : block.slice(0, containerAt) + entry + block.slice(containerAt)
    return xml.slice(0, ext.start) + patched + xml.slice(closeAt)
  }
  const element =
    `<ext uri="${wiring.uri}" xmlns:${wiring.prefix}="${wiring.namespace}"` +
    ` xmlns:r="${REL_NS}">` +
    `<${container}>${entry}</${container}></ext>`
  const extLstClose = xml.lastIndexOf('</extLst>')
  const worksheetLevel =
    hostClose === '</worksheet>' ? extLstClose > xml.lastIndexOf('</sheetData>') : true
  if (extLstClose !== -1 && worksheetLevel) {
    return xml.slice(0, extLstClose) + element + xml.slice(extLstClose)
  }
  const hostCloseAt = xml.lastIndexOf(hostClose)
  if (hostCloseAt === -1) throw new SlicerAddError('The host element is malformed.')
  return xml.slice(0, hostCloseAt) + `<extLst>${element}</extLst>` + xml.slice(hostCloseAt)
}

/// Removes the ext entry; prunes the container, the ext, and the extLst when
/// they become empty. Returns the input unchanged when the entry is absent.
/// All surgery is scoped to OUR ext block — worksheets carry foreign ext
/// elements of their own (cells, sparklines) that must not be touched.
function removeExtRef(xml: string, wiring: ExtWiring, relId: string): string {
  const ext = extOpenTag(xml, wiring)
  if (!ext) return xml
  const closeAt = xml.indexOf('</ext>', ext.end)
  if (closeAt === -1) throw new SlicerAddError('A slicer ext element is malformed.')
  const block = xml.slice(ext.start, closeAt + '</ext>'.length)
  const entryPattern = new RegExp(
    `<${wiring.prefix}:${wiring.item}\\b[^>]*\\br:id="${escapeRegExp(relId)}"[^>]*/>\\s*`,
  )
  const strippedEntry = block.replace(entryPattern, '')
  if (strippedEntry === block) return xml

  // Prune an emptied container inside the block.
  let inner = strippedEntry.slice(0, strippedEntry.length - '</ext>'.length)
  const containerOpen = new RegExp(`<${wiring.prefix}:${wiring.container}\\b[^>]*>`)
  const containerMatch = containerOpen.exec(inner)
  if (containerMatch) {
    const containerClose = `</${wiring.prefix}:${wiring.container}>`
    const containerCloseAt = inner.lastIndexOf(containerClose)
    const containerInner = inner.slice(containerMatch.index, containerCloseAt)
    if (!new RegExp(`<${wiring.prefix}:${wiring.item}\\b`).test(containerInner)) {
      inner =
        inner.slice(0, containerMatch.index) + inner.slice(containerCloseAt + containerClose.length)
    }
  }
  // An emptied ext (open tag to close tag) drops out entirely.
  const extInner = inner.slice(ext.end - ext.start)
  let replacement = ''
  if (extInner.trim() !== '') {
    replacement = inner + '</ext>'
  }
  let patched = xml.slice(0, ext.start) + replacement + xml.slice(closeAt + '</ext>'.length)
  if (replacement === '') {
    const lstClose = patched.lastIndexOf('</extLst>')
    if (lstClose !== -1) {
      const lstOpen = patched.lastIndexOf('<extLst>', lstClose)
      if (lstOpen !== -1 && patched.slice(lstOpen + '<extLst>'.length, lstClose).trim() === '') {
        patched = patched.slice(0, lstOpen) + patched.slice(lstClose + '</extLst>'.length)
      }
    }
  }
  return patched
}

// ─── Relationship helpers ──────────────────────────────────────────────────

/// Relationship of `type` in `relsXml` whose resolved target is `partPath`.
function findRelationshipTo(
  relsXml: string,
  type: string,
  fromPart: string,
  partPath: string,
): { id: string; target: string } | null {
  for (const entry of parseRelationships(relsXml)) {
    if (entry.type !== type || entry.id === undefined) continue
    if (resolveRelTarget(fromPart, entry.target) === partPath) {
      return { id: entry.id, target: entry.target }
    }
  }
  return null
}

async function appendOrReuseRelationship(
  pkg: MutablePackage,
  relsPath: string,
  type: string,
  fromPart: string,
  partPath: string,
  touchedEntries: Set<string>,
): Promise<string> {
  if (await pkg.has(relsPath)) {
    const relsXml = await pkg.readText(relsPath)
    const existing = findRelationshipTo(relsXml, type, fromPart, partPath)
    if (existing) return existing.id
  }
  const id = await appendRelationship(pkg, relsPath, type, relativeTarget(fromPart, partPath))
  touchedEntries.add(relsPath)
  return id
}

async function ensureWorksheetRef(
  pkg: MutablePackage,
  worksheetPath: string,
  partPath: string,
  relType: string,
  wiring: ExtWiring,
  touchedEntries: Set<string>,
): Promise<void> {
  const relId = await appendOrReuseRelationship(
    pkg,
    relsPathFor(worksheetPath),
    relType,
    worksheetPath,
    partPath,
    touchedEntries,
  )
  const worksheetXml = await pkg.readText(worksheetPath)
  const patched = addExtRef(worksheetXml, wiring, relId, '</worksheet>')
  if (patched !== worksheetXml) {
    pkg.write(worksheetPath, patched)
    touchedEntries.add(worksheetPath)
  }
}

async function removeWorksheetRef(
  pkg: MutablePackage,
  worksheetPath: string,
  partPath: string,
  relType: string,
  wiring: ExtWiring,
  touchedEntries: Set<string>,
): Promise<void> {
  const relsPath = relsPathFor(worksheetPath)
  let relId: string | undefined
  if (await pkg.has(relsPath)) {
    const relsXml = await pkg.readText(relsPath)
    const existing = findRelationshipTo(relsXml, relType, worksheetPath, partPath)
    if (existing) {
      pkg.write(relsPath, removeRelationshipById(relsXml, existing.id))
      touchedEntries.add(relsPath)
      relId = existing.id
    }
  }
  if (!(await pkg.has(worksheetPath))) return
  const worksheetXml = await pkg.readText(worksheetPath)
  const patched = relId === undefined ? worksheetXml : removeExtRef(worksheetXml, wiring, relId)
  if (patched !== worksheetXml) {
    pkg.write(worksheetPath, patched)
    touchedEntries.add(worksheetPath)
  }
}

async function removePartWithSatellites(
  pkg: MutablePackage,
  path: string,
  touchedEntries: Set<string>,
): Promise<void> {
  if (!(await pkg.has(path))) return
  pkg.remove(path)
  touchedEntries.add(path)
  // Satellite rels of the removed part die with it (our writer creates
  // none; a foreign part's rels are removed rather than orphaned).
  const relsPath = relsPathFor(path)
  if (await pkg.has(relsPath)) {
    pkg.remove(relsPath)
    touchedEntries.add(relsPath)
  }
  const contentTypesPath = '[Content_Types].xml'
  const xml = await pkg.readText(contentTypesPath)
  const stripped = removePartOverride(xml, path)
  if (stripped !== xml) {
    pkg.write(contentTypesPath, stripped)
    touchedEntries.add(contentTypesPath)
  }
}

// ─── Existing-name collection (merge never collides) ───────────────────────

async function collectCacheNames(pkg: MutablePackage, directory: string): Promise<Set<string>> {
  const names = new Set<string>()
  for (const path of await pkg.paths()) {
    if (!path.startsWith(directory) || !path.endsWith('.xml')) continue
    const name = attr(await pkg.readText(path), 'name')
    if (name !== undefined) names.add(decodeAttribute(name))
  }
  return names
}

async function collectControlNames(
  pkg: MutablePackage,
  directory: string,
  tag: 'slicer' | 'timeline',
): Promise<Set<string>> {
  const names = new Set<string>()
  for (const path of await pkg.paths()) {
    if (!path.startsWith(directory) || !path.endsWith('.xml')) continue
    for (const match of (await pkg.readText(path)).matchAll(
      new RegExp(`<${tag}(?:\\s[^>]*)?/>`, 'g'),
    )) {
      const name = attr(match[0], 'name')
      if (name !== undefined) names.add(decodeAttribute(name))
    }
  }
  return names
}

/// Which worksheet parts carry a rel to `partPath`. Our own files reference a
/// control from exactly one worksheet; the walk stays general so a foreign
/// file's references all get stripped on removal.
async function worksheetsReferencing(
  pkg: MutablePackage,
  partPath: string,
  relType: string,
): Promise<string[]> {
  const hosts: string[] = []
  for (const path of await pkg.paths()) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(path)) continue
    const relsPath = relsPathFor(path)
    if (!(await pkg.has(relsPath))) continue
    if (findRelationshipTo(await pkg.readText(relsPath), relType, path, partPath) !== null) {
      hosts.push(path)
    }
  }
  return hosts
}

// ─── Write entry point ─────────────────────────────────────────────────────

async function writeControl(
  pkg: MutablePackage,
  control: SlicerControlWrite | TimelineControlWrite,
  kind: 'slicer' | 'timeline',
  workbookXml: string,
  touchedEntries: Set<string>,
  takenControlNames: Set<string>,
  takenCacheNames: Set<string>,
): Promise<string> {
  if (!(await pkg.has(control.worksheetPath))) {
    throw new SlicerAddError(`Slicer host worksheet ${control.worksheetPath} is missing.`)
  }
  const binding = await pivotBinding(pkg, workbookXml, control.pivotPath)
  const pivotCacheId = await pivotCacheIdOf(pkg, control.pivotPath)

  // Identity: reuse the restored part paths / names, else allocate fresh ones
  // that do not collide with anything already in the package.
  const partPath =
    control.partPath ??
    (await allocatePartPath(
      pkg,
      kind === 'slicer' ? 'xl/slicers/slicer' : 'xl/timelines/timeline',
      '.xml',
    ))
  const cachePath =
    control.cachePath ??
    (await allocatePartPath(
      pkg,
      kind === 'slicer' ? 'xl/slicerCaches/slicerCache' : 'xl/timelineCaches/timelineCache',
      '.xml',
    ))
  const cacheName =
    control.cacheName ??
    uniqueName(
      `${kind === 'slicer' ? 'Slicer' : 'Timeline'}_${sanitizeNameBase(control.sourceName)}`,
      takenCacheNames,
    )
  const name =
    control.name ??
    uniqueName(sanitizeNameBase(control.caption || control.sourceName), takenControlNames)

  const cacheXml =
    kind === 'slicer'
      ? buildSlicerCacheXml(
          cacheName,
          control.sourceName,
          binding,
          pivotCacheId,
          (control as SlicerControlWrite).items,
        )
      : buildTimelineCacheXml(
          cacheName,
          control.sourceName,
          binding,
          pivotCacheId,
          (control as TimelineControlWrite).range,
        )
  if (await pkg.has(cachePath)) pkg.write(cachePath, cacheXml)
  else pkg.add(cachePath, cacheXml)
  touchedEntries.add(cachePath)
  await registerContentTypeOverride(
    pkg,
    cachePath,
    kind === 'slicer' ? SLICER_CACHE_CONTENT_TYPE : TIMELINE_CACHE_CONTENT_TYPE,
    touchedEntries,
  )

  // The workbook rel + ext entry point at the cache part.
  const cacheRelType = kind === 'slicer' ? SLICER_CACHE_REL_TYPE : TIMELINE_CACHE_REL_TYPE
  const relId = await appendOrReuseRelationship(
    pkg,
    'xl/_rels/workbook.xml.rels',
    cacheRelType,
    'xl/workbook.xml',
    cachePath,
    touchedEntries,
  )
  const wiredWorkbook = addExtRef(
    workbookXml,
    kind === 'slicer' ? WORKBOOK_SLICER_WIRING : WORKBOOK_TIMELINE_WIRING,
    relId,
    '</workbook>',
  )

  // The control part (rewrites an existing one, preserving its geometry).
  const existingPart = (await pkg.has(partPath)) ? await pkg.readText(partPath) : null
  const carried = existingPart === null ? '' : carryOverGeometry(existingPart, kind)
  const partXml =
    kind === 'slicer'
      ? buildSlicerPartXml(name, cacheName, control.caption, carried)
      : buildTimelinePartXml(name, cacheName, control.caption, carried)
  if (existingPart === null) pkg.add(partPath, partXml)
  else pkg.write(partPath, partXml)
  touchedEntries.add(partPath)
  await registerContentTypeOverride(
    pkg,
    partPath,
    kind === 'slicer' ? SLICER_CONTENT_TYPE : TIMELINE_CONTENT_TYPE,
    touchedEntries,
  )

  await ensureWorksheetRef(
    pkg,
    control.worksheetPath,
    partPath,
    kind === 'slicer' ? SLICER_REL_TYPE : TIMELINE_REL_TYPE,
    kind === 'slicer' ? WORKSHEET_SLICER_WIRING : WORKSHEET_TIMELINE_WIRING,
    touchedEntries,
  )
  return wiredWorkbook
}

/// Materializes every slicer/timeline control and cleans removals. Returns
/// the workbook XML with the cache ext entries merged in. Removals run first
/// so a deleted control's parts free their names for reuse.
export async function applySlicerState(
  pkg: MutablePackage,
  state: SlicerStateWrite,
  workbookXml: string,
  touchedEntries: Set<string>,
): Promise<string> {
  const written = new Set(
    [...state.slicers, ...state.timelines].flatMap((control) =>
      [
        ...(control.partPath ? [control.partPath] : []),
        ...(control.cachePath ? [control.cachePath] : []),
      ].filter((path) =>
        state.removals.some((removal) => removal.partPath === path || removal.cachePath === path),
      ),
    ),
  )
  if (written.size > 0) {
    throw new SlicerAddError('A removed slicer part cannot also be written in the same save.')
  }

  for (const removal of state.removals) {
    const relType = removal.kind === 'slicer' ? SLICER_REL_TYPE : TIMELINE_REL_TYPE
    const cacheRelType = removal.kind === 'slicer' ? SLICER_CACHE_REL_TYPE : TIMELINE_CACHE_REL_TYPE
    const wiring = removal.kind === 'slicer' ? WORKSHEET_SLICER_WIRING : WORKSHEET_TIMELINE_WIRING
    const cacheWiring =
      removal.kind === 'slicer' ? WORKBOOK_SLICER_WIRING : WORKBOOK_TIMELINE_WIRING
    for (const host of await worksheetsReferencing(pkg, removal.partPath, relType)) {
      await removeWorksheetRef(pkg, host, removal.partPath, relType, wiring, touchedEntries)
    }
    const relsPath = 'xl/_rels/workbook.xml.rels'
    if (await pkg.has(relsPath)) {
      const relsXml = await pkg.readText(relsPath)
      const existing = findRelationshipTo(
        relsXml,
        cacheRelType,
        'xl/workbook.xml',
        removal.cachePath,
      )
      if (existing) {
        pkg.write(relsPath, removeRelationshipById(relsXml, existing.id))
        touchedEntries.add(relsPath)
        workbookXml = removeExtRef(workbookXml, cacheWiring, existing.id)
      }
    }
    await removePartWithSatellites(pkg, removal.partPath, touchedEntries)
    await removePartWithSatellites(pkg, removal.cachePath, touchedEntries)
  }

  if (state.slicers.length > 0) {
    const takenNames = await collectControlNames(pkg, 'xl/slicers/', 'slicer')
    const takenCaches = await collectCacheNames(pkg, 'xl/slicerCaches/')
    for (const slicer of state.slicers) {
      workbookXml = await writeControl(
        pkg,
        slicer,
        'slicer',
        workbookXml,
        touchedEntries,
        takenNames,
        takenCaches,
      )
    }
  }
  if (state.timelines.length > 0) {
    const takenNames = await collectControlNames(pkg, 'xl/timelines/', 'timeline')
    const takenCaches = await collectCacheNames(pkg, 'xl/timelineCaches/')
    for (const timeline of state.timelines) {
      workbookXml = await writeControl(
        pkg,
        timeline,
        'timeline',
        workbookXml,
        touchedEntries,
        takenNames,
        takenCaches,
      )
    }
  }
  return workbookXml
}

// ─── Read side ─────────────────────────────────────────────────────────────

export interface SlicerPartState {
  readonly kind: 'slicer' | 'timeline'
  /// Sidecar sheet id (`sheet-<workbook sheetId attribute>`).
  readonly sheetId: string
  readonly worksheetPath: string
  readonly partPath: string
  readonly cachePath: string
  /// slicer@name / timeline@name.
  readonly name: string
  /// slicerCacheDefinition@name / timelineCacheDefinition@name.
  readonly cacheName: string
  /// slicerCacheDefinition@sourceName (the bound cache field's name).
  readonly sourceName: string
  readonly caption: string
  /// pivotCacheDefinition part path resolved from the cache binding; null
  /// when unresolvable.
  readonly pivotCacheDefinitionPath: string | null
  /// Slicer: sharedItems indices of the selected members; null = the cache
  /// carried no item list (treat as unfiltered).
  readonly selectedItems: readonly number[] | null
  /// Timeline: inclusive month range; null = unfiltered.
  readonly range: { readonly startMonth: number; readonly endMonth: number } | null
}

export interface SlicerPartReader {
  paths(): Promise<readonly string[]>
  readText(path: string): Promise<string>
}

function monthKeyOfIso(value: string): number | null {
  const match = /^(\d{4})-(\d{2})/.exec(value)
  if (!match) return null
  const year = Number(match[1])
  const month = Number(match[2])
  if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) return null
  return year * 12 + month - 1
}

interface ParsedCachePart {
  readonly path: string
  readonly name: string
  readonly sourceName: string
  /// tabular@pivotCacheId (slicers) or root@pivotCacheId (timelines, our own
  /// writer); null when the part carries none.
  readonly pivotCacheId: number | null
  /// pivotTables/pivotTable@tabId, for the fallback binding.
  readonly tabId: number | null
  /// Slicer: true when the data section is tabular (OLAP caches cannot be
  /// represented by the session model).
  readonly tabular: boolean
  readonly selectedItems: readonly number[] | null
  readonly range: { readonly startMonth: number; readonly endMonth: number } | null
}

function parseCachePart(
  kind: 'slicer' | 'timeline',
  path: string,
  xml: string,
): ParsedCachePart | null {
  const root = kind === 'slicer' ? 'slicerCacheDefinition' : 'timelineCacheDefinition'
  const element = new RegExp(`<${root}(?:\\s[^>]*)?(?:/>|[\\s\\S]*?</${root}>)`).exec(xml)?.[0]
  if (element === undefined) return null
  const name = attr(element, 'name')
  const sourceName = attr(element, 'sourceName')
  if (name === undefined || sourceName === undefined) return null

  const tabular = /<tabular\b/.test(element)
  const pivotCacheIdRaw =
    attr(/<tabular\b[^>]*>/.exec(element)?.[0] ?? '', 'pivotCacheId') ??
    attr(element, 'pivotCacheId')
  const pivotCacheId = pivotCacheIdRaw === undefined ? null : Number(pivotCacheIdRaw)
  const tabIdRaw = attr(/<pivotTable\b[^>]*\/?>/.exec(element)?.[0] ?? '', 'tabId')
  const tabId = tabIdRaw === undefined ? null : Number(tabIdRaw)

  let selectedItems: readonly number[] | null = null
  const itemsBlock = /<items\b[^>]*>([\s\S]*?)<\/items>/.exec(element)?.[1]
  if (itemsBlock !== undefined) {
    const selected: number[] = []
    for (const item of itemsBlock.matchAll(/<i\b[^>]*\/>/g)) {
      const x = attr(item[0], 'x')
      if (x === undefined) continue
      const index = Number(x)
      if (!Number.isInteger(index) || index < 0) continue
      if (attr(item[0], 's') === '1') selected.push(index)
    }
    selectedItems = selected
  }

  let range: { readonly startMonth: number; readonly endMonth: number } | null = null
  const selection = /<selection\b[^>]*\/?>/.exec(element)?.[0]
  if (selection !== undefined) {
    const start = attr(selection, 'startDate')
    const end = attr(selection, 'endDate')
    const startMonth = start === undefined ? null : monthKeyOfIso(start)
    const endMonth = end === undefined ? null : monthKeyOfIso(end)
    if (startMonth !== null && endMonth !== null) {
      range = { startMonth, endMonth: Math.max(startMonth, endMonth) }
    }
  }
  return {
    path,
    name: decodeAttribute(name),
    sourceName: decodeAttribute(sourceName),
    pivotCacheId: pivotCacheId !== null && Number.isInteger(pivotCacheId) ? pivotCacheId : null,
    tabId: tabId !== null && Number.isInteger(tabId) ? tabId : null,
    tabular,
    selectedItems,
    range,
  }
}

/// Reads every slicer/timeline control the package carries, with enough
/// identity (part paths, names) for the session to update-in-place on save.
export async function parseSlicerState(reader: SlicerPartReader): Promise<SlicerPartState[]> {
  const paths = await reader.paths()
  const slicerParts = paths.filter((path) => /^xl\/slicers\/[^/]+\.xml$/.test(path)).sort()
  const timelineParts = paths.filter((path) => /^xl\/timelines\/[^/]+\.xml$/.test(path)).sort()
  if (slicerParts.length === 0 && timelineParts.length === 0) return []

  const workbookXml = await reader.readText('xl/workbook.xml')
  const workbookRelsXml = (await hasEntry(reader, 'xl/_rels/workbook.xml.rels'))
    ? await reader.readText('xl/_rels/workbook.xml.rels')
    : null
  const workbookRels = workbookRelsXml === null ? [] : parseRelationships(workbookRelsXml)
  const relTarget = (relId: string): string | null => {
    const entry = workbookRels.find((rel) => rel.id === relId)
    if (!entry) return null
    return resolveRelTarget('xl/workbook.xml', entry.target)
  }

  // pivotCacheId → pivotCacheDefinition part path (workbook pivotCaches).
  const cachePathByCacheId = new Map<number, string>()
  for (const pivotCache of workbookXml.matchAll(/<pivotCache\b[^>]*>/g)) {
    const cacheId = attr(pivotCache[0], 'cacheId')
    const relId = attr(pivotCache[0], 'r:id')
    if (cacheId === undefined || relId === undefined) continue
    const parsed = Number(cacheId)
    const target = relTarget(relId)
    if (target === null || !Number.isInteger(parsed)) continue
    cachePathByCacheId.set(parsed, target)
  }

  // Worksheet part path → sidecar sheet id (sheet-<sheetId attr>).
  const sheetIdByWorksheet = new Map<string, string>()
  for (const sheet of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const sheetId = attr(sheet[0], 'sheetId')
    const relId = attr(sheet[0], 'r:id')
    if (sheetId === undefined || relId === undefined) continue
    const target = relTarget(relId)
    if (target !== null) sheetIdByWorksheet.set(target, `sheet-${decodeAttribute(sheetId)}`)
  }

  // Cache parts by @name.
  const cachesByName = new Map<string, ParsedCachePart>()
  for (const path of paths) {
    const kind = /^xl\/slicerCaches\/[^/]+\.xml$/.test(path)
      ? 'slicer'
      : /^xl\/timelineCaches\/[^/]+\.xml$/.test(path)
        ? 'timeline'
        : null
    if (kind === null) continue
    const parsed = parseCachePart(kind, path, await reader.readText(path))
    if (parsed) cachesByName.set(parsed.name, parsed)
  }

  // Worksheet rels: which sheet hosts which control part.
  const hostByPart = new Map<string, string>()
  for (const path of paths) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(path)) continue
    const relsPath = relsPathFor(path)
    if (!(await hasEntry(reader, relsPath))) continue
    for (const entry of parseRelationships(await reader.readText(relsPath))) {
      if (entry.type !== SLICER_REL_TYPE && entry.type !== TIMELINE_REL_TYPE) continue
      hostByPart.set(resolveRelTarget(path, entry.target), path)
    }
  }

  const states: SlicerPartState[] = []
  const readControls = async (parts: readonly string[], kind: 'slicer' | 'timeline') => {
    for (const partPath of parts) {
      const xml = await reader.readText(partPath)
      for (const match of xml.matchAll(new RegExp(`<${kind}(?:\\s[^>]*)?/>`, 'g'))) {
        const control = match[0]
        const name = attr(control, 'name')
        const cacheName = attr(control, 'cache')
        if (name === undefined || cacheName === undefined) continue
        const cache = cachesByName.get(decodeAttribute(cacheName))
        const host = hostByPart.get(partPath)
        if (cache === undefined || host === undefined) continue
        if (kind === 'slicer' && !cache.tabular) continue // OLAP caches: skip
        // Binding: prefer tabular/root pivotCacheId; fall back to our tabId
        // rule over the sorted pivot table parts.
        let pivotCacheDefinitionPath: string | null = null
        if (cache.pivotCacheId !== null) {
          pivotCacheDefinitionPath = cachePathByCacheId.get(cache.pivotCacheId) ?? null
        }
        if (pivotCacheDefinitionPath === null && cache.tabId !== null) {
          const pivotParts = paths
            .filter((path) => /^xl\/pivotTables\/[^/]+\.xml$/.test(path))
            .sort()
          const pivotPath = pivotParts[cache.tabId - 1]
          if (pivotPath !== undefined) {
            const pivotCacheId = attr(await reader.readText(pivotPath), 'cacheId')
            const parsed = pivotCacheId === undefined ? Number.NaN : Number(pivotCacheId)
            if (Number.isInteger(parsed)) {
              pivotCacheDefinitionPath = cachePathByCacheId.get(parsed) ?? null
            }
          }
        }
        states.push({
          kind,
          sheetId: sheetIdByWorksheet.get(host) ?? '',
          worksheetPath: host,
          partPath,
          cachePath: cache.path,
          name: decodeAttribute(name),
          cacheName: cache.name,
          sourceName: cache.sourceName,
          caption: decodeAttribute(attr(control, 'caption') ?? name),
          pivotCacheDefinitionPath,
          selectedItems: kind === 'slicer' ? cache.selectedItems : null,
          range: kind === 'timeline' ? cache.range : null,
        })
      }
    }
  }
  await readControls(slicerParts, 'slicer')
  await readControls(timelineParts, 'timeline')
  return states
}

async function hasEntry(reader: SlicerPartReader, path: string): Promise<boolean> {
  return (await reader.paths()).includes(path)
}
