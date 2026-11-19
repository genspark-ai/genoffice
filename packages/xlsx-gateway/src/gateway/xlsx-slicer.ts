/// Slicer OOXML persistence (MS-XLSX x14 slicers). Timelines (x15) are not
/// persisted by this writer: their parts ride a save untouched, and the
/// session timeline feature stays session-only (to land separately).
///
/// Write side: session slicer panels are materialized as
///   xl/slicerCaches/slicerCacheN.xml   (pivot binding + member selection)
///   xl/slicers/slicerN.xml             (the controls: name/caption/geometry;
///     one part per worksheet, several controls can share it)
/// wired up by
///   worksheet rels (…/2007/relationships/slicer) + worksheet extLst
///     x14:slicerList ext entries,
///   workbook rels (…/slicerCache) + workbook extLst x14:slicerCaches ext
///     entries,
///   the bound pivotCacheDefinition's own x14 ext (pivotCacheId, its own
///     slicer-space id — not the workbook cacheId) — tabular@pivotCacheId
///     must equal it or Excel drops the whole slicerCache part as
///     'Removed Part: slicerCache',
///   one definedName per slicerCache named exactly like the cache
///     (slicerCacheDefinition@name, "#N/A", no hidden attribute — a different
///     name or hidden="1" makes Excel repair the slicer cache),
///   [Content_Types].xml overrides per part.
///
/// Read side mirrors the writer (the reader keys on container element names
/// and relationship types). The x14 ext URIs below are the MS-XLSX 2.2.2
/// values — Excel writes exactly these, and addExtRef/removeExtRef key on the
/// worksheet one, so a wrong URI both hides our ext from Excel and makes
/// Excel-authored parts invisible here. A test pins the literal strings: a
/// writer-mirrors-reader round trip cannot catch a wrong constant.
///
/// Pivot binding: slicerCacheDefinition carries tabular@pivotCacheId (the
/// bound pivot cache definition's x14 pivotCacheId — Excel keeps those in an
/// id space separate from workbook pivotCache@cacheId) and a
/// pivotTables/pivotTable entry whose tabId is the hosting worksheet's sheetId
/// and whose name is the pivot table's own name (MS-XLSX) — Excel does not
/// bind the slicer to its pivot without them. The reader resolves the id
/// through each definition's x14 ext and only falls back to workbook cacheId
/// for packages whose definitions carry no ext at all.

import {
  allocatePartPath,
  appendRelationship,
  ensureSheetDrawing,
  registerContentTypeOverride,
  relsPathFor,
  relativeTarget,
  resolveRelTarget,
  type MutablePackage,
} from './xlsx-drawing-add'
import { insertRootExt } from './xlsx-root-tail'
import { parseRelationships, removePartOverride, removeRelationshipById } from './xlsx-sheets'

export class SlicerAddError extends Error {}

// ─── Constants ─────────────────────────────────────────────────────────────

const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

export const SLICER_REL_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
export const SLICER_CACHE_REL_TYPE =
  'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
/// ECMA-376: the worksheet → pivotTable part relationship, used to find which
/// sheet hosts a pivot (the slicerCache binding's tabId is that sheet's sheetId)
const PIVOT_TABLE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'

const SLICER_CONTENT_TYPE = 'application/vnd.ms-excel.slicer+xml'
const SLICER_CACHE_CONTENT_TYPE = 'application/vnd.ms-excel.slicerCache+xml'

/// Ext URIs per MS-XLSX 2.2.2. The worksheet URI must be this exact value:
/// a wrong one made Excel treat the worksheet ext as foreign, never
/// instantiate the slicer, and discard the slicerCache + drawing on save.
/// It is also the key addExtRef/removeExtRef use to find Excel's own
/// x14:slicerList ext.
const WORKSHEET_SLICER_LIST_URI = '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}'
const WORKBOOK_SLICER_CACHES_URI = '{BBE1A952-AA13-448e-AADC-164F8A28A991}'
/// The x14 pivotCacheDefinition ext that makes a pivot cache slicer-bound
/// (MS-XLSX): its @pivotCacheId is what tabular@pivotCacheId must equal. The
/// id lives in its own space, unrelated to the workbook pivotCache@cacheId.
const PIVOT_CACHE_SLICER_EXT_URI = '{725AE2AE-9491-48be-B2B4-4EB974FC3084}'

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

export interface SlicerRemovalWrite {
  readonly partPath: string
  readonly cachePath: string
}

export interface SlicerStateWrite {
  readonly slicers: readonly SlicerControlWrite[]
  readonly removals: readonly SlicerRemovalWrite[]
}

// ─── Part XML builders ─────────────────────────────────────────────────────

/// SlicerCache pivot binding per MS-XLSX: pivotTables/pivotTable
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

/// The slicer-space id a pivotCacheDefinition's x14 ext carries, or null when
/// the definition has no slicer ext. This is NOT the workbook
/// pivotCache@cacheId: Excel keeps the two in separate id spaces (a genuine
/// file pairs cacheId="110" with x14 pivotCacheId="2124480673"), and only this
/// id is what slicerCache tabular@pivotCacheId refers to.
function x14PivotCacheIdOf(definitionXml: string): number | null {
  const ext = new RegExp(
    `<ext\\b[^>]*\\buri="${PIVOT_CACHE_SLICER_EXT_URI}"[^>]*>[\\s\\S]*?</ext>`,
  ).exec(definitionXml)?.[0]
  if (ext === undefined) return null
  const raw = attr(/<x14:pivotCacheDefinition\b[^>]*>/.exec(ext)?.[0] ?? '', 'pivotCacheId')
  const parsed = raw === undefined ? Number.NaN : Number(raw)
  return Number.isInteger(parsed) && parsed > 0 ? parsed : null
}

/// Every pivotCacheDefinition part the workbook's pivotCaches list resolves to.
async function pivotCacheDefinitionPaths(
  pkg: MutablePackage,
  workbookXml: string,
): Promise<string[]> {
  const paths: string[] = []
  for (const pivotCache of workbookXml.matchAll(/<pivotCache\b[^>]*>/g)) {
    const cacheId = attr(pivotCache[0], 'cacheId')
    if (cacheId === undefined) continue
    const parsed = Number(cacheId)
    if (!Number.isInteger(parsed)) continue
    const path = await pivotCacheDefinitionPathOf(pkg, workbookXml, parsed)
    if (path !== null && !paths.includes(path)) paths.push(path)
  }
  return paths
}

/// A fresh x14 pivotCacheId: one past the highest id any pivot cache
/// definition in this workbook already carries (1 when none does). The ids are
/// their own space, MUST be non-zero when a slicer cache uses the cache, and
/// must not collide.
async function nextX14PivotCacheId(pkg: MutablePackage, workbookXml: string): Promise<number> {
  let max = 0
  for (const path of await pivotCacheDefinitionPaths(pkg, workbookXml)) {
    if (!(await pkg.has(path))) continue
    const id = x14PivotCacheIdOf(await pkg.readText(path))
    if (id !== null && id > max) max = id
  }
  return max + 1
}

interface RawAttribute {
  readonly name: string
  /// The attribute exactly as written (escaped), ready to concatenate.
  readonly raw: string
}

/// Attributes of a start-tag string with their original entity escaping, so
/// Excel-authored attributes survive a part rewrite verbatim.
function rawAttributes(startTag: string): RawAttribute[] {
  const attributes: RawAttribute[] = []
  for (const match of startTag.matchAll(/([A-Za-z_][\w.:-]*)\s*=\s*("[^"]*"|'[^']*')/g)) {
    attributes.push({ name: match[1]!, raw: `${match[1]}=${match[2]}` })
  }
  return attributes
}

/// One control inside a slicers part. Excel keeps every slicer of a worksheet
/// in one xl/slicers/slicerN.xml part, so a rewrite of one control must carry
/// the others through; the identity attributes (name/cache/caption) belong to
/// the session, every other attribute and the child content stay verbatim.
interface CarriedSlicerControl {
  readonly name: string
  readonly cacheName: string
  readonly caption: string
  readonly attributes: string
  readonly inner: string
  /// The session owns this control: a missing rowHeight gets Excel's default.
  /// Preserved controls keep exactly the attributes the file carried.
  readonly requireRowHeight?: boolean
}

interface CarriedSlicerPart {
  readonly root: string
  readonly controls: readonly CarriedSlicerControl[]
}

/// Every <slicer> in a part, self-closing or with children (an Excel slicer
/// with an extLst is the paired form). Names/cache/caption are decoded;
/// attributes keep the raw start-tag escapes of everything else.
function parseSlicerControls(xml: string): CarriedSlicerControl[] {
  const controls: CarriedSlicerControl[] = []
  const tag = /<slicer\b[^>]*?(\/?)>/g
  let match: RegExpExecArray | null
  while ((match = tag.exec(xml)) !== null) {
    const start = match[0]
    const startAt = match.index
    let inner = ''
    if (match[1] !== '/') {
      const close = xml.indexOf('</slicer>', startAt + start.length)
      if (close === -1) break
      inner = xml.slice(startAt + start.length, close)
      tag.lastIndex = close + '</slicer>'.length
    }
    const name = attr(start, 'name')
    const cache = attr(start, 'cache')
    if (name === undefined || cache === undefined) continue
    let attributes = ''
    for (const attribute of rawAttributes(start)) {
      if (attribute.name === 'name' || attribute.name === 'cache' || attribute.name === 'caption') {
        continue
      }
      attributes += ' ' + attribute.raw
    }
    controls.push({
      name: decodeAttribute(name),
      cacheName: decodeAttribute(cache),
      caption: decodeAttribute(attr(start, 'caption') ?? name),
      attributes,
      inner,
    })
  }
  return controls
}

function carryOverSlicerPart(existingXml: string | null): CarriedSlicerPart {
  const xml = existingXml ?? ''
  const rootTag = /<slicers\b[^>]*>/.exec(xml)?.[0] ?? ''
  let root = ''
  for (const attribute of rawAttributes(rootTag)) {
    // We state the slicer namespace ourselves; every other declaration (and
    // mc:Ignorable) belongs to the file.
    if (attribute.name === 'xmlns') continue
    root += ' ' + attribute.raw
  }
  return { root, controls: parseSlicerControls(xml) }
}

/// One <slicer> of the rebuilt part. A session control that carried no
/// Excel-authored rowHeight gets Excel's own default (241300 EMU, about a
/// quarter inch); preserved controls keep exactly their file attributes.
function slicerControlXml(control: CarriedSlicerControl): string {
  const rowHeight =
    control.requireRowHeight === true && !control.attributes.includes('rowHeight=')
      ? ' rowHeight="241300"'
      : ''
  const element =
    '<slicer name="' +
    escapeAttribute(control.name) +
    '" cache="' +
    escapeAttribute(control.cacheName) +
    '" caption="' +
    escapeAttribute(control.caption) +
    '"' +
    control.attributes +
    rowHeight
  return control.inner === '' ? element + '/>' : element + '>' + control.inner + '</slicer>'
}

function buildSlicerPartXml(carried: CarriedSlicerPart): string {
  const controls = carried.controls.map(slicerControlXml).join('')
  return XML_DECL + '<slicers xmlns="' + X14_NS + '"' + carried.root + '>' + controls + '</slicers>'
}

/// Replace the control named by the session, appending it when the shared part
/// does not carry it yet; every other control rides through untouched.
function upsertSlicerControl(
  carried: CarriedSlicerPart,
  updated: CarriedSlicerControl,
): CarriedSlicerPart {
  const found = carried.controls.some((control) => control.name === updated.name)
  return {
    root: carried.root,
    controls: found
      ? carried.controls.map((control) => (control.name === updated.name ? updated : control))
      : [...carried.controls, updated],
  }
}

/// Strip one control by name. The part and its worksheet rel/ext only go away
/// when no control is left in it.
function withoutSlicerControl(carried: CarriedSlicerPart, name: string): CarriedSlicerPart {
  return {
    root: carried.root,
    controls: carried.controls.filter((control) => control.name !== name),
  }
}

function buildSlicerCacheXml(
  cacheName: string,
  sourceName: string,
  binding: { readonly tabId: number; readonly name: string },
  x14PivotCacheId: number,
  items: readonly SlicerItemWrite[],
  existingXml: string | null,
): string {
  const existing = existingXml ?? ''
  // Attributes Excel owns (xr10:uid, mc:Ignorable, its namespace
  // declarations) and the tabular's ordering flags (sortOrder, customListSort,
  // showMissing, crossFilter) survive a rewrite; the session owns only the
  // cache identity, the binding, and the item selection.
  let root = ''
  const rootTag = /<slicerCacheDefinition\b[^>]*>/.exec(existing)?.[0] ?? ''
  for (const attribute of rawAttributes(rootTag)) {
    if (['name', 'sourceName', 'xmlns', 'xmlns:r'].includes(attribute.name)) continue
    root += ` ${attribute.raw}`
  }
  let tabular = ''
  let hasCrossFilter = false
  const tabularTag = /<tabular\b[^>]*>/.exec(existing)?.[0] ?? ''
  for (const attribute of rawAttributes(tabularTag)) {
    if (attribute.name === 'pivotCacheId') continue
    if (attribute.name === 'crossFilter') hasCrossFilter = true
    tabular += ` ${attribute.raw}`
  }
  if (!hasCrossFilter) tabular += ' crossFilter="showItemsWithDataAtTop"'
  const extLstStart = existing.indexOf('<extLst')
  const extLstClose = existing.lastIndexOf('</extLst>')
  const extLst =
    extLstStart !== -1 && extLstClose > extLstStart
      ? existing.slice(extLstStart, extLstClose + '</extLst>'.length)
      : ''
  // An Excel slicer can be connected to several pivot tables (Report
  // Connections); its pivotTables block lists them all. Keep the file's block
  // on rewrite — the session model binds one pivot, which is only the default
  // for a fresh part.
  const pivotTablesBlock =
    /<pivotTables\b[^>]*?(?:\/>|>[\s\S]*?<\/pivotTables>)/.exec(existing)?.[0] ??
    `<pivotTables><pivotTable tabId="${binding.tabId}" name="${escapeAttribute(binding.name)}"/></pivotTables>`
  const itemsXml = items
    .map((item) => `<i x="${item.x}"${item.selected ? ' s="1"' : ''}/>`)
    .join('')
  return (
    `${XML_DECL}<slicerCacheDefinition xmlns="${X14_NS}" xmlns:r="${REL_NS}"${root}` +
    ` name="${escapeAttribute(cacheName)}" sourceName="${escapeAttribute(sourceName)}">` +
    pivotTablesBlock +
    '<data>' +
    // tabular@pivotCacheId is the pivot cache definition's x14 id below — the
    // two must match or Excel repairs/drops the file ('Removed Part: slicerCache')
    `<tabular pivotCacheId="${x14PivotCacheId}"${tabular}>` +
    `<items count="${items.length}">${itemsXml}</items>` +
    `</tabular></data>${extLst}</slicerCacheDefinition>`
  )
}

/// The pivotCacheDefinition part a workbook cacheId resolves to (through
/// workbook pivotCaches and their rels), or null when it cannot be found.
async function pivotCacheDefinitionPathOf(
  pkg: MutablePackage,
  workbookXml: string,
  cacheId: number,
): Promise<string | null> {
  const relsXml = (await pkg.has('xl/_rels/workbook.xml.rels'))
    ? await pkg.readText('xl/_rels/workbook.xml.rels')
    : ''
  const rels = parseRelationships(relsXml)
  for (const pivotCache of workbookXml.matchAll(/<pivotCache\b[^>]*>/g)) {
    if (attr(pivotCache[0], 'cacheId') !== String(cacheId)) continue
    const relId = attr(pivotCache[0], 'r:id')
    const entry = relId === undefined ? undefined : rels.find((rel) => rel.id === relId)
    if (entry !== undefined) return resolveRelTarget('xl/workbook.xml', entry.target)
  }
  return null
}

/// The pivotCacheDefinition's x14 ext (MS-XLSX 2.6.33): the slicer-space
/// pivotCacheId, which MUST be non-zero when a slicer cache uses the cache
/// (tabular@pivotCacheId carries the same value). No slicerData: that flag is
/// only for OLAP slicer sources and MUST be "false" for our tabular
/// (worksheet-pivot) caches — writing slicerData="1" made Excel remove the
/// whole slicer cache part as a repair ("Removed Feature: Slicer Cache").
///
/// Returns the x14 id to put in the slicerCache's tabular. An existing ext
/// keeps its id (other slicer caches and Excel refer to it, and it is not the
/// workbook cacheId); only a definition without a slicer ext gets a fresh id.
async function ensurePivotCacheSlicerExt(
  pkg: MutablePackage,
  workbookXml: string,
  pivotPath: string,
  cacheId: number,
  touchedEntries: Set<string>,
): Promise<number> {
  const definitionPath = await pivotCacheDefinitionPathOf(pkg, workbookXml, cacheId)
  if (definitionPath === null) {
    throw new SlicerAddError(`No pivotCacheDefinition carries cacheId ${cacheId}.`)
  }

  const xml = await pkg.readText(definitionPath)
  const x14CacheId = x14PivotCacheIdOf(xml) ?? (await nextX14PivotCacheId(pkg, workbookXml))
  const extElement =
    `<ext uri="${PIVOT_CACHE_SLICER_EXT_URI}" xmlns:x14="${X14_NS}">` +
    `<x14:pivotCacheDefinition pivotCacheId="${x14CacheId}"/></ext>`
  let patched: string
  const existing = new RegExp(
    `<ext\\b[^>]*\\buri="${PIVOT_CACHE_SLICER_EXT_URI}"[^>]*>[\\s\\S]*?</ext>`,
  ).exec(xml)
  if (existing) {
    patched =
      xml.slice(0, existing.index) + extElement + xml.slice(existing.index + existing[0].length)
  } else if (xml.includes('</pivotCacheDefinition>')) {
    // extLst is the schema-last child; the definition's own one takes the
    // ext (cache fields carry nested lists), else the element gains one
    const inserted = insertRootExt(xml, 'pivotCacheDefinition', extElement)
    if (inserted === null) {
      throw new SlicerAddError(`Pivot cache definition ${definitionPath} is malformed.`)
    }
    patched = inserted
  } else {
    // a self-closing definition opens up to carry the extLst
    const selfClosing = /<pivotCacheDefinition\b[^>]*?\/>/.exec(xml)
    if (selfClosing === null) {
      throw new SlicerAddError(`Pivot cache definition ${definitionPath} is malformed.`)
    }
    patched =
      xml.slice(0, selfClosing.index) +
      selfClosing[0].slice(0, -2) +
      `><extLst>${extElement}</extLst></pivotCacheDefinition>` +
      xml.slice(selfClosing.index + selfClosing[0].length)
  }
  pkg.write(definitionPath, patched)
  touchedEntries.add(definitionPath)
  return x14CacheId
}

/// Drops the x14 slicer ext from the pivot cache definition that carries
/// `x14CacheId`, once no remaining slicerCache part binds that id anymore.
async function removePivotCacheSlicerExt(
  pkg: MutablePackage,
  workbookXml: string,
  x14CacheId: number,
  touchedEntries: Set<string>,
): Promise<void> {
  for (const path of await pkg.paths()) {
    if (!/^xl\/slicerCaches\/[^/]+\.xml$/.test(path)) continue
    const bound = Number(
      attr(/<tabular\b[^>]*>/.exec(await pkg.readText(path))?.[0] ?? '', 'pivotCacheId') ?? NaN,
    )
    if (bound === x14CacheId) return // another slicer still binds this cache
  }
  for (const definitionPath of await pivotCacheDefinitionPaths(pkg, workbookXml)) {
    if (!(await pkg.has(definitionPath))) continue
    const xml = await pkg.readText(definitionPath)
    if (x14PivotCacheIdOf(xml) !== x14CacheId) continue
    const ext = new RegExp(
      `<ext\\b[^>]*\\buri="${PIVOT_CACHE_SLICER_EXT_URI}"[^>]*>[\\s\\S]*?</ext>\\s*`,
    )
    let patched = xml.replace(ext, '')
    const open = patched.lastIndexOf('<extLst>')
    if (open !== -1) {
      const close = patched.indexOf('</extLst>', open)
      if (close !== -1 && patched.slice(open + '<extLst>'.length, close).trim() === '') {
        patched = patched.slice(0, open) + patched.slice(close + '</extLst>'.length)
      }
    }
    if (patched !== xml) {
      pkg.write(definitionPath, patched)
      touchedEntries.add(definitionPath)
    }
    return
  }
}

// ─── On-sheet control (drawing graphicFrame) ────────────────────────────────

/// The control's visible face lives in the worksheet's drawing part as a
/// graphicFrame (sle:slicer). The slicer part alone is plumbing: without the
/// frame Excel opens the file with no control on the sheet. An anchor that
/// already names this control is left untouched — that is where Excel-authored
/// geometry survives a rewrite.
async function ensureControlFrame(
  pkg: MutablePackage,
  worksheetPath: string,
  name: string,
  touchedEntries: Set<string>,
): Promise<void> {
  const drawing = await ensureSheetDrawing(pkg, worksheetPath, touchedEntries)
  const xml = await pkg.readText(drawing.path)
  if (new RegExp(`<sle:slicer\\b[^>]*\\bname="${escapeRegExp(name)}"`).test(xml)) return
  // No session geometry: new controls land in a column of slots below the
  // used range's first rows, offset per existing anchor so they do not stack.
  const anchors = (xml.match(/<xdr:twoCellAnchor\b/g) ?? []).length
  const fromCol = 8 + (anchors % 4) * 4
  const fromRow = 2 + Math.floor(anchors / 4) * 14
  const marker = (col: number, row: number) =>
    `<xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff>` +
    `<xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff>`
  // cNvPr ids are unique per drawing part: max-plus-one, not anchor-count
  // (an Excel-authored part can carry arbitrary ids)
  let maxId = 1
  for (const id of xml.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)) {
    const parsed = Number(id[1])
    if (Number.isInteger(parsed) && parsed > maxId) maxId = parsed
  }
  const control =
    `<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer"` +
    ` name="${escapeAttribute(name)}"/>`
  const frame =
    '<xdr:twoCellAnchor editAs="oneCell">' +
    `<xdr:from>${marker(fromCol, fromRow)}</xdr:from>` +
    `<xdr:to>${marker(fromCol + 3, fromRow + 12)}</xdr:to>` +
    '<mc:AlternateContent xmlns:mc="' +
    MC_NS +
    '">' +
    `<mc:Choice xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" Requires="sle">` +
    '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>' +
    `<xdr:cNvPr id="${maxId + 1}" name="${escapeAttribute(name)}"/>` +
    '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer">' +
    control +
    '</a:graphicData></a:graphic></xdr:graphicFrame>' +
    '</mc:Choice>' +
    // a reader without the slicer extension draws the fallback shape instead
    // of repairing the file (what Excel itself writes for down-level clients)
    '<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr>' +
    `<xdr:cNvPr id="${maxId + 2}" name="${escapeAttribute(name)} Fallback"/>` +
    '<xdr:cNvSpPr/><xdr:nvPr/></xdr:nvSpPr>' +
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>' +
    '<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></xdr:txBody>' +
    '</xdr:sp></mc:Fallback>' +
    '</mc:AlternateContent>' +
    // clientData belongs to the anchor, not to the graphicFrame: inside the
    // frame Excel repairs the drawing ('Repaired Records: Drawing').
    '<xdr:clientData/>' +
    '</xdr:twoCellAnchor>'
  pkg.write(drawing.path, xml.replace('</xdr:wsDr>', `${frame}</xdr:wsDr>`))
  touchedEntries.add(drawing.path)
}

/// Drops the twoCellAnchor whose sle:slicer names `name` — removal without it
/// strands an orphan control in the drawing (Excel repairs it away).
async function removeControlFrame(
  pkg: MutablePackage,
  worksheetPath: string,
  name: string,
  touchedEntries: Set<string>,
): Promise<void> {
  const relsPath = relsPathFor(worksheetPath)
  if (!(await pkg.has(relsPath))) return
  const relsXml = await pkg.readText(relsPath)
  let drawingPath: string | null = null
  for (const entry of parseRelationships(relsXml)) {
    if (
      entry.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
    ) {
      drawingPath = resolveRelTarget(worksheetPath, entry.target)
      break
    }
  }
  if (drawingPath === null || !(await pkg.has(drawingPath))) return
  const xml = await pkg.readText(drawingPath)
  const pattern = new RegExp(
    `<xdr:twoCellAnchor\\b[^>]*>(?:(?!</xdr:twoCellAnchor>)[\\s\\S])*?` +
      `<sle:slicer\\b[^>]*\\bname="${escapeRegExp(name)}"[^>]*/>` +
      `[\\s\\S]*?</xdr:twoCellAnchor>\\s*`,
  )
  const stripped = xml.replace(pattern, '')
  if (stripped !== xml) {
    pkg.write(drawingPath, stripped)
    touchedEntries.add(drawingPath)
  }
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
const WORKBOOK_SLICER_WIRING: ExtWiring = {
  uri: WORKBOOK_SLICER_CACHES_URI,
  prefix: 'x14',
  namespace: X14_NS,
  container: 'slicerCaches',
  item: 'slicerCache',
}

function extOpenTag(xml: string, wiring: ExtWiring): { start: number; end: number } | null {
  const match = new RegExp(`<ext\\b[^>]*\\buri="${escapeRegExp(wiring.uri)}"[^>]*>`, 'i').exec(xml)
  if (!match) return null
  return { start: match.index, end: match.index + match[0].length }
}

/// Adds `<{prefix}:{item} r:id="…"/>` inside the wiring's ext (creating the
/// ext, its container, and the host extLst as needed). Only the host's own
/// extLst qualifies: cells and linked conditional-formatting rules carry
/// nested ones.
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
  const patched = insertRootExt(
    xml,
    hostClose === '</worksheet>' ? 'worksheet' : 'workbook',
    element,
  )
  if (patched === null) throw new SlicerAddError('The host element is malformed.')
  return patched
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

/// A worksheet that already owns a slicer part (Excel keeps every control of
/// a sheet in one part) takes a new control into that part instead of gaining
/// a second part Excel would never author.
async function slicerPartPathForWorksheet(
  pkg: MutablePackage,
  worksheetPath: string,
): Promise<string | null> {
  const relsPath = relsPathFor(worksheetPath)
  if (!(await pkg.has(relsPath))) return null
  for (const entry of parseRelationships(await pkg.readText(relsPath))) {
    if (entry.type !== SLICER_REL_TYPE) continue
    const target = resolveRelTarget(worksheetPath, entry.target)
    if (await pkg.has(target)) return target
  }
  return null
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

async function collectControlNames(pkg: MutablePackage, directory: string): Promise<Set<string>> {
  const names = new Set<string>()
  for (const path of await pkg.paths()) {
    if (!path.startsWith(directory) || !path.endsWith('.xml')) continue
    for (const control of parseSlicerControls(await pkg.readText(path))) names.add(control.name)
  }
  return names
}

/// True when a control still in the package (after the target was pruned)
/// references cacheName: its slicerCache part and workbook wiring must stay.
async function otherControlUsesCache(pkg: MutablePackage, cacheName: string): Promise<boolean> {
  for (const path of await pkg.paths()) {
    if (!/^xl\/slicers\/[^/]+\.xml$/.test(path)) continue
    if (!(await pkg.has(path))) continue
    for (const control of parseSlicerControls(await pkg.readText(path))) {
      if (control.cacheName === cacheName) return true
    }
  }
  return false
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

/// Workbook members that follow <definedNames> in the CT_Workbook sequence.
/// Creating the container has to land before the first of them: an existing
/// <extLst> (we add one for the slicer caches) must stay the last child.
const AFTER_DEFINED_NAMES = [
  '<calcPr',
  '<oleSize',
  '<customWorkbookViews',
  '<pivotCaches',
  '<smartTagPr',
  '<smartTagTypes',
  '<webPublishing',
  '<fileRecoveryPr',
  '<webPublishObjects',
  '<extLst',
] as const

/// Excel requires exactly one definedName per slicerCache, named exactly like
/// the cache (`slicerCacheDefinition@name`) with `#N/A` as the value: the name
/// is a marker, not a reference. Keying it on the control name instead leaves an
/// invalid/absent name (spaces, a fresh caption) that makes Excel repair the
/// file, and a restore-then-save must refresh the existing entry rather than add
/// a second one.
function addSlicerDefinedName(workbookXml: string, cacheName: string): string {
  const element = `<definedName name="${escapeAttribute(cacheName)}">#N/A</definedName>`
  const existing = new RegExp(
    `<definedName\\b[^>]*\\bname="${escapeRegExp(escapeAttribute(cacheName))}"[^>]*>[^<]*</definedName>`,
  )
  if (existing.test(workbookXml)) return workbookXml.replace(existing, element)
  // A self-closing <definedNames/> (openpyxl and others write it) is opened up
  // in place; appending a second container after </sheets> is a schema
  // violation Excel repairs.
  const selfClosing = /<definedNames\b[^>]*\/>/.exec(workbookXml)
  if (selfClosing) {
    return (
      workbookXml.slice(0, selfClosing.index) +
      `<definedNames>${element}</definedNames>` +
      workbookXml.slice(selfClosing.index + selfClosing[0].length)
    )
  }
  const close = workbookXml.lastIndexOf('</definedNames>')
  if (close !== -1) return workbookXml.slice(0, close) + element + workbookXml.slice(close)
  const sheetsClose = workbookXml.lastIndexOf('</sheets>')
  const workbookClose = workbookXml.lastIndexOf('</workbook>')
  if (sheetsClose === -1 || workbookClose === -1) {
    throw new SlicerAddError('The workbook part is malformed.')
  }
  let insertAt = workbookClose
  const afterSheets = sheetsClose + '</sheets>'.length
  for (const member of AFTER_DEFINED_NAMES) {
    const match = workbookXml.indexOf(member, afterSheets)
    if (match !== -1 && match < insertAt) insertAt = match
  }
  return (
    workbookXml.slice(0, insertAt) +
    `<definedNames>${element}</definedNames>` +
    workbookXml.slice(insertAt)
  )
}

function removeSlicerDefinedName(workbookXml: string, cacheName: string): string {
  const pattern = new RegExp(
    `<definedName\\b[^>]*\\bname="${escapeRegExp(escapeAttribute(cacheName))}"[^>]*>[^<]*</definedName>\\s*`,
  )
  let patched = workbookXml.replace(pattern, '')
  // an emptied <definedNames> container drops out
  const open = patched.lastIndexOf('<definedNames>')
  if (open !== -1) {
    const close = patched.indexOf('</definedNames>', open)
    if (close !== -1 && patched.slice(open + '<definedNames>'.length, close).trim() === '') {
      patched = patched.slice(0, open) + patched.slice(close + '</definedNames>'.length)
    }
  }
  return patched
}

async function writeControl(
  pkg: MutablePackage,
  control: SlicerControlWrite,
  workbookXml: string,
  touchedEntries: Set<string>,
  takenControlNames: Set<string>,
  takenCacheNames: Set<string>,
): Promise<string> {
  if (!(await pkg.has(control.worksheetPath))) {
    throw new SlicerAddError(`Slicer host worksheet ${control.worksheetPath} is missing.`)
  }
  const binding = await pivotBinding(pkg, workbookXml, control.pivotPath)
  const workbookCacheId = await pivotCacheIdOf(pkg, control.pivotPath)
  // The bound pivot cache definition gains (or keeps) its x14 slicer id — the
  // slicerCache's tabular@pivotCacheId must equal that id (MS-XLSX), and it is
  // not the workbook cacheId.
  const x14PivotCacheId = await ensurePivotCacheSlicerExt(
    pkg,
    workbookXml,
    control.pivotPath,
    workbookCacheId,
    touchedEntries,
  )

  // Identity: reuse the restored part paths / names, else allocate fresh ones
  // that do not collide with anything already in the package.
  const partPath =
    control.partPath ??
    (await slicerPartPathForWorksheet(pkg, control.worksheetPath)) ??
    (await allocatePartPath(pkg, 'xl/slicers/slicer', '.xml'))

  const cachePath =
    control.cachePath ?? (await allocatePartPath(pkg, 'xl/slicerCaches/slicerCache', '.xml'))
  const cacheName =
    control.cacheName ??
    uniqueName(`Slicer_${sanitizeNameBase(control.sourceName)}`, takenCacheNames)
  const name =
    control.name ??
    uniqueName(sanitizeNameBase(control.caption || control.sourceName), takenControlNames)

  // Rewrites keep what Excel authored on the parts themselves (geometry,
  // style, ordering flags, extLst); a fresh part starts from our defaults.
  const existingPart = (await pkg.has(partPath)) ? await pkg.readText(partPath) : null
  const existingCache = (await pkg.has(cachePath)) ? await pkg.readText(cachePath) : null
  const cacheXml = buildSlicerCacheXml(
    cacheName,
    control.sourceName,
    binding,
    x14PivotCacheId,
    control.items,
    existingCache,
  )
  if (existingCache !== null) pkg.write(cachePath, cacheXml)
  else pkg.add(cachePath, cacheXml)
  touchedEntries.add(cachePath)
  await registerContentTypeOverride(pkg, cachePath, SLICER_CACHE_CONTENT_TYPE, touchedEntries)

  // The workbook rel + ext entry point at the cache part, and the
  // cache-named definedName joins the workbook's name space.
  const relId = await appendOrReuseRelationship(
    pkg,
    'xl/_rels/workbook.xml.rels',
    SLICER_CACHE_REL_TYPE,
    'xl/workbook.xml',
    cachePath,
    touchedEntries,
  )
  let wiredWorkbook = addExtRef(workbookXml, WORKBOOK_SLICER_WIRING, relId, '</workbook>')
  wiredWorkbook = addSlicerDefinedName(wiredWorkbook, cacheName)

  // The control part: Excel keeps every slicer of a worksheet in one part, so
  // a rewrite replaces only the control keyed by its name (appending it when
  // the part does not carry it) and preserves the other controls verbatim.
  const carried = carryOverSlicerPart(existingPart)
  const previous = carried.controls.find((control) => control.name === name)
  const partXml = buildSlicerPartXml(
    upsertSlicerControl(carried, {
      name,
      cacheName,
      caption: control.caption,
      attributes: previous?.attributes ?? '',
      inner: previous?.inner ?? '',
      requireRowHeight: true,
    }),
  )

  if (existingPart === null) pkg.add(partPath, partXml)
  else pkg.write(partPath, partXml)
  touchedEntries.add(partPath)
  await registerContentTypeOverride(pkg, partPath, SLICER_CONTENT_TYPE, touchedEntries)

  await ensureWorksheetRef(
    pkg,
    control.worksheetPath,
    partPath,
    SLICER_REL_TYPE,
    WORKSHEET_SLICER_WIRING,
    touchedEntries,
  )
  await ensureControlFrame(pkg, control.worksheetPath, name, touchedEntries)
  return wiredWorkbook
}

/// Materializes every slicer control and cleans removals. Returns the workbook
/// XML with the cache ext entries merged in. Removals run first so a deleted
/// control's parts free their names for reuse.
export async function applySlicerState(
  pkg: MutablePackage,
  state: SlicerStateWrite,
  workbookXml: string,
  touchedEntries: Set<string>,
): Promise<string> {
  // A removal and a write may share the slicer part (several controls live in
  // one part); only the cache part is exclusive to one control.
  const writtenCachePaths = new Set(
    state.slicers.flatMap((control) =>
      control.cachePath === undefined ? [] : [control.cachePath],
    ),
  )
  for (const removal of state.removals) {
    if (writtenCachePaths.has(removal.cachePath)) {
      throw new SlicerAddError('A removed slicer cache cannot also be written in the same save.')
    }
    // Removal is keyed by the control's cache part: the shared slicers part may
    // carry other controls (Excel keeps all slicers of a sheet together), and
    // only the part's last control takes the part/rel/ext with it.
    const partXml = (await pkg.has(removal.partPath)) ? await pkg.readText(removal.partPath) : null
    const cacheXml = (await pkg.has(removal.cachePath))
      ? await pkg.readText(removal.cachePath)
      : null
    const cacheNameRaw = attr(
      /<slicerCacheDefinition\b[^>]*>/.exec(cacheXml ?? '')?.[0] ?? '',
      'name',
    )
    const cacheName = cacheNameRaw === undefined ? undefined : decodeAttribute(cacheNameRaw)
    const carried = carryOverSlicerPart(partXml)
    // Match on the cache name only: a cache part that is already gone must not
    // make removal strip an arbitrary control from a shared part.
    const target =
      cacheName === undefined
        ? undefined
        : carried.controls.find((control) => control.cacheName === cacheName)
    const name = target?.name ?? null
    const remaining = name === null ? carried : withoutSlicerControl(carried, name)
    const boundX14CacheId = Number(
      attr(/<tabular\b[^>]*>/.exec(cacheXml ?? '')?.[0] ?? '', 'pivotCacheId') ?? NaN,
    )
    for (const host of await worksheetsReferencing(pkg, removal.partPath, SLICER_REL_TYPE)) {
      if (name !== null) await removeControlFrame(pkg, host, name, touchedEntries)
      if (remaining.controls.length === 0) {
        await removeWorksheetRef(
          pkg,
          host,
          removal.partPath,
          SLICER_REL_TYPE,
          WORKSHEET_SLICER_WIRING,
          touchedEntries,
        )
      }
    }
    if (name !== null) {
      if (remaining.controls.length > 0) {
        pkg.write(removal.partPath, buildSlicerPartXml(remaining))
        touchedEntries.add(removal.partPath)
      } else {
        await removePartWithSatellites(pkg, removal.partPath, touchedEntries)
      }
    }
    // Another slicer can share this cache; its part, workbook wiring and the
    // one-per-cache definedName all stay.
    const cacheStillUsed = cacheName !== undefined && (await otherControlUsesCache(pkg, cacheName))
    if (!cacheStillUsed) {
      if (cacheName !== undefined) {
        workbookXml = removeSlicerDefinedName(workbookXml, cacheName)
      }
      const relsPath = 'xl/_rels/workbook.xml.rels'
      if (await pkg.has(relsPath)) {
        const relsXml = await pkg.readText(relsPath)
        const existing = findRelationshipTo(
          relsXml,
          SLICER_CACHE_REL_TYPE,
          'xl/workbook.xml',
          removal.cachePath,
        )
        if (existing) {
          pkg.write(relsPath, removeRelationshipById(relsXml, existing.id))
          touchedEntries.add(relsPath)
          workbookXml = removeExtRef(workbookXml, WORKBOOK_SLICER_WIRING, existing.id)
        }
      }
      await removePartWithSatellites(pkg, removal.cachePath, touchedEntries)
      if (Number.isInteger(boundX14CacheId)) {
        await removePivotCacheSlicerExt(pkg, workbookXml, boundX14CacheId, touchedEntries)
      }
    }
  }

  if (state.slicers.length > 0) {
    const takenNames = await collectControlNames(pkg, 'xl/slicers/')
    const takenCaches = await collectCacheNames(pkg, 'xl/slicerCaches/')
    for (const slicer of state.slicers) {
      workbookXml = await writeControl(
        pkg,
        slicer,
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
  /// Sidecar sheet id (`sheet-<workbook sheetId attribute>`).
  readonly sheetId: string
  readonly worksheetPath: string
  readonly partPath: string
  readonly cachePath: string
  /// slicer@name.
  readonly name: string
  /// slicerCacheDefinition@name.
  readonly cacheName: string
  /// slicerCacheDefinition@sourceName (the bound cache field's name).
  readonly sourceName: string
  readonly caption: string
  /// pivotCacheDefinition part path resolved from the cache binding; null
  /// when unresolvable.
  readonly pivotCacheDefinitionPath: string | null
  /// sharedItems indices of the selected members; null = the cache carried
  /// no item list (treat as unfiltered).
  readonly selectedItems: readonly number[] | null
}

export interface SlicerPartReader {
  paths(): Promise<readonly string[]>
  readText(path: string): Promise<string>
}

interface ParsedCachePart {
  readonly path: string
  readonly name: string
  readonly sourceName: string
  /// tabular@pivotCacheId — the bound pivot cache definition's x14 slicer id
  /// (its own id space, not the workbook cacheId); null when the part
  /// carries none.
  readonly pivotCacheId: number | null
  /// true when the data section is tabular (OLAP caches cannot be
  /// represented by the session model).
  readonly tabular: boolean
  readonly selectedItems: readonly number[] | null
}

function parseCachePart(path: string, xml: string): ParsedCachePart | null {
  const root = 'slicerCacheDefinition'
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

  return {
    path,
    name: decodeAttribute(name),
    sourceName: decodeAttribute(sourceName),
    pivotCacheId: pivotCacheId !== null && Number.isInteger(pivotCacheId) ? pivotCacheId : null,
    tabular,
    selectedItems,
  }
}

/// Reads every slicer control the package carries, with enough identity (part
/// paths, names) for the session to update-in-place on save. Timeline parts
/// are not read: the session has no timeline persistence yet, and an unread
/// foreign part rides a save untouched.
export async function parseSlicerState(reader: SlicerPartReader): Promise<SlicerPartState[]> {
  const paths = await reader.paths()
  const slicerParts = paths.filter((path) => /^xl\/slicers\/[^/]+\.xml$/.test(path)).sort()
  if (slicerParts.length === 0) return []

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

  // Workbook pivotCaches: cacheId → pivotCacheDefinition part path.
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
  // tabular@pivotCacheId refers to the pivot cache definition's x14
  // pivotCacheId, an id space of its own — Excel pairs workbook cacheId="110"
  // with x14 pivotCacheId="2124480673". Resolve through the definitions'
  // x14 exts; the workbook cacheId map is only for packages whose definitions
  // carry no ext at all (older files here wrote the cacheId in both places).
  const cachePathByX14Id = new Map<number, string>()
  const cachePathByLegacyCacheId = new Map<number, string>()
  for (const [cacheId, definitionPath] of cachePathByCacheId) {
    if (!(await hasEntry(reader, definitionPath))) continue
    const x14CacheId = x14PivotCacheIdOf(await reader.readText(definitionPath))
    if (x14CacheId === null) {
      cachePathByLegacyCacheId.set(cacheId, definitionPath)
    } else if (!cachePathByX14Id.has(x14CacheId)) {
      cachePathByX14Id.set(x14CacheId, definitionPath)
    }
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
    if (!/^xl\/slicerCaches\/[^/]+\.xml$/.test(path)) continue
    const parsed = parseCachePart(path, await reader.readText(path))
    if (parsed) cachesByName.set(parsed.name, parsed)
  }

  // Worksheet rels: which sheet hosts which control part.
  const hostByPart = new Map<string, string>()
  for (const path of paths) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(path)) continue
    const relsPath = relsPathFor(path)
    if (!(await hasEntry(reader, relsPath))) continue
    for (const entry of parseRelationships(await reader.readText(relsPath))) {
      if (entry.type !== SLICER_REL_TYPE) continue
      hostByPart.set(resolveRelTarget(path, entry.target), path)
    }
  }

  const states: SlicerPartState[] = []
  for (const partPath of slicerParts) {
    const xml = await reader.readText(partPath)
    for (const control of parseSlicerControls(xml)) {
      const cache = cachesByName.get(control.cacheName)
      const host = hostByPart.get(partPath)
      if (cache === undefined || host === undefined) continue
      if (!cache.tabular) continue // OLAP caches: skip
      // Binding: tabular@pivotCacheId is the bound definition's x14 slicer id;
      // only a definition without any x14 ext falls back to the workbook
      // cacheId (tabId is a sheetId, never a cache discriminator).
      let pivotCacheDefinitionPath: string | null = null
      if (cache.pivotCacheId !== null) {
        pivotCacheDefinitionPath =
          cachePathByX14Id.get(cache.pivotCacheId) ??
          cachePathByLegacyCacheId.get(cache.pivotCacheId) ??
          null
      }
      states.push({
        sheetId: sheetIdByWorksheet.get(host) ?? '',
        worksheetPath: host,
        partPath,
        cachePath: cache.path,
        name: control.name,
        cacheName: cache.name,
        sourceName: cache.sourceName,
        caption: control.caption,
        pivotCacheDefinitionPath,
        selectedItems: cache.selectedItems,
      })
    }
  }
  return states
}

async function hasEntry(reader: SlicerPartReader, path: string): Promise<boolean> {
  return (await reader.paths()).includes(path)
}
