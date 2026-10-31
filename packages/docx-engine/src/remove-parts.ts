/**
 * remove_parts: drop metadata-carrying package parts (customXml stores and
 * web-extension/add-in parts) from a saved .docx, with the package wiring that
 * would otherwise dangle (relationships, [Content_Types].xml entries, empty
 * folders). Used by the docs `remove_parts` op for pre-filing scrubs.
 *
 * The pass is deliberately narrow: it only removes the parts that were asked
 * for, plus their own relationship parts and now-dangling relationship
 * elements / content-type entries. Every other part keeps its bytes.
 *
 * A removed customXml store can back block-level content controls through
 * w:dataBinding/w:storeItemID. By default the save refuses (the control would
 * be orphaned); `dataBinding: "strip"` removes the w:dataBinding element and
 * keeps the control with its current text.
 */
import JSZip from 'jszip'

import { resolveMainDocumentPath, resolveRelationshipTargetPath } from './parse-package'
import { loadDocxZip } from './zip-load'

/** Metadata-carrying part families the docs remove_parts op can target. */
export const REMOVE_PARTS_KINDS = ['customXml', 'webExtensions'] as const
export type RemovePartsKind = (typeof REMOVE_PARTS_KINDS)[number]

const KIND_PREFIXES: Record<RemovePartsKind, string> = {
  /** template / DMS data stores: item*.xml, itemProps*.xml and their rels */
  customXml: 'customXml/',
  /** task panes / add-ins: webextension*.xml, taskpanes.xml and their rels */
  webExtensions: 'word/webextensions/',
}

/** Structural package parts an explicit path may not remove (removal would corrupt the file). */
const PROTECTED_PART_PATHS = new Set(['[Content_Types].xml', '_rels/.rels'])

/** Story parts that can carry a block-level w:sdt with a w:dataBinding. */
const STORY_PART_PATHS = [
  'word/document.xml',
  'word/footnotes.xml',
  'word/endnotes.xml',
  'word/glossary/document.xml',
]
const STORY_PART_RE = /^word\/(?:header|footer)[^/]*\.xml$/i

const RELATIONSHIP_TAG_RE =
  /<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*(?:\/\s*>|>\s*<\/(?:[A-Za-z_][\w.-]*:)?Relationship\s*>)/g
const OVERRIDE_TAG_RE =
  /<(?:[A-Za-z_][\w.-]*:)?Override\b[^>]*(?:\/\s*>|>\s*<\/(?:[A-Za-z_][\w.-]*:)?Override\s*>)/g
const DEFAULT_TAG_RE =
  /<(?:[A-Za-z_][\w.-]*:)?Default\b[^>]*(?:\/\s*>|>\s*<\/(?:[A-Za-z_][\w.-]*:)?Default\s*>)/g
const SDT_PR_RE = /<w:sdtPr\b[^>]*>[\s\S]*?<\/w:sdtPr>/g
const DATA_BINDING_RE = /<w:dataBinding\b[^>]*(?:\/>|>[\s\S]*?<\/w:dataBinding>)/

export interface RemovePartsOptions {
  /** explicit package paths (relative to the package root, no leading slash); unknown paths are reported and ignored */
  paths?: string[]
  /** whole part families to remove */
  kinds?: RemovePartsKind[]
  /** what to do when a removed customXml store backs a w:dataBinding control: refuse (default) or strip */
  dataBinding?: 'refuse' | 'strip'
}

/** One content control bound to a customXml store that is being removed. */
export interface RemovedPartBinding {
  /** story part carrying the control (word/document.xml, word/header1.xml, ...) */
  part: string
  /** w:sdtPr/w:tag of the control, when present */
  tag?: string
  /** w:sdtPr/w:id of the control, when present */
  controlId?: string
  /** ds:itemID the control is bound to, as written in w:storeItemID */
  storeItemID: string
  /** the customXml/itemProps*.xml part declaring that itemID */
  itemPropsPath: string
}

export interface RemovePartsPlan {
  /** package paths removal drops (or has already dropped), sorted */
  removedPaths: string[]
  /** explicit paths that are not present in the package (ignored) */
  unknownPaths: string[]
  /** content controls bound to removed customXml stores */
  dataBindings: RemovedPartBinding[]
}

function decodeXmlEntities(value: string): string {
  const codePoint = (match: string, raw: string, radix: number) => {
    const value = parseInt(raw, radix)
    return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
      ? String.fromCodePoint(value)
      : match
  }
  return value
    .replace(/&#x([0-9a-f]+);/gi, (match, hex: string) => codePoint(match, hex, 16))
    .replace(/&#(\d+);/g, (match, decimal: string) => codePoint(match, decimal, 10))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

/** An attribute by local name, tolerating any namespace prefix (ds:itemID, w:val, ...) and both quote styles. */
function attributeOf(tag: string, localName: string): string | undefined {
  const escaped = localName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const match = new RegExp(
    `(?:^|\\s)(?:[A-Za-z_][\\w.-]*:)?${escaped}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`,
  ).exec(tag)
  const value = match?.[1] ?? match?.[2]
  return value === undefined ? undefined : decodeXmlEntities(value)
}

function normalizePartPath(path: string): string | null {
  const parts: string[] = []
  for (const segment of path.replace(/\\/g, '/').replace(/^\/+/, '').split('/')) {
    if (!segment || segment === '.') continue
    if (segment === '..') {
      if (parts.length === 0) return null
      parts.pop()
    } else {
      parts.push(segment)
    }
  }
  return parts.join('/') || null
}

function decodePartUri(uri: string): string {
  try {
    return decodeURIComponent(uri)
  } catch {
    return uri
  }
}

function extensionOf(path: string): string | null {
  const name = path.slice(path.lastIndexOf('/') + 1)
  const dot = name.lastIndexOf('.')
  return dot > 0 ? name.slice(dot + 1).toLowerCase() : null
}

/** The source part a relationships part belongs to ("_rels/.rels" is the package root). */
function relationshipSourcePath(relsPath: string): string | null {
  if (relsPath === '_rels/.rels') return ''
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/.exec(relsPath)
  return match ? `${match[1] ?? ''}${match[2]}` : null
}

/** The relationships part that belongs to a part ("customXml/item1.xml" -> "customXml/_rels/item1.xml.rels"). */
function relationshipPartPath(partPath: string): string {
  const slash = partPath.lastIndexOf('/')
  const dir = slash >= 0 ? partPath.slice(0, slash + 1) : ''
  const filename = slash >= 0 ? partPath.slice(slash + 1) : partPath
  return `${dir}_rels/${filename}.rels`
}

/** GUID comparison ignores braces and case; Word writes storeItemID/itemID with braces. */
function normalizeItemId(value: string): string {
  return value.replace(/[{}]/g, '').trim().toUpperCase()
}

/** The parts an option set selects, plus explicit paths that are not in the package. */
function wantedParts(
  zip: JSZip,
  options: RemovePartsOptions,
): { paths: Set<string>; unknownPaths: string[] } {
  const paths = new Set<string>()
  for (const kind of options.kinds ?? []) {
    const prefix = KIND_PREFIXES[kind]
    if (!prefix) continue
    for (const name of Object.keys(zip.files)) {
      if (!zip.files[name]!.dir && name.startsWith(prefix)) paths.add(name)
    }
  }
  const unknownPaths: string[] = []
  for (const raw of options.paths ?? []) {
    const path = normalizePartPath(String(raw))
    if (!path || !zip.file(path) || PROTECTED_PART_PATHS.has(path)) {
      unknownPaths.push(String(raw))
      continue
    }
    paths.add(path)
  }
  return { paths, unknownPaths }
}

/** Add the relationship part of every selected part (it belongs to the part and leaves with it). */
function withOwnRels(zip: JSZip, paths: Set<string>): Set<string> {
  const out = new Set(paths)
  for (const path of paths) {
    if (/\.rels$/i.test(path)) continue
    const relsPath = relationshipPartPath(path)
    if (zip.file(relsPath)) out.add(relsPath)
  }
  return out
}

/** ds:itemID of every removed customXml itemProps part, keyed by the normalized GUID. */
async function itemIdsOf(zip: JSZip, removeSet: Set<string>): Promise<Map<string, string>> {
  const ids = new Map<string, string>()
  for (const path of removeSet) {
    if (!/^customXml\/itemProps[^/]*\.xml$/i.test(path)) continue
    const file = zip.file(path)
    if (!file) continue
    const xml = await file.async('string')
    const root = /<(?:[A-Za-z_][\w.-]*:)?datastoreItem\b[^>]*>/.exec(xml)?.[0]
    const itemID = root ? attributeOf(root, 'itemID') : undefined
    if (itemID) ids.set(normalizeItemId(itemID), path)
  }
  return ids
}

/** The story parts that may hold block-level content controls. */
async function storyPartsOf(zip: JSZip): Promise<string[]> {
  const paths = new Set(STORY_PART_PATHS)
  const docPath = await resolveMainDocumentPath(zip)
  if (docPath) paths.add(docPath)
  for (const name of Object.keys(zip.files)) {
    if (!zip.files[name]!.dir && STORY_PART_RE.test(name)) paths.add(name)
  }
  return [...paths].filter((path) => zip.file(path)).sort()
}

/**
 * Find (and strip, when asked) w:dataBinding elements whose w:storeItemID names
 * a removed customXml store. w:sdtPr does not nest, so the per-control regex walk
 * is exact; the control keeps its tag, id and current text.
 */
async function scanDataBindings(
  zip: JSZip,
  itemIds: Map<string, string>,
  storyParts: string[],
  strip: boolean,
): Promise<{ bindings: RemovedPartBinding[]; stripped: Map<string, string> }> {
  const bindings: RemovedPartBinding[] = []
  const stripped = new Map<string, string>()
  if (itemIds.size === 0) return { bindings, stripped }
  for (const part of storyParts) {
    const file = zip.file(part)
    if (!file) continue
    const xml = await file.async('string')
    const found: RemovedPartBinding[] = []
    const next = xml.replace(SDT_PR_RE, (sdtPr) => {
      const dataBinding = DATA_BINDING_RE.exec(sdtPr)?.[0]
      if (!dataBinding) return sdtPr
      const storeItemID = attributeOf(dataBinding, 'storeItemID')
      const itemPropsPath = storeItemID ? itemIds.get(normalizeItemId(storeItemID)) : undefined
      if (!storeItemID || !itemPropsPath) return sdtPr
      found.push({
        part,
        tag: attributeOf(/<w:tag\b[^>]*>/.exec(sdtPr)?.[0] ?? '', 'val'),
        controlId: attributeOf(/<w:id\b[^>]*>/.exec(sdtPr)?.[0] ?? '', 'val'),
        storeItemID,
        itemPropsPath,
      })
      return sdtPr.replace(dataBinding, '')
    })
    if (found.length === 0) continue
    bindings.push(...found)
    if (strip && next !== xml) stripped.set(part, next)
  }
  return { bindings, stripped }
}

function describeControl(binding: RemovedPartBinding): string {
  const label = binding.tag ? `"${binding.tag}"` : 'untagged control'
  const id = binding.controlId ? ` (w:id ${binding.controlId})` : ''
  return `${label}${id} in ${binding.part} is bound to storeItemID ${binding.storeItemID} of ${binding.itemPropsPath}`
}

/** The guided error a 'refuse' save throws when bound controls would be orphaned (the CLI reports the same text). */
export function removePartsRefusalMessage(bindings: RemovedPartBinding[]): string {
  return (
    `remove_parts would orphan content controls: ${bindings.map(describeControl).join('; ')}. ` +
    'Pass dataBinding: "strip" to unbind them and keep their current text, or remove the customXml parts in Word first.'
  )
}

interface ZipPlan {
  plan: RemovePartsPlan
  /** story part XML with the dataBinding elements removed (only when stripping) */
  stripped: Map<string, string>
}

async function planInZip(zip: JSZip, options: RemovePartsOptions): Promise<ZipPlan> {
  const { paths, unknownPaths } = wantedParts(zip, options)
  const removeSet = withOwnRels(zip, paths)
  const itemIds = await itemIdsOf(zip, removeSet)
  const storyParts = await storyPartsOf(zip)
  const strip = (options.dataBinding ?? 'refuse') === 'strip'
  const { bindings, stripped } = await scanDataBindings(zip, itemIds, storyParts, strip)
  return {
    plan: {
      removedPaths: [...removeSet].sort(),
      unknownPaths,
      dataBindings: bindings,
    },
    stripped,
  }
}

/** Rewrite every remaining relationships part so relations to removed parts are gone. */
async function dropDanglingRelationships(zip: JSZip, removed: Set<string>): Promise<void> {
  for (const path of Object.keys(zip.files)) {
    if (zip.files[path]!.dir || !path.endsWith('.rels')) continue
    const source = relationshipSourcePath(path)
    if (source === null) continue
    const file = zip.file(path)
    if (!file) continue
    const xml = await file.async('string')
    const next = xml.replace(RELATIONSHIP_TAG_RE, (tag) => {
      if (/TargetMode\s*=\s*["']external["']/i.test(tag)) return tag
      const target = attributeOf(tag, 'Target')
      if (!target) return tag
      const resolved = resolveRelationshipTargetPath(source, target)
      return resolved && removed.has(resolved) ? '' : tag
    })
    if (next !== xml) zip.file(path, next, { date: file.date })
  }
}

/** Drop overrides of removed parts, and Default extensions no remaining part uses. */
async function cleanContentTypes(zip: JSZip, removed: Set<string>): Promise<void> {
  const path = '[Content_Types].xml'
  const file = zip.file(path)
  if (!file) return
  const xml = await file.async('string')
  let next = xml.replace(OVERRIDE_TAG_RE, (tag) => {
    const partName = attributeOf(tag, 'PartName')
    if (!partName) return tag
    const normalized = normalizePartPath(decodePartUri(partName))
    return normalized && removed.has(normalized) ? '' : tag
  })
  const usedExtensions = new Set<string>()
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name]!.dir) continue
    const extension = extensionOf(name)
    if (extension) usedExtensions.add(extension)
  }
  next = next.replace(DEFAULT_TAG_RE, (tag) => {
    const extension = (attributeOf(tag, 'Extension') ?? '').toLowerCase()
    return extension && !usedExtensions.has(extension) ? '' : tag
  })
  if (next !== xml) zip.file(path, next, { date: file.date })
}

/** Drop explicit folder entries no remaining part sits under (some producers write them). */
function removeEmptyFolders(zip: JSZip): void {
  // snapshot first: zip.remove() on a folder also drops its entries from zip.files
  const entries = Object.entries(zip.files)
  const files = entries.filter(([, entry]) => !entry.dir).map(([name]) => name)
  for (const [name, entry] of entries) {
    if (!entry.dir) continue
    if (!files.some((file) => file.startsWith(name))) zip.remove(name)
  }
}

async function removePartsFromZip(zip: JSZip, removedPaths: string[]): Promise<void> {
  const removed = new Set(removedPaths)
  for (const path of removed) zip.remove(path)
  await dropDanglingRelationships(zip, removed)
  await cleanContentTypes(zip, removed)
  removeEmptyFolders(zip)
}

/**
 * Plan a remove_parts pass without changing anything: which package parts would
 * go, which explicit paths are missing, and which content controls are bound to
 * the removed customXml stores. Never throws: a caller that previews a 'refuse'
 * plan reports the bindings (and removePartsRefusalMessage) instead of writing.
 */
export async function planRemoveParts(
  bytes: Uint8Array,
  options: RemovePartsOptions,
): Promise<RemovePartsPlan> {
  const zip = await loadDocxZip(bytes)
  const { plan } = await planInZip(zip, options)
  return plan
}

/**
 * Apply a remove_parts pass to an assembled save zip: drop the selected parts,
 * their own rels parts, dangling relationship elements, content-type entries
 * and empty folders. Other parts keep their bytes. Throws when a bound control
 * would be orphaned and dataBinding is 'refuse'.
 */
export async function applyRemoveParts(
  zip: JSZip,
  options: RemovePartsOptions,
): Promise<RemovePartsPlan> {
  const { plan, stripped } = await planInZip(zip, options)
  if (plan.removedPaths.length === 0) return plan
  if (plan.dataBindings.length > 0) {
    if ((options.dataBinding ?? 'refuse') !== 'strip')
      throw new Error(removePartsRefusalMessage(plan.dataBindings))
    for (const [part, xml] of stripped) {
      const entry = zip.file(part)
      zip.file(part, xml, { date: entry?.date })
    }
  }
  await removePartsFromZip(zip, plan.removedPaths)
  return plan
}
