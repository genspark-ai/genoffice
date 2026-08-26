/// Copying a slide between two decks. A bundle carries the slide's XML plus every
/// package part it depends on (media, charts and their embedded workbooks, …) so
/// the target deck can be a different file — or a different process.
///
/// Two things keep this simple. Relationship ids only have to be unique inside one
/// part's .rels, so the copied slide keeps the source's rIds verbatim and its XML
/// never needs rewriting. And the layout is not copied: the slide is re-pointed at
/// a layout in the destination, which is what PowerPoint's "use destination theme"
/// paste does.

import { PackageArchive, relsPathFor, resolveTarget, type Relationship } from './zip'

const LAYOUT_REL = '/slideLayout'
const NOTES_REL = '/notesSlide'

export interface SlideBundleRel {
  readonly id: string
  readonly type: string
  /** absolute source path for internal targets; the raw target for external ones */
  readonly target: string
  readonly external?: boolean
  /** true for the slideLayout rel, which the destination re-points */
  readonly layout?: boolean
}

/** The source layout→master→theme chain, for "keep source formatting" pastes. */
export interface SlideBundleChain {
  readonly layoutPath: string
  readonly masterPath: string
  readonly themePath?: string
  /** the chain parts and everything they reference, base64 like `parts` */
  readonly parts: Readonly<Record<string, string>>
  readonly contentTypes: Readonly<Record<string, string>>
}

export interface SlideBundle {
  readonly slideXml: string
  readonly rels: readonly SlideBundleRel[]
  /** every dependency part, base64 so the bundle survives IPC and JSON */
  readonly parts: Readonly<Record<string, string>>
  /** ContentType per dependency: keyed by lower-case extension or by part path */
  readonly contentTypes: Readonly<Record<string, string>>
  /** cSld/@name of the source layout, used to find its counterpart by name */
  readonly layoutName?: string
  /** source deck's slide size, so the caller can warn about a mismatch */
  readonly slideSize?: { cx: number; cy: number }
  readonly chain?: SlideBundleChain
}

function extOf(path: string): string {
  const dot = path.lastIndexOf('.')
  return dot < 0 ? '' : path.slice(dot + 1).toLowerCase()
}

function overrideContentType(ct: string | null, partPath: string): string | undefined {
  if (!ct) return undefined
  const re = new RegExp(
    `<Override[^>]*PartName="/${partPath.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*ContentType="([^"]+)"`,
  )
  return re.exec(ct)?.[1]
}

function defaultContentType(ct: string | null, ext: string): string | undefined {
  if (!ct) return undefined
  return new RegExp(`<Default[^>]*Extension="${ext}"[^>]*ContentType="([^"]+)"`, 'i').exec(ct)?.[1]
}

/** Walks a part's rel graph, collecting bytes + content types for everything reachable. */
export function collectPart(
  archive: PackageArchive,
  partPath: string,
  parts: Record<string, string>,
  contentTypes: Record<string, string>,
): void {
  if (parts[partPath]) return
  const bytes = archive.readBytes(partPath)
  if (!bytes) return
  parts[partPath] = Buffer.from(bytes).toString('base64')

  const ct = archive.readText('[Content_Types].xml')
  const override = overrideContentType(ct, partPath)
  if (override) contentTypes[partPath] = override
  else {
    const ext = extOf(partPath)
    const dflt = defaultContentType(ct, ext)
    if (ext && dflt) contentTypes[ext] = dflt
  }

  const relsPath = relsPathFor(partPath)
  const relsXml = archive.readText(relsPath)
  if (!relsXml) return
  parts[relsPath] = Buffer.from(relsXml, 'utf8').toString('base64')
  for (const rel of archive.readRels(partPath).values()) {
    if (rel.targetMode === 'External') continue
    collectPart(archive, resolveTarget(partPath, rel.target), parts, contentTypes)
  }
}

function layoutNameOf(archive: PackageArchive, layoutPath: string): string | undefined {
  const xml = archive.readText(layoutPath)
  return xml ? /<p:cSld[^>]*\sname="([^"]*)"/.exec(xml)?.[1] : undefined
}

/**
 * Snapshot a slide and its dependencies. `slideXml` is passed in so callers can
 * hand over the patched (unsaved-edits-included) XML.
 */
export function collectSlideBundle(
  archive: PackageArchive,
  slidePath: string,
  slideXml: string,
): SlideBundle {
  const parts: Record<string, string> = {}
  const contentTypes: Record<string, string> = {}
  const rels: SlideBundleRel[] = []
  let layoutName: string | undefined

  for (const rel of archive.readRels(slidePath).values()) {
    if (rel.type.endsWith(NOTES_REL)) continue
    if (rel.targetMode === 'External') {
      rels.push({ id: rel.id, type: rel.type, target: rel.target, external: true })
      continue
    }
    const absolute = resolveTarget(slidePath, rel.target)
    if (rel.type.endsWith(LAYOUT_REL)) {
      layoutName = layoutNameOf(archive, absolute)
      rels.push({ id: rel.id, type: rel.type, target: absolute, layout: true })
      continue
    }
    collectPart(archive, absolute, parts, contentTypes)
    rels.push({ id: rel.id, type: rel.type, target: absolute })
  }

  let slideSize: { cx: number; cy: number } | undefined
  try {
    slideSize = archive.readPresentation().size
  } catch {
    /* malformed source presentation: the size is only used for a warning */
  }

  let chain: SlideBundleChain | undefined
  const c = archive.resolveSlideChain(slidePath)
  if (c.layoutPath && c.masterPath) {
    const chainParts: Record<string, string> = {}
    const chainCT: Record<string, string> = {}
    // collectPart walks the rel graph: layout → master → theme + the master's other layouts + media
    collectPart(archive, c.layoutPath, chainParts, chainCT)
    chain = {
      layoutPath: c.layoutPath,
      masterPath: c.masterPath,
      ...(c.themePath ? { themePath: c.themePath } : {}),
      parts: chainParts,
      contentTypes: chainCT,
    }
  }

  return {
    slideXml,
    rels,
    parts,
    contentTypes,
    layoutName,
    ...(slideSize ? { slideSize } : {}),
    ...(chain ? { chain } : {}),
  }
}

/** First `<dir>/<base><n>.<ext>` free in the target archive and not yet allocated this pass. */
function freePath(archive: PackageArchive, sourcePath: string, taken: ReadonlySet<string>): string {
  const slash = sourcePath.lastIndexOf('/')
  const dir = sourcePath.slice(0, slash)
  const file = sourcePath.slice(slash + 1)
  const dot = file.lastIndexOf('.')
  const stem = dot < 0 ? file : file.slice(0, dot)
  const ext = dot < 0 ? '' : file.slice(dot)
  const base = /^(.*?)(\d+)$/.exec(stem)?.[1] ?? stem
  for (let n = 1; ; n += 1) {
    const candidate = `${dir}/${base}${n}${ext}`
    if (!archive.has(candidate) && !taken.has(candidate)) return candidate
  }
}

function relative(fromPart: string, toPath: string): string {
  const from = fromPart.slice(0, fromPart.lastIndexOf('/')).split('/').filter(Boolean)
  const to = toPath.split('/').filter(Boolean)
  let shared = 0
  while (shared < from.length && shared < to.length && from[shared] === to[shared]) shared += 1
  return [...Array.from({ length: from.length - shared }, () => '..'), ...to.slice(shared)].join(
    '/',
  )
}

function escapeAttr(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;')
}

/** All layout paths in the target, with their display names. */
export function listLayouts(archive: PackageArchive): { path: string; name?: string }[] {
  const layouts: { path: string; name?: string }[] = []
  for (const path of archive.entries.keys()) {
    if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path)) {
      layouts.push({ path, name: layoutNameOf(archive, path) })
    }
  }
  return layouts
}

/**
 * Writes a bundle's dependency parts into `archive` under fresh names and returns
 * the slide's .rels XML, with every rId preserved and the layout rel pointed at
 * `layoutPath`. The caller writes the slide part itself.
 */
/**
 * Writes a bag of bundled parts into `archive` under non-conflicting names,
 * rewriting each part's own rels to wherever its dependencies landed. Returns
 * the source→destination path map.
 */
function writeParts(
  archive: PackageArchive,
  parts: Readonly<Record<string, string>>,
  contentTypes: Readonly<Record<string, string>>,
  reusable?: ReadonlySet<string>,
): Map<string, string> {
  // rels parts follow their owner, so map the owners first
  const pathMap = new Map<string, string>()
  const taken = new Set<string>()
  for (const sourcePath of Object.keys(parts)) {
    if (sourcePath.includes('/_rels/')) continue
    if (reusable?.has(sourcePath)) {
      pathMap.set(sourcePath, sourcePath)
      continue
    }
    const dest =
      archive.has(sourcePath) || taken.has(sourcePath)
        ? freePath(archive, sourcePath, taken)
        : sourcePath
    taken.add(dest)
    pathMap.set(sourcePath, dest)
  }

  for (const [sourcePath, base64] of Object.entries(parts)) {
    if (sourcePath.includes('/_rels/') || reusable?.has(sourcePath)) continue
    const targetPath = pathMap.get(sourcePath)
    if (!targetPath) continue
    archive.entries.set(targetPath, new Uint8Array(Buffer.from(base64, 'base64')))

    const sourceRels = parts[relsPathFor(sourcePath)]
    if (!sourceRels) continue
    let relsXml = Buffer.from(sourceRels, 'base64').toString('utf8')
    // rewrite this part's own targets to wherever their parts landed
    relsXml = relsXml.replaceAll(/Target="([^"]+)"/g, (whole, target: string) => {
      if (/^[a-z]+:/i.test(target)) return whole
      const absolute = resolveTarget(sourcePath, target)
      const mapped = pathMap.get(absolute)
      return mapped ? `Target="${escapeAttr(relative(targetPath, mapped))}"` : whole
    })
    archive.entries.set(relsPathFor(targetPath), new Uint8Array(Buffer.from(relsXml, 'utf8')))
  }

  ensureContentTypes(archive, contentTypes, pathMap)
  return pathMap
}

/**
 * Parts the target archive already holds byte-identically (same path, same
 * bytes, same rels, dependencies transitively identical too) — those can be
 * shared instead of re-imported, so a same-deck paste never duplicates media.
 */
function reusableParts(
  archive: PackageArchive,
  parts: Readonly<Record<string, string>>,
): Set<string> {
  const memo = new Map<string, boolean>()
  const identical = (sourcePath: string): boolean => {
    const cached = memo.get(sourcePath)
    if (cached !== undefined) return cached
    memo.set(sourcePath, false) // cycle guard
    let ok = false
    const existing = archive.readBytes(sourcePath)
    if (existing && Buffer.from(existing).equals(Buffer.from(parts[sourcePath]!, 'base64'))) {
      const relsPath = relsPathFor(sourcePath)
      const bundledRels = parts[relsPath]
      const existingRels = archive.readText(relsPath)
      if (bundledRels == null) ok = existingRels == null
      else if (existingRels != null) {
        ok = Buffer.from(bundledRels, 'base64').toString('utf8') === existingRels
        if (ok) {
          for (const rel of archive.readRels(sourcePath).values()) {
            if (rel.targetMode === 'External') continue
            const dep = resolveTarget(sourcePath, rel.target)
            if (parts[dep] != null && !identical(dep)) {
              ok = false
              break
            }
          }
        }
      }
    }
    memo.set(sourcePath, ok)
    return ok
  }
  const reusable = new Set<string>()
  for (const sourcePath of Object.keys(parts)) {
    if (!sourcePath.includes('/_rels/') && identical(sourcePath)) reusable.add(sourcePath)
  }
  return reusable
}

/**
 * Lands an element-clipboard's dependency parts in `archive` (paste target).
 * Byte-identical parts already present are shared — a same-deck paste writes
 * nothing; a cross-deck paste imports media/charts under non-conflicting names.
 * Returns the source→destination path map.
 */
export function importClipboardParts(
  archive: PackageArchive,
  parts: Readonly<Record<string, string>>,
  contentTypes: Readonly<Record<string, string>>,
): Map<string, string> {
  return writeParts(archive, parts, contentTypes, reusableParts(archive, parts))
}

export function materializeSlideBundle(
  archive: PackageArchive,
  bundle: SlideBundle,
  slidePath: string,
  layoutPath: string,
): string {
  const pathMap = writeParts(archive, bundle.parts, bundle.contentTypes)

  const relLines = bundle.rels.map((rel) => {
    if (rel.layout) {
      return `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeAttr(relative(slidePath, layoutPath))}"/>`
    }
    if (rel.external) {
      return `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeAttr(rel.target)}" TargetMode="External"/>`
    }
    const mapped = pathMap.get(rel.target) ?? rel.target
    return `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeAttr(relative(slidePath, mapped))}"/>`
  })
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `${relLines.join('')}</Relationships>`
  )
}

function ensureContentTypes(
  archive: PackageArchive,
  contentTypes: Readonly<Record<string, string>>,
  pathMap: Map<string, string>,
): void {
  const ctPath = '[Content_Types].xml'
  let ct = archive.readText(ctPath)
  if (!ct) return
  for (const [key, contentType] of Object.entries(contentTypes)) {
    if (key.includes('/')) {
      const targetPath = pathMap.get(key)
      if (!targetPath) continue
      const override = `<Override PartName="/${targetPath}" ContentType="${contentType}"/>`
      if (!ct.includes(`PartName="/${targetPath}"`))
        ct = ct.replace('</Types>', `${override}</Types>`)
    } else if (!new RegExp(`<Default[^>]*Extension="${key}"`, 'i').test(ct)) {
      ct = ct.replace(
        '</Types>',
        `<Default Extension="${key}" ContentType="${contentType}"/></Types>`,
      )
    }
  }
  archive.entries.set(ctPath, new Uint8Array(Buffer.from(ct, 'utf8')))
}

const PRES_PATH = 'ppt/presentation.xml'
const MASTER_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'

function partText(parts: Readonly<Record<string, string>>, path: string): string | null {
  const base64 = parts[path]
  return base64 == null ? null : Buffer.from(base64, 'base64').toString('utf8')
}

/**
 * A target layout whose XML, owning master and theme are byte-identical to the
 * bundled chain — i.e. this chain was already imported (or the decks share the
 * template), so "keep source formatting" can reuse it instead of adding a master.
 */
function findIdenticalLayout(
  archive: PackageArchive,
  masterXml: string,
  themeXml: string | null,
  layoutXml: string,
): string | null {
  for (const path of archive.entries.keys()) {
    if (!/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(path)) continue
    if (archive.readText(path) !== masterXml) continue
    const rels = [...archive.readRels(path).values()]
    if (themeXml != null) {
      const themeRel = rels.find((r) => r.type.endsWith('/theme'))
      const themePath = themeRel ? resolveTarget(path, themeRel.target) : null
      if (!themePath || archive.readText(themePath) !== themeXml) continue
    }
    for (const rel of rels) {
      if (!rel.type.endsWith(LAYOUT_REL)) continue
      const layoutPath = resolveTarget(path, rel.target)
      if (archive.readText(layoutPath) === layoutXml) return layoutPath
    }
  }
  return null
}

/** Register an imported slideMaster with presentation.xml(.rels). */
function registerMaster(archive: PackageArchive, masterPath: string): boolean {
  const presRelsPath = relsPathFor(PRES_PATH)
  const pres = archive.readText(PRES_PATH)
  const presRels = archive.readText(presRelsPath)
  if (!pres || !presRels || !pres.includes('</p:sldMasterIdLst>')) return false
  let maxRid = 0
  for (const m of presRels.matchAll(/Id="rId(\d+)"/g)) maxRid = Math.max(maxRid, Number(m[1]))
  const rid = `rId${maxRid + 1}`
  archive.entries.set(
    presRelsPath,
    Buffer.from(
      presRels.replace(
        '</Relationships>',
        `<Relationship Id="${rid}" Type="${MASTER_REL_TYPE}" Target="${escapeAttr(relative(PRES_PATH, masterPath))}"/></Relationships>`,
      ),
      'utf8',
    ),
  )
  // sldMasterId ids must be unique and ≥ 2147483648
  let maxId = 2147483647
  for (const m of pres.matchAll(/<p:sldMasterId\s[^>]*id="(\d+)"/g))
    maxId = Math.max(maxId, Number(m[1]))
  archive.entries.set(
    PRES_PATH,
    Buffer.from(
      pres.replace(
        '</p:sldMasterIdLst>',
        `<p:sldMasterId id="${maxId + 1}" r:id="${rid}"/></p:sldMasterIdLst>`,
      ),
      'utf8',
    ),
  )
  return true
}

/**
 * "Keep source formatting": land the bundled layout→master→theme chain in the
 * target (registering the master) and return the layout path the pasted slide
 * should point at. Returns null when the bundle has no chain or the target's
 * presentation.xml can't take another master — callers fall back to chooseLayout.
 */
export function importSourceLayout(archive: PackageArchive, bundle: SlideBundle): string | null {
  const chain = bundle.chain
  if (!chain) return null
  const layoutXml = partText(chain.parts, chain.layoutPath)
  const masterXml = partText(chain.parts, chain.masterPath)
  if (layoutXml == null || masterXml == null) return null
  const themeXml = chain.themePath ? partText(chain.parts, chain.themePath) : null

  const existing = findIdenticalLayout(archive, masterXml, themeXml, layoutXml)
  if (existing) return existing

  const pres = archive.readText(PRES_PATH)
  if (!pres || !pres.includes('</p:sldMasterIdLst>')) return null
  const pathMap = writeParts(archive, chain.parts, chain.contentTypes)
  const masterPath = pathMap.get(chain.masterPath)
  const layoutPath = pathMap.get(chain.layoutPath)
  if (!masterPath || !layoutPath || !registerMaster(archive, masterPath)) return null
  return layoutPath
}

/** The destination layout for a pasted slide: same name first, else the neighbour's. */
export function chooseLayout(
  archive: PackageArchive,
  bundle: SlideBundle,
  neighbourSlidePath?: string,
): string | null {
  const layouts = listLayouts(archive)
  if (bundle.layoutName) {
    const byName = layouts.find((layout) => layout.name === bundle.layoutName)
    if (byName) return byName.path
  }
  if (neighbourSlidePath) {
    const chain = archive.resolveSlideChain(neighbourSlidePath)
    if (chain.layoutPath) return chain.layoutPath
  }
  return layouts[0]?.path ?? null
}

export type { Relationship }
