/**
 * Slide master / layout edit view support (6.16).
 *
 * Exception to the fidelity rule: only layout/master parts the user actively edited
 * in master view are written back, reusing the slide byte-surgery path (scanSlide
 * anchors + patchSlideXml); untouched elements pass through with original bytes.
 * This module only enumerates and parses; the caller writes back via patchSlideXml
 * and then writes the archive entry.
 */
import type { PackageArchive } from './zip'
import { resolveTarget } from './zip'
import { parseClrMap, parseTheme } from './theme'
import { parseSlide, parseDecorations, type ParseContext } from './parse'
import { parsePlaceholderMap, parseMasterTextStyles } from './placeholder'
import type { Slide } from './types'

export interface MasterPartInfo {
  /** Path inside the zip, e.g. ppt/slideMasters/slideMaster1.xml */
  partPath: string
  kind: 'master' | 'layout'
  /** <p:cSld name="…">, falling back to the file name */
  name: string
}

const partNum = (p: string) => parseInt(/(\d+)\.xml$/.exec(p)?.[1] ?? '0', 10)

function cSldName(xml: string | null | undefined, fallback: string): string {
  const m = xml ? /<p:cSld\s[^>]*name="([^"]*)"/.exec(xml) : null
  return m?.[1] || fallback
}

/** Enumerate parts for the master edit view: each master first, followed by its layouts. */
export function listMasterParts(archive: PackageArchive): MasterPartInfo[] {
  const masters = [...archive.entries.keys()]
    .filter((p) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(p))
    .sort((a, b) => partNum(a) - partNum(b))
  const out: MasterPartInfo[] = []
  for (const m of masters) {
    out.push({ partPath: m, kind: 'master', name: cSldName(archive.readText(m), baseName(m)) })
    const layouts = [...archive.readRels(m).values()]
      .filter((r) => r.type.endsWith('/slideLayout'))
      .map((r) => resolveTarget(m, r.target))
      .sort((a, b) => partNum(a) - partNum(b))
    for (const l of layouts) {
      out.push({ partPath: l, kind: 'layout', name: cSldName(archive.readText(l), baseName(l)) })
    }
  }
  return out
}

function baseName(p: string): string {
  return p.slice(p.lastIndexOf('/') + 1, -4)
}

/** A part's own image rels (same logic as index.ts partMediaRels, duplicated to avoid a circular dependency). */
function partMedia(archive: PackageArchive, partPath: string): Map<string, string> {
  const media = new Map<string, string>()
  for (const rel of archive.readRels(partPath).values()) {
    if (rel.type.endsWith('/image')) media.set(rel.id, resolveTarget(partPath, rel.target))
  }
  return media
}

/**
 * Parse a layout/master part into an editable Slide (its spTree structure is isomorphic
 * to a slide's, so scanSlide applies directly).
 * - master: text styles come from its own <p:txStyles>; placeholders always have an
 *   explicit xfrm, so no geometry inheritance is needed.
 * - layout: placeholder geometry/text styles/background inherit from its master; the
 *   master decoration layer is rendered underneath.
 */
export function parseMasterPart(archive: PackageArchive, partPath: string): Slide | null {
  const xml = archive.readText(partPath)
  if (!xml) return null
  const isMaster = partPath.includes('/slideMasters/')

  let masterPath: string | undefined = isMaster ? partPath : undefined
  if (!isMaster) {
    for (const rel of archive.readRels(partPath).values()) {
      if (rel.type.endsWith('/slideMaster')) {
        masterPath = resolveTarget(partPath, rel.target)
        break
      }
    }
  }

  const ctx: ParseContext = {}
  if (masterPath) {
    for (const rel of archive.readRels(masterPath).values()) {
      if (rel.type.endsWith('/theme')) {
        const themeXml = archive.readText(resolveTarget(masterPath, rel.target))
        if (themeXml) ctx.theme = parseTheme(themeXml)
        break
      }
    }
  }
  ctx.mediaRels = partMedia(archive, partPath)
  ctx.tableStyles = archive.readText('ppt/tableStyles.xml') ?? undefined

  const masterXml = !isMaster && masterPath ? archive.readText(masterPath) : undefined
  if (ctx.theme)
    ctx.theme.clrMap = isMaster ? parseClrMap(xml) : parseClrMap(masterXml ?? undefined, xml)
  if (isMaster) {
    ctx.masterTextStyles = parseMasterTextStyles(xml, ctx.theme)
  } else if (masterXml) {
    ctx.masterPlaceholders = parsePlaceholderMap(masterXml, ctx.theme)
    ctx.masterTextStyles = parseMasterTextStyles(masterXml, ctx.theme)
    ctx.masterBg = masterXml
    if (masterPath) ctx.masterMediaRels = partMedia(archive, masterPath)
  }

  const slide = parseSlide({ path: partPath, slideXml: xml, masterPath, ctx })

  // Layout view renders master concrete shapes underneath (master decorations stay visible when editing a layout)
  if (!isMaster && masterXml && masterPath) {
    if (!/<p:sldLayout\b[^>]*showMasterSp="(?:0|false)"/.test(xml)) {
      const dctx: ParseContext = { theme: ctx.theme, mediaRels: partMedia(archive, masterPath) }
      const dec = parseDecorations(masterXml, dctx, {})
      if (dec.length) slide.decorations = dec
    }
  }
  return slide
}
