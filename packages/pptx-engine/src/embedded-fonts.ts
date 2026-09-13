/**
 * Embedded font extraction — <p:embeddedFontLst> declares document-embedded faces whose
 * fntdata parts are EOT containers. PowerPoint renders missing families with these faces,
 * so fidelity needs them registered for measuring and drawing. Only uncompressed payloads
 * are usable directly; MicroType-Express-compressed EOTs (flag 0x4) inflate through the
 * vendored MTX decoder (vendor/mtx).
 */
import { XMLParser } from 'fast-xml-parser'
import { mtxToSfnt } from './vendor/mtx'
import { PackageArchive, resolveTarget } from './zip'
import { asXmlNode, xmlArray } from './xml-utils'
import type { GroupElement, Slide, SlideElement, TableElement, TextElement, TextRun } from './types'

export type EmbeddedFontStyle = 'regular' | 'bold' | 'italic' | 'boldItalic'

export interface EmbeddedFontFace {
  typeface: string
  style: EmbeddedFontStyle
  /** Single-face sfnt bytes (TTF/OTF), ready for opentype.js / FontFace */
  sfnt: Uint8Array
}

const EOT_FLAG_COMPRESSED = 0x00000004 // TTEMBED_TTCOMPRESSED (MicroType Express)
const EOT_FLAG_XOR = 0x10000000 // TTEMBED_XORENCRYPTDATA (payload XORed with 0x50)

const SFNT_MAGICS = [0x00010000, 0x4f54544f, 0x74727565, 0x74746366] // TrueType / 'OTTO' / 'true' / 'ttcf'

/** Unwrap an EOT container to its sfnt payload; null when compressed or malformed. */
export function eotToSfnt(bytes: Uint8Array): Uint8Array | null {
  if (bytes.length < 16) return null
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  // Some producers (LibreOffice) store the bare sfnt without an EOT wrapper
  if (SFNT_MAGICS.includes(dv.getUint32(0, false))) return bytes
  const eotSize = dv.getUint32(0, true)
  const fontDataSize = dv.getUint32(4, true)
  const flags = dv.getUint32(12, true)
  if (eotSize !== bytes.length || fontDataSize === 0 || fontDataSize > bytes.length - 16)
    return null
  // The variable-length header precedes the payload, so the payload sits at the tail
  let sfnt: Uint8Array = bytes.slice(bytes.length - fontDataSize)
  if (flags & EOT_FLAG_XOR) sfnt = sfnt.map((b) => b ^ 0x50)
  if (flags & EOT_FLAG_COMPRESSED) {
    const inflated = mtxToSfnt(sfnt)
    if (!inflated) return null
    sfnt = inflated
  }
  const magic = new DataView(sfnt.buffer, sfnt.byteOffset, 4).getUint32(0, false)
  if (!SFNT_MAGICS.includes(magic)) return null
  return sfnt
}

/**
 * Code points each embedded part must cover, keyed by `${typeface}|${style}`
 * (typeface lower-cased; '*' collects runs with no resolved family). PowerPoint
 * subsets every style file on its own, so a bold run only asks of the bold part.
 */
export type GlyphDemand = Map<string, Set<number>>

const EMBEDDED_STYLES: EmbeddedFontStyle[] = ['regular', 'bold', 'italic', 'boldItalic']

export function demandKey(typeface: string, style: EmbeddedFontStyle): string {
  return `${typeface.toLowerCase()}|${style}`
}

function runStyle(r: TextRun): EmbeddedFontStyle {
  if (r.bold && r.italic) return 'boldItalic'
  if (r.bold) return 'bold'
  if (r.italic) return 'italic'
  return 'regular'
}

const EAST_ASIAN_RE =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\u3000-\u303F\uFF00-\uFFEF]/u

function collectRuns(el: SlideElement, out: TextRun[]): void {
  if (el.type === 'group') {
    for (const c of (el as GroupElement).children) collectRuns(c, out)
    return
  }
  if (el.type === 'table') {
    for (const row of (el as TableElement).rows)
      for (const cell of row) for (const p of cell.text?.paragraphs ?? []) out.push(...p.runs)
    return
  }
  if (el.type === 'text' || el.type === 'shape')
    for (const p of (el as TextElement).text?.paragraphs ?? []) out.push(...p.runs)
}

/**
 * Glyphs the edited text needs, per typeface. Only edited elements count: the
 * producer's subsets already cover the untouched text. Null when nothing is edited.
 */
export function editedGlyphDemand(deck: { slides: Slide[] }): GlyphDemand | null {
  const runs: TextRun[] = []
  let edited = false
  for (const s of deck.slides) {
    for (const el of s.elements) {
      if (!s.structureDirty && !el.dirty) continue
      edited = true
      collectRuns(el, runs)
    }
  }
  if (!edited) return null
  const demand: GlyphDemand = new Map()
  for (const r of runs) {
    const style = runStyle(r)
    const fam = r.fontFamily ?? '*'
    const latin = r.latinFamily
    for (const ch of r.text) {
      const cp = ch.codePointAt(0)!
      if (cp <= 0x20 || /\s/u.test(ch)) continue
      const key = demandKey(latin && !EAST_ASIAN_RE.test(ch) ? latin : fam, style)
      let set = demand.get(key)
      if (!set) demand.set(key, (set = new Set()))
      set.add(cp)
    }
  }
  return demand
}

/** Code-point lookup over the sfnt's cmap (formats 4 and 12); null when unreadable. */
export function sfntCmapLookup(sfnt: Uint8Array): ((cp: number) => boolean) | null {
  try {
    const dv = new DataView(sfnt.buffer, sfnt.byteOffset, sfnt.byteLength)
    let base = 0
    if (dv.getUint32(0, false) === 0x74746366) base = dv.getUint32(12, false)
    const numTables = dv.getUint16(base + 4, false)
    let cmap = -1
    for (let i = 0; i < numTables; i++) {
      const rec = base + 12 + 16 * i
      if (dv.getUint32(rec, false) === 0x636d6170) cmap = dv.getUint32(rec + 8, false)
    }
    if (cmap < 0) return null
    const lookups: ((cp: number) => boolean)[] = []
    const n = dv.getUint16(cmap + 2, false)
    for (let i = 0; i < n; i++) {
      const sub = cmap + dv.getUint32(cmap + 4 + 8 * i + 4, false)
      const format = dv.getUint16(sub, false)
      if (format === 4) {
        const segCount = dv.getUint16(sub + 6, false) >> 1
        const ends = sub + 14
        const starts = ends + segCount * 2 + 2
        const deltas = starts + segCount * 2
        const rangeOffsets = deltas + segCount * 2
        lookups.push((cp) => {
          if (cp > 0xffff) return false
          for (let k = 0; k < segCount; k++) {
            const end = dv.getUint16(ends + k * 2, false)
            if (cp > end) continue
            const start = dv.getUint16(starts + k * 2, false)
            if (cp < start) return false
            const delta = dv.getUint16(deltas + k * 2, false)
            const ro = dv.getUint16(rangeOffsets + k * 2, false)
            if (ro === 0) return ((cp + delta) & 0xffff) !== 0
            const g = dv.getUint16(rangeOffsets + k * 2 + ro + (cp - start) * 2, false)
            return g !== 0 && ((g + delta) & 0xffff) !== 0
          }
          return false
        })
      } else if (format === 12) {
        const nGroups = dv.getUint32(sub + 12, false)
        lookups.push((cp) => {
          for (let k = 0; k < nGroups; k++) {
            const g = sub + 16 + k * 12
            const start = dv.getUint32(g, false)
            const end = dv.getUint32(g + 4, false)
            if (cp >= start && cp <= end) return dv.getUint32(g + 8, false) + (cp - start) !== 0
          }
          return false
        })
      }
    }
    if (!lookups.length) return null
    return (cp) => lookups.some((f) => f(cp))
  } catch {
    return null
  }
}

interface EmbeddedFontBlock {
  xml: string
  typeface: string
  relIds: string[]
  parts: Partial<Record<EmbeddedFontStyle, string>>
}

function embeddedFontBlocks(pres: string): EmbeddedFontBlock[] {
  const out: EmbeddedFontBlock[] = []
  for (const m of pres.matchAll(/<p:embeddedFont>[\s\S]*?<\/p:embeddedFont>/g)) {
    const typeface = /<p:font\b[^>]*\btypeface="([^"]*)"/.exec(m[0])?.[1] ?? ''
    const parts: EmbeddedFontBlock['parts'] = {}
    for (const r of m[0].matchAll(/<p:(regular|bold|italic|boldItalic)\b[^>]*\br:id="([^"]*)"/g))
      parts[r[1] as EmbeddedFontStyle] = r[2]!
    out.push({ xml: m[0], typeface, relIds: Object.values(parts), parts })
  }
  return out
}

/**
 * Embedded typefaces whose subsets no longer cover the glyphs the edited text
 * asks of them. Each style is checked against its own part; a style with no
 * part of its own renders from the regular one (synthetic bold/italic). A part
 * that cannot be decoded or has no readable cmap counts as stale when anything
 * demands it: unverifiable is not safe.
 */
export function staleEmbeddedTypefaces(archive: PackageArchive, demand: GlyphDemand): string[] {
  const pres = archive.readText('ppt/presentation.xml')
  if (!pres || !pres.includes('embeddedFont')) return []
  const rels = archive.readRels('ppt/presentation.xml')
  const lookups = new Map<string, ((cp: number) => boolean) | null>()
  const lookupFor = (relId: string) => {
    if (!lookups.has(relId)) {
      const rel = rels.get(relId)
      const bytes = rel
        ? archive.readBytes(resolveTarget('ppt/presentation.xml', rel.target))
        : null
      const sfnt = bytes ? eotToSfnt(bytes) : null
      lookups.set(relId, sfnt ? sfntCmapLookup(sfnt) : null)
    }
    return lookups.get(relId) ?? null
  }
  const stale: string[] = []
  for (const face of embeddedFontBlocks(pres)) {
    if (!face.relIds.length) continue
    const covers = EMBEDDED_STYLES.every((style) => {
      const needed = new Set([
        ...(demand.get(demandKey(face.typeface, style)) ?? []),
        ...(demand.get(demandKey('*', style)) ?? []),
      ])
      if (!needed.size) return true
      const relId = face.parts[style] ?? face.parts.regular ?? face.relIds[0]!
      const lookup = lookupFor(relId)
      return !!lookup && [...needed].every(lookup)
    })
    if (!covers) stale.push(face.typeface)
  }
  return stale
}

/**
 * Remove embedded faces from the package: their <p:embeddedFont> entries, font
 * relationships, fntdata parts and [Content_Types] overrides. With no typeface
 * filter every face goes; once no face remains the <p:embeddedFontLst> itself,
 * the embedTrueTypeFonts/saveSubsetFonts attributes and the fntdata Default are
 * dropped too. Idempotent; returns whether anything was removed.
 */
export function stripEmbeddedFonts(
  archive: PackageArchive,
  typefaces?: ReadonlySet<string>,
): boolean {
  const presPath = 'ppt/presentation.xml'
  const pres = archive.readText(presPath)
  if (!pres || !pres.includes('embeddedFont')) return false
  const wanted = typefaces && new Set([...typefaces].map((t) => t.toLowerCase()))
  const blocks = embeddedFontBlocks(pres)
  const removed = blocks.filter((b) => !wanted || wanted.has(b.typeface.toLowerCase()))
  if (!removed.length) return false
  const removeAll = removed.length === blocks.length

  let nextPres = pres
  for (const b of removed) nextPres = nextPres.replace(b.xml, '')
  if (removeAll) {
    nextPres = nextPres
      .replace(
        /<p:embeddedFontLst\b[^>]*\/>|<p:embeddedFontLst\b[^>]*>[\s\S]*?<\/p:embeddedFontLst>/,
        '',
      )
      .replace(/ embedTrueTypeFonts="[^"]*"/, '')
      .replace(/ saveSubsetFonts="[^"]*"/, '')
  }
  archive.entries.set(presPath, Buffer.from(nextPres, 'utf8'))

  const relIds = new Set(removed.flatMap((b) => b.relIds))
  const parts = new Set<string>()
  const relsPath = 'ppt/_rels/presentation.xml.rels'
  const rels = archive.readText(relsPath)
  if (rels) {
    const nextRels = rels.replace(/<Relationship\b[^>]*\/>/g, (tag) => {
      const isFont = /\bType="[^"]*\/font"/.test(tag)
      const id = /\bId="([^"]*)"/.exec(tag)?.[1] ?? ''
      if (!isFont || !(removeAll || relIds.has(id))) return tag
      const target = /\bTarget="([^"]*)"/.exec(tag)?.[1]
      if (target) parts.add(resolveTarget(presPath, target))
      return ''
    })
    if (nextRels !== rels) archive.entries.set(relsPath, Buffer.from(nextRels, 'utf8'))
  }
  if (removeAll)
    for (const path of archive.entries.keys())
      if (/^ppt\/fonts\/[^/]+\.fntdata$/.test(path)) parts.add(path)
  for (const path of parts) archive.entries.delete(path)

  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (ct) {
    let nextCt = ct.replace(/<Override\b[^>]*\/>/g, (tag) => {
      const partName = /\bPartName="([^"]*)"/.exec(tag)?.[1] ?? ''
      return parts.has(partName.replace(/^\//, '')) ? '' : tag
    })
    if (removeAll) nextCt = nextCt.replace(/<Default\b[^>]*\bExtension="fntdata"[^>]*\/>/g, '')
    if (nextCt !== ct) archive.entries.set(ctPath, Buffer.from(nextCt, 'utf8'))
  }
  return true
}

/**
 * Save-time guard: an embedded subset cannot grow, so text edited into glyphs
 * it lacks renders as invisible text in WPS (which honors the subset strictly,
 * unlike Chromium's per-glyph fallback in the app). Faces the edits outgrew are
 * stripped so viewers fall back to system fonts; covered faces, and untouched
 * decks, stay byte-identical. Returns the stripped typefaces.
 */
export function stripStaleEmbeddedFonts(
  deck: { slides: Slide[] },
  archive: PackageArchive,
): string[] {
  const demand = editedGlyphDemand(deck)
  if (!demand) return []
  const stale = staleEmbeddedTypefaces(archive, demand)
  if (stale.length) stripEmbeddedFonts(archive, new Set(stale))
  return stale
}

/** Usable (uncompressed) embedded faces of a package; empty when none are declared. */
export function listEmbeddedFonts(archive: PackageArchive): EmbeddedFontFace[] {
  const presXml = archive.readText('ppt/presentation.xml')
  if (!presXml || !presXml.includes('embeddedFont')) return []
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    isArray: (name) => name === 'p:embeddedFont',
  })
  let root: Record<string, unknown>
  try {
    root = asXmlNode(asXmlNode(parser.parse(presXml))['p:presentation'])
  } catch {
    return []
  }
  const lst = asXmlNode(root['p:embeddedFontLst'] ?? {})
  const rels = archive.readRels('ppt/presentation.xml')
  const out: EmbeddedFontFace[] = []
  for (const ef of xmlArray(lst['p:embeddedFont'])) {
    const typeface = String(asXmlNode(asXmlNode(ef)['p:font'] ?? {})['@_typeface'] ?? '')
    if (!typeface) continue
    const styles: EmbeddedFontStyle[] = ['regular', 'bold', 'italic', 'boldItalic']
    for (const style of styles) {
      const relId = asXmlNode(asXmlNode(ef)[`p:${style}`] ?? {})['@_r:id']
      if (relId == null) continue
      const rel = rels.get(String(relId))
      if (!rel) continue
      const part = resolveTarget('ppt/presentation.xml', rel.target)
      const bytes = archive.readBytes(part)
      if (!bytes) continue
      const sfnt = eotToSfnt(bytes)
      if (sfnt) out.push({ typeface, style, sfnt })
    }
  }
  return out
}
