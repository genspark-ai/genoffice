/**
 * Embedded font extraction — <p:embeddedFontLst> declares document-embedded faces whose
 * fntdata parts are EOT containers. PowerPoint renders missing families with these faces,
 * so fidelity needs them registered for measuring and drawing. Only uncompressed payloads
 * are usable here: MicroType-Express-compressed EOTs (flag 0x4) need a licensed
 * decompressor and are skipped.
 */
import { XMLParser } from 'fast-xml-parser'
import { PackageArchive, resolveTarget } from './zip'
import { asXmlNode, xmlArray } from './xml-utils'

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
  if (flags & EOT_FLAG_COMPRESSED) return null
  // The variable-length header precedes the payload, so the payload sits at the tail
  let sfnt = bytes.slice(bytes.length - fontDataSize)
  if (flags & EOT_FLAG_XOR) sfnt = sfnt.map((b) => b ^ 0x50)
  const magic = new DataView(sfnt.buffer, sfnt.byteOffset, 4).getUint32(0, false)
  if (!SFNT_MAGICS.includes(magic)) return null
  return sfnt
}

/**
 * Remove the package's embedded fonts: <p:embeddedFontLst> and the
 * embedTrueTypeFonts/saveSubsetFonts attributes in presentation.xml, the font
 * relationships in its rels, the ppt/fonts/*.fntdata parts, and the
 * [Content_Types] fntdata declaration.
 *
 * Why: the embedded subsets cannot be re-generated after a text edit (MicroType
 * Express compression needs a licensed compressor), and WPS honors a stale
 * subset strictly — glyphs it doesn't cover render as invisible text while the
 * in-app system-font fallback shows them fine. Stripping makes every viewer
 * fall back to system fonts. Idempotent; returns whether anything was stripped.
 */
export function stripEmbeddedFonts(archive: PackageArchive): boolean {
  const presPath = 'ppt/presentation.xml'
  const pres = archive.readText(presPath)
  if (!pres || !pres.includes('embeddedFont')) return false

  const nextPres = pres
    .replace(/<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/, '')
    .replace(/ embedTrueTypeFonts="[^"]*"/, '')
    .replace(/ saveSubsetFonts="[^"]*"/, '')
  archive.entries.set(presPath, Buffer.from(nextPres, 'utf8'))

  const relsPath = 'ppt/_rels/presentation.xml.rels'
  const rels = archive.readText(relsPath)
  if (rels) {
    const nextRels = rels.replace(/<Relationship\b[^>]*\bType="[^"]*\/font"[^>]*\/>/g, '')
    if (nextRels !== rels) archive.entries.set(relsPath, Buffer.from(nextRels, 'utf8'))
  }

  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (ct) {
    const nextCt = ct
      .replace(/<Default\b[^>]*\bExtension="fntdata"[^>]*\/>/g, '')
      .replace(/<Override\b[^>]*\bPartName="\/ppt\/fonts\/[^"]*"[^>]*\/>/g, '')
    if (nextCt !== ct) archive.entries.set(ctPath, Buffer.from(nextCt, 'utf8'))
  }

  for (const path of [...archive.entries.keys()]) {
    if (/^ppt\/fonts\/[^/]+\.fntdata$/.test(path)) archive.entries.delete(path)
  }
  return true
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
