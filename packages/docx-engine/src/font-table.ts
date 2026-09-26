import type JSZip from 'jszip'
import { parseRels, resolveRelationshipTargetPath } from './parse-package'
import type {
  EmbeddedFontLineMetrics,
  EmbeddedFont,
  EmbeddedFontRef,
  EmbeddedFontSlot,
  FontTableEntry,
} from './types'
import { attrsOf, childrenOf, findChild, nameOf, xmlParser, type XNode } from './xml-utils'

export const FONT_TABLE_PART_PATH = 'word/fontTable.xml'

/** word/fontTable.xml → substitution hints (altName / PANOSE / family / pitch) */
export function parseFontTable(xml: string): FontTableEntry[] {
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(xml) as XNode[]
  } catch {
    return []
  }
  const root = parsed.find((n) => nameOf(n) === 'w:fonts')
  if (!root) return []
  const out: FontTableEntry[] = []
  for (const node of childrenOf(root)) {
    if (nameOf(node) !== 'w:font') continue
    const name = attrsOf(node)['w:name']
    if (!name) continue
    const val = (tag: string) => attrsOf(findChild(node, tag) ?? {})['w:val'] || undefined
    const entry: FontTableEntry = { name }
    const altName = val('w:altName')
    if (altName) entry.altName = altName
    const panose = val('w:panose1')
    if (panose) entry.panose = panose
    const family = val('w:family')
    if (family) entry.family = family
    const pitch = val('w:pitch')
    if (pitch) entry.pitch = pitch
    const embedded = parseEmbedSlots(node)
    if (embedded) entry.embedded = embedded
    out.push(entry)
  }
  return out
}

const EMBED_SLOTS: ReadonlyArray<[tag: string, slot: EmbeddedFontSlot]> = [
  ['w:embedRegular', 'regular'],
  ['w:embedBold', 'bold'],
  ['w:embedItalic', 'italic'],
  ['w:embedBoldItalic', 'boldItalic'],
]

function parseEmbedSlots(font: XNode): FontTableEntry['embedded'] | undefined {
  let out: FontTableEntry['embedded']
  for (const [tag, slot] of EMBED_SLOTS) {
    const attrs = attrsOf(findChild(font, tag) ?? {})
    const rId = attrs['r:id']
    if (!rId) continue
    const ref: EmbeddedFontRef = { rId }
    if (attrs['w:fontKey']) ref.fontKey = attrs['w:fontKey']
    ;(out ??= {})[slot] = ref
  }
  return out
}

const SFNT_MAGICS = new Set([0x00010000, 0x4f54544f, 0x74727565])

export function isSfnt(bytes: Uint8Array): boolean {
  if (bytes.length < 12) return false
  const magic = ((bytes[0] << 24) | (bytes[1] << 16) | (bytes[2] << 8) | bytes[3]) >>> 0
  return SFNT_MAGICS.has(magic)
}

/** hhea line box of a single sfnt face in em units; null without head/hhea */
export function sfntLineMetrics(bytes: Uint8Array): EmbeddedFontLineMetrics | null {
  if (!isSfnt(bytes)) return null
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const numTables = dv.getUint16(4)
  let unitsPerEm = 0
  let hhea = -1
  for (let t = 0; t < numTables; t++) {
    const rec = 12 + 16 * t
    if (rec + 16 > bytes.length) return null
    const tag = String.fromCharCode(bytes[rec], bytes[rec + 1], bytes[rec + 2], bytes[rec + 3])
    const off = dv.getUint32(rec + 8)
    const len = dv.getUint32(rec + 12)
    if (tag === 'head' && len >= 20 && off + 20 <= bytes.length) unitsPerEm = dv.getUint16(off + 18)
    if (tag === 'hhea' && len >= 10 && off + 10 <= bytes.length) hhea = off
  }
  if (unitsPerEm === 0 || hhea < 0) return null
  return {
    ascent: dv.getInt16(hhea + 4) / unitsPerEm,
    descent: -dv.getInt16(hhea + 6) / unitsPerEm,
    lineGap: dv.getInt16(hhea + 8) / unitsPerEm,
  }
}

/**
 * Undo the ECMA-376 17.8.1 obfuscation: the first 32 bytes are XORed with the
 * fontKey GUID's 16 bytes taken in reverse order. A missing/zero key or a part
 * that already carries an sfnt magic is returned as is.
 */
export function deobfuscateOdttf(bytes: Uint8Array, fontKey: string | undefined): Uint8Array {
  const hex = fontKey?.replace(/[^0-9a-f]/gi, '') ?? ''
  if (hex.length !== 32 || /^0+$/.test(hex) || isSfnt(bytes)) return bytes
  const key = new Uint8Array(16)
  for (let i = 0; i < 16; i++) key[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16)
  const out = bytes.slice()
  const n = Math.min(32, out.length)
  for (let i = 0; i < n; i++) out[i] ^= key[15 - (i % 16)]
  return out
}

export const FONT_TABLE_RELS_PATH = 'word/_rels/fontTable.xml.rels'

/** fontTable embed slots → de-obfuscated faces; unparseable parts are skipped */
export async function readEmbeddedFonts(
  zip: JSZip,
  fontTable: readonly FontTableEntry[],
): Promise<EmbeddedFont[]> {
  if (!fontTable.some((e) => e.embedded)) return []
  const rels = await parseRels(zip, FONT_TABLE_RELS_PATH)
  const out: EmbeddedFont[] = []
  for (const entry of fontTable) {
    for (const [slot, ref] of Object.entries(entry.embedded ?? {}) as Array<
      [EmbeddedFontSlot, EmbeddedFontRef]
    >) {
      const rel = rels.get(ref.rId)
      if (!rel || rel.targetMode === 'External') continue
      const path = resolveRelationshipTargetPath(FONT_TABLE_PART_PATH, rel.target)
      if (!path) continue
      const file = zip.file(path)
      if (!file) continue
      const data = deobfuscateOdttf(await file.async('uint8array'), ref.fontKey)
      if (!isSfnt(data)) continue
      const lineMetrics = sfntLineMetrics(data)
      out.push({
        family: entry.name,
        bold: slot === 'bold' || slot === 'boldItalic',
        italic: slot === 'italic' || slot === 'boldItalic',
        data,
        ...(lineMetrics ? { lineMetrics } : {}),
      })
    }
  }
  return out
}
