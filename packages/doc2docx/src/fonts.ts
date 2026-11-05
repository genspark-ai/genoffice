/**
 * The Word 97-2003 font table (SttbfFfn, FIB pair 15): an STTB whose strings
 * are FFN records. Each FFN starts with cbFfnM1 (its size minus one), then
 * ffid, wWeight, chs, ixchSzAlt, a 10-byte PANOSE and a 24-byte FontSignature,
 * and the UTF-16 font name (xszFfn, null-terminated) at offset 40. Character
 * runs refer to fonts by their index in this table (ftc). A malformed table
 * keeps the names read so far and records a warning.
 */
import type { FibInfo } from './fib'

const FFN_NAME_OFFSET = 40
const MAX_FONTS = 0x7fff

function u16At(bytes: Uint8Array, offset: number): number {
  return bytes[offset]! | (bytes[offset + 1]! << 8)
}

/** One SttbfFfn entry, the source of a w:font element in word/fontTable.xml */
export interface FontEntry {
  name: string
  altName?: string
  /** ffid family bits (FF_*): roman, swiss, modern, script, decorative or auto */
  family: 'roman' | 'swiss' | 'modern' | 'script' | 'decorative' | 'auto'
  /** ffid prq bits */
  pitch: 'default' | 'fixed' | 'variable'
  /** chs: Windows character set (w:charset, hex) */
  charset: number
  /** 10 PANOSE bytes */
  panose: Uint8Array
  /** FontSignature: fsUsb[4] then fsCsb[2] (w:sig) */
  signature: number[]
}

const FAMILIES: FontEntry['family'][] = ['auto', 'roman', 'swiss', 'modern', 'script', 'decorative']
const PITCHES: FontEntry['pitch'][] = ['default', 'fixed', 'variable']

function u32At(bytes: Uint8Array, offset: number): number {
  return (u16At(bytes, offset) | (u16At(bytes, offset + 2) << 16)) >>> 0
}

/** Font names by ftc; empty when the document has no font table. */
export function parseFontTable(fib: FibInfo, table: Uint8Array, warnings: string[]): string[] {
  return parseFontEntries(fib, table, warnings).map((entry) => entry.name)
}

/** The FFN records by ftc (name, alternative name, charset, PANOSE, family, pitch, signature). */
export function parseFontEntries(fib: FibInfo, table: Uint8Array, warnings: string[]): FontEntry[] {
  const { fcSttbfFfn: fc, lcbSttbfFfn: lcb } = fib
  if (lcb === 0) return []
  const entries: FontEntry[] = []
  const warn = (detail: string): void => {
    warnings.push(
      'the font table is invalid (' + detail + '); font names after that point are unknown',
    )
  }
  if (fc > table.length || lcb > table.length - fc || lcb < 4) {
    warn('its range lies outside the table stream')
    return entries
  }
  const end = fc + lcb
  const count = u16At(table, fc)
  const cbExtra = u16At(table, fc + 2)
  if (count > MAX_FONTS) {
    warn('it declares ' + count + ' fonts')
    return entries
  }
  let offset = fc + 4
  for (let index = 0; index < count; index++) {
    if (offset + 1 > end) {
      warn('font ' + index + ' is missing')
      break
    }
    const size = table[offset]! + 1
    if (offset + size > end) {
      warn('font ' + index + ' overruns the table')
      break
    }
    const strings: string[] = []
    let name = ''
    for (let at = offset + FFN_NAME_OFFSET; at + 2 <= offset + size; at += 2) {
      const code = u16At(table, at)
      if (code === 0) {
        strings.push(name)
        name = ''
        continue
      }
      name += String.fromCharCode(code)
    }
    if (name !== '') strings.push(name)
    const ffid = size > 1 ? table[offset + 1]! : 0
    const ixchSzAlt = size > 5 ? table[offset + 5]! : 0
    const altName = ixchSzAlt > 0 ? strings[1] : undefined
    const header = size >= FFN_NAME_OFFSET
    entries.push({
      name: strings[0] ?? '',
      ...(altName ? { altName } : {}),
      family: FAMILIES[(ffid >> 4) & 0x07] ?? 'auto',
      pitch: PITCHES[ffid & 0x03] ?? 'default',
      charset: size > 4 ? table[offset + 4]! : 0,
      panose: header ? table.slice(offset + 6, offset + 16) : new Uint8Array(10),
      signature: header
        ? [0, 4, 8, 12, 16, 20].map((at) => u32At(table, offset + 16 + at))
        : [0, 0, 0, 0, 0, 0],
    })
    offset += size + cbExtra
  }
  return entries
}

/** The resolved font names of a run (w:rFonts ascii/hAnsi, eastAsia, cs). */
export interface FontNames {
  ascii?: string
  eastAsia?: string
  cs?: string
}

/** Default ftc values (STSHI rgftcStandardChpStsh) used when a run names no font. */
export interface DefaultFontIndices {
  ascii?: number
  eastAsia?: number
  other?: number
  cs?: number
}

/** The font table plus the stylesheet defaults the resolver falls back to. */
export interface FontTable {
  names: string[]
  defaults: DefaultFontIndices
}

/** Map the ftc indices of a run to names; undefined slots stay unset. */
export function fontNamesFor(
  fonts: FontTable,
  ftc: { ascii?: number; hAnsi?: number; eastAsia?: number; cs?: number },
): FontNames | undefined {
  const nameOf = (index: number | undefined): string | undefined => {
    if (index === undefined) return undefined
    const name = fonts.names[index]
    return name === undefined || name === '' ? undefined : name
  }
  const ascii = nameOf(ftc.ascii ?? ftc.hAnsi ?? fonts.defaults.ascii)
  const eastAsia = nameOf(ftc.eastAsia ?? fonts.defaults.eastAsia)
  const cs = nameOf(ftc.cs ?? fonts.defaults.cs)
  if (ascii === undefined && eastAsia === undefined && cs === undefined) return undefined
  return {
    ...(ascii !== undefined ? { ascii } : {}),
    ...(eastAsia !== undefined ? { eastAsia } : {}),
    ...(cs !== undefined ? { cs } : {}),
  }
}
