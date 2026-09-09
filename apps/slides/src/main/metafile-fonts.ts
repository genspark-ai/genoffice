/**
 * Facenames requested by a deck's EMF/WMF pictures (Excel/Word OLE previews). Metafile text is
 * drawn through canvas fonts, so Office-private faces (DFonts) must be resolved and registered
 * before the pictures rasterize — the deck's own runs never mention these families.
 */
import { gunzipSync } from 'node:zlib'
import type { PackageArchive } from '@genoffice/pptx-engine'

export interface MetafileFontRequest {
  family: string
  bold: boolean
  italic: boolean
}

const EMR_COMMENT = 70
const EMR_EXTCREATEFONTINDIRECTW = 82
const EMFPLUS_SIGNATURE = 0x2b464d45 // 'EMF+'
const EMFPLUS_OBJECT = 0x4008
const EMFPLUS_OBJECT_FONT = 6
const META_CREATEFONTINDIRECT = 0x02fb
const FONTSTYLE_BOLD = 1
const FONTSTYLE_ITALIC = 2
// eslint-disable-next-line no-control-regex
const CONTROL_RE = /[\u0000-\u001f\u007f]/g

function utf16(view: DataView, off: number, maxChars: number): string {
  let s = ''
  for (let i = 0; i < maxChars && off + i * 2 + 1 < view.byteLength; i++) {
    const c = view.getUint16(off + i * 2, true)
    if (c === 0) break
    s += String.fromCharCode(c)
  }
  return s
}

function add(
  out: Map<string, MetafileFontRequest>,
  family: string,
  bold: boolean,
  italic: boolean,
): void {
  const clean = family.replace(CONTROL_RE, '').trim()
  if (!clean) return
  const key = `${clean.toLowerCase()}|${bold ? 1 : 0}${italic ? 1 : 0}`
  if (!out.has(key)) out.set(key, { family: clean, bold, italic })
}

function scanEmfPlus(
  view: DataView,
  start: number,
  length: number,
  out: Map<string, MetafileFontRequest>,
): void {
  const end = Math.min(start + length, view.byteLength)
  let off = start
  while (off + 12 <= end) {
    const type = view.getUint16(off, true)
    const flags = view.getUint16(off + 2, true)
    const size = view.getUint32(off + 4, true)
    if (size < 12 || off + size > end) break
    if (type === EMFPLUS_OBJECT && ((flags >> 8) & 0x7f) === EMFPLUS_OBJECT_FONT && size >= 40) {
      // EmfPlusFont: Version(4) EmSize(4) SizeUnit(4) FontStyleFlags(4) Reserved(4) Length(4) FamilyName
      const data = off + 12
      const style = view.getInt32(data + 12, true)
      const nameLen = view.getUint32(data + 20, true)
      if (nameLen > 0 && data + 24 + nameLen * 2 <= off + size)
        add(
          out,
          utf16(view, data + 24, nameLen),
          (style & FONTSTYLE_BOLD) !== 0,
          (style & FONTSTYLE_ITALIC) !== 0,
        )
    }
    off += size
  }
}

function scanEmf(view: DataView, out: Map<string, MetafileFontRequest>): void {
  let off = 0
  while (off + 8 <= view.byteLength) {
    const type = view.getUint32(off, true)
    const size = view.getUint32(off + 4, true)
    if (size < 8 || off + size > view.byteLength) break
    if (type === EMR_EXTCREATEFONTINDIRECTW && size >= 12 + 92) {
      // ihFont(4), then LOGFONTW: Height Width Escapement Orientation Weight (5×4), Italic(1) …, FaceName[32] at +28
      const lf = off + 12
      const weight = view.getInt32(lf + 16, true)
      const italic = view.getUint8(lf + 20) !== 0
      add(out, utf16(view, lf + 28, 32), weight >= 600, italic)
    } else if (type === EMR_COMMENT && size >= 16) {
      const dataSize = view.getUint32(off + 8, true)
      if (dataSize >= 4 && view.getUint32(off + 12, true) === EMFPLUS_SIGNATURE)
        scanEmfPlus(view, off + 16, Math.min(dataSize - 4, size - 16), out)
    }
    off += size
  }
}

function scanWmf(view: DataView, out: Map<string, MetafileFontRequest>): void {
  // Optional placeable header (22 bytes), then META_HEADER (18 bytes)
  let off = (view.getUint32(0, true) === 0x9ac6cdd7 ? 22 : 0) + 18
  while (off + 6 <= view.byteLength) {
    const size = view.getUint32(off, true) * 2
    const fn = view.getUint16(off + 4, true)
    if (size < 6 || off + size > view.byteLength) break
    if (fn === META_CREATEFONTINDIRECT && size >= 6 + 18) {
      // LOGFONT16: Height Width Escapement Orientation Weight (5×2), Italic(1) …, FaceName at +18
      const lf = off + 6
      const weight = view.getInt16(lf + 8, true)
      const italic = view.getUint8(lf + 10) !== 0
      let name = ''
      for (let i = 0; i < 32 && lf + 18 + i < off + size; i++) {
        const c = view.getUint8(lf + 18 + i)
        if (c === 0) break
        name += String.fromCharCode(c)
      }
      add(out, name, weight >= 600, italic)
    }
    off += size
  }
}

/** Font requests found in one EMF/WMF (optionally gzip-compressed .emz/.wmz) payload. */
export function scanMetafileFonts(bytes: Uint8Array): MetafileFontRequest[] {
  let u8 = bytes
  if (u8.length > 2 && u8[0] === 0x1f && u8[1] === 0x8b) {
    try {
      u8 = new Uint8Array(gunzipSync(u8))
    } catch {
      return []
    }
  }
  if (u8.length < 18) return []
  const view = new DataView(u8.buffer, u8.byteOffset, u8.byteLength)
  const out = new Map<string, MetafileFontRequest>()
  try {
    const isEmf =
      view.getUint32(0, true) === 1 && u8.length >= 44 && view.getUint32(40, true) === 0x464d4520
    if (isEmf) scanEmf(view, out)
    else scanWmf(view, out)
  } catch {
    // truncated record tail: keep whatever was collected
  }
  return [...out.values()]
}

/** Every font request across the deck's metafile media parts. */
export function listMetafileFonts(archive: PackageArchive): MetafileFontRequest[] {
  const out = new Map<string, MetafileFontRequest>()
  for (const [path, bytes] of archive.entries) {
    if (!/^ppt\/media\/[^/]+\.(emf|wmf|emz|wmz)$/i.test(path)) continue
    for (const f of scanMetafileFonts(bytes))
      out.set(`${f.family.toLowerCase()}|${f.bold ? 1 : 0}${f.italic ? 1 : 0}`, f)
  }
  return [...out.values()]
}
