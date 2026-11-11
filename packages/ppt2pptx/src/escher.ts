/** Escher (MS-ODRAW) reading for the PPT stage-3 surfaces: shape geometry,
 *  the picture store, and the master's drawing. Records use the same 8-byte
 *  header as MS-PPT but live inside PPDrawing; this reader is deliberately
 *  tolerant (a malformed child ends that shape's parse, not the conversion). */

import type { TextStyle } from './text-style'
import { parseStyleTextProp } from './text-style'
import { decodeTextBytes, decodeTextChars, PPT_TYPES } from './records'

export const ESCHER = {
  dggContainer: 0xf000,
  bstoreContainer: 0xf001,
  dgContainer: 0xf002,
  spgrContainer: 0xf003,
  spContainer: 0xf004,
  fbse: 0xf007,
  fspgr: 0xf009,
  fsp: 0xf00a,
  opt: 0xf00b,
  clientTextbox: 0xf00d,
  childAnchor: 0xf00f,
  clientAnchor: 0xf010,
  secondaryOpt: 0xf121,
  tertiaryOpt: 0xf122,
} as const

export const OPT = {
  pib: 0x0104,
  fillColor: 0x0181,
  lineColor: 0x01c0,
  lineWidth: 0x01cb,
} as const

/** OfficeArtFSP flags (MS-ODRAW 2.2.40): group/patriarch/deleted plus flips. */
export const FSP = { group: 0x1, patriarch: 0x4, deleted: 0x8, flipH: 0x40, flipV: 0x80 } as const

/** shape types from MS-ODRAW msospt*, the subset DrawingML has a preset for */
export const SHAPE_PRESETS: Record<number, string> = {
  1: 'rect',
  2: 'roundRect',
  3: 'ellipse',
  4: 'diamond',
  5: 'triangle',
  9: 'line',
  75: 'rect', // msosptPicture is drawn as a <p:pic>, not a preset
  202: 'rect', // msosptTextBox
}

export interface EscherRect {
  left: number
  top: number
  right: number
  bottom: number
}

export interface PptShapeText {
  kind: number
  text: string
  style: TextStyle
  align: 'l' | 'ctr' | 'r' | 'just' | null
}

export interface PptShape {
  spid: number
  shapeType: number
  flags: number
  /** master units (576/inch); null means the shape had no anchor */
  rect: EscherRect | null
  props: Map<number, number>
  text: PptShapeText | null
  /** 1-based blip index from OPT_PIB */
  blipIndex: number | null
}

export interface PptImage {
  bytes: Uint8Array
  ext: 'png' | 'jpg'
  contentType: 'image/png' | 'image/jpeg'
}

export interface PptMaster {
  shapes: PptShape[]
}

interface EscherRecord {
  version: number
  instance: number
  type: number
  start: number
  end: number
}

const MAX_DEPTH = 32
const MAX_SHAPES = 65536

function viewOf(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

/** Tolerant record header: null at a boundary instead of throwing. */
function recordAt(bytes: Uint8Array, offset: number, limit: number): EscherRecord | null {
  if (offset < 0 || offset + 8 > limit || offset + 8 > bytes.length) return null
  const dv = viewOf(bytes)
  const head = dv.getUint16(offset, true)
  const type = dv.getUint16(offset + 2, true)
  const length = dv.getUint32(offset + 4, true)
  const start = offset + 8
  if (start + length > bytes.length) return null
  return {
    version: head & 0xf,
    instance: head >> 4,
    type,
    start,
    end: Math.min(start + length, limit),
  }
}

function* children(bytes: Uint8Array, from: number, to: number): Generator<EscherRecord> {
  let at = from
  while (at + 8 <= to) {
    const record = recordAt(bytes, at, to)
    if (!record) return
    yield record
    at = record.end
  }
}

function rectOf(bytes: Uint8Array, record: EscherRecord): EscherRect | null {
  if (record.end - record.start < 16) return null
  const dv = viewOf(bytes)
  const rect: EscherRect = {
    left: dv.getInt32(record.start, true),
    top: dv.getInt32(record.start + 4, true),
    right: dv.getInt32(record.start + 8, true),
    bottom: dv.getInt32(record.start + 12, true),
  }
  if (rect.right <= rect.left || rect.bottom <= rect.top) return null
  return rect
}

/** Parse the OfficeArtClientTextbox's MS-PPT children, including its styling. */
function shapeTextOf(bytes: Uint8Array, record: EscherRecord): PptShapeText | null {
  let kind = 4
  let text: string | null = null
  let charCount = 0
  let stylePayload: Uint8Array | null = null
  let styleCount = 0
  for (const inner of children(bytes, record.start, record.end)) {
    if (inner.type === PPT_TYPES.textHeaderAtom) {
      if (inner.end - inner.start >= 4) kind = viewOf(bytes).getInt32(inner.start, true)
      continue
    }
    if (inner.type === PPT_TYPES.textCharsAtom || inner.type === PPT_TYPES.textBytesAtom) {
      const payload = bytes.subarray(inner.start, inner.end)
      text =
        inner.type === PPT_TYPES.textCharsAtom ? decodeTextChars(payload) : decodeTextBytes(payload)
      charCount = inner.type === PPT_TYPES.textCharsAtom ? payload.length / 2 : payload.length
      continue
    }
    if (inner.type === PPT_TYPES.styleTextPropAtom) {
      stylePayload = bytes.subarray(inner.start, inner.end)
      styleCount = inner.instance
    }
  }
  if (text === null || text.trim().length === 0) return null
  const count = styleCount > 0 ? styleCount : Math.round(charCount) || text.length
  const style = stylePayload
    ? parseStyleTextProp(stylePayload, count)
    : { paragraphs: [], characters: [] }
  return { kind, text, style, align: style.paragraphs[0]?.align ?? null }
}

function parseShape(bytes: Uint8Array, record: EscherRecord): PptShape {
  const shape: PptShape = {
    spid: 0,
    shapeType: 0,
    flags: 0,
    rect: null,
    props: new Map(),
    text: null,
    blipIndex: null,
  }
  const dv = viewOf(bytes)
  for (const child of children(bytes, record.start, record.end)) {
    if (child.type === ESCHER.fsp && child.end - child.start >= 8) {
      shape.spid = dv.getUint32(child.start, true)
      shape.flags = dv.getUint32(child.start + 4, true)
      shape.shapeType = child.instance
    } else if (child.type === ESCHER.childAnchor || child.type === ESCHER.clientAnchor) {
      shape.rect = rectOf(bytes, child) ?? shape.rect
    } else if (
      child.type === ESCHER.opt ||
      child.type === ESCHER.secondaryOpt ||
      child.type === ESCHER.tertiaryOpt
    ) {
      for (let index = 0; index < child.instance; index++) {
        const at = child.start + index * 6
        if (at + 6 > child.end) break
        const entry = dv.getUint16(at, true)
        // complex entries carry a length instead of a value and their data
        // follows the table; skip them rather than desynchronising the walk
        if ((entry & 0x4000) !== 0) continue
        const pid = entry & 0x3fff
        const value = dv.getUint32(at + 2, true)
        shape.props.set(pid, value)
        if (pid === OPT.pib) shape.blipIndex = value
      }
    } else if (child.type === ESCHER.clientTextbox) {
      shape.text = shapeTextOf(bytes, child)
    }
  }
  return shape
}

/** Every leaf SpContainer under a drawing, with its groups flattened away. */
export function shapesOf(bytes: Uint8Array, from: number, to: number): PptShape[] {
  const out: PptShape[] = []
  const walk = (start: number, end: number, depth: number): void => {
    if (depth > MAX_DEPTH || out.length >= MAX_SHAPES) return
    for (const record of children(bytes, start, end)) {
      if (record.type === ESCHER.spContainer) {
        const shape = parseShape(bytes, record)
        if ((shape.flags & (FSP.group | FSP.patriarch | FSP.deleted)) === 0) out.push(shape)
        continue
      }
      if (record.version === 0xf) walk(record.start, record.end, depth + 1)
    }
  }
  walk(from, to, 0)
  return out
}

/* OfficeArtBlip record types (MS-ODRAW 2.2.23-2.2.30) */
const BLIP_PNG = 0xf01e
const BLIP_JPEG = new Set([0xf01d, 0xf02a])
const TWO_UID_INSTANCES: Record<number, number[]> = {
  [BLIP_PNG]: [0x6e1],
  0xf01d: [0x46b, 0x6e3],
  0xf02a: [0x46b, 0x6e3],
}

/** Decode one embedded OfficeArtBlip: UID bytes, then a tag byte, then data. */
function readBlip(bytes: Uint8Array, offset: number, limit: number): PptImage | null {
  const record = recordAt(bytes, offset, limit)
  if (!record) return null
  let image: PptImage
  if (record.type === BLIP_PNG) {
    image = { bytes: new Uint8Array(0), ext: 'png', contentType: 'image/png' }
  } else if (BLIP_JPEG.has(record.type)) {
    image = { bytes: new Uint8Array(0), ext: 'jpg', contentType: 'image/jpeg' }
  } else {
    return null // metafiles, DIB, TIFF and PICT are not carried by stage 3
  }
  const twoUid = TWO_UID_INSTANCES[record.type]?.includes(record.instance) ?? false
  const start = record.start + (twoUid ? 32 : 16) + 1
  if (start >= record.end) return null
  const data = bytes.slice(start, record.end)
  // sanity: a PNG signature or a JPEG SOI marker, so a garbage blip is dropped
  if (image.ext === 'png' && !(data[0] === 0x89 && data[1] === 0x50)) return null
  if (image.ext === 'jpg' && !(data[0] === 0xff && data[1] === 0xd8)) return null
  return { ...image, bytes: data }
}

/** The embedded blips of the global BStore, in shape-reference order (1-based). */
export function imagesOf(
  document: Uint8Array,
  pictures: Uint8Array | null,
): Array<PptImage | null> {
  const out: Array<PptImage | null> = []
  const scan = (from: number, to: number, depth: number): void => {
    if (depth > MAX_DEPTH) return
    for (const record of children(document, from, to)) {
      if (record.type === ESCHER.bstoreContainer) {
        for (const entry of children(document, record.start, record.end)) {
          if (entry.type === ESCHER.fbse) {
            // FBSE: 36-byte header, cbName at +33, then an optional embedded
            // blip; foDelay at +28 points into the Pictures stream otherwise
            if (entry.end - entry.start < 36) {
              out.push(null)
              continue
            }
            const dv = viewOf(document)
            const nameBytes = document[entry.start + 33] ?? 0
            const embedded = entry.start + 36 + nameBytes
            const fromBlip = readBlip(document, embedded, entry.end)
            if (fromBlip) {
              out.push(fromBlip)
            } else {
              let fallback: PptImage | null = null
              if (pictures) {
                const delay = dv.getUint32(entry.start + 28, true)
                fallback = readBlip(pictures, delay, pictures.length)
              }
              out.push(fallback)
            }
          } else {
            out.push(readBlip(document, entry.start - 8, entry.end))
          }
        }
        continue
      }
      if (record.version === 0xf) scan(record.start, record.end, depth + 1)
    }
  }
  scan(0, document.length, 0)
  return out
}

/** The main master's drawing, so its shapes can be inherited by every slide. */
export function masterOf(document: Uint8Array, from: number, to: number): PptMaster | null {
  for (const record of children(document, from, to)) {
    if (record.type === PPT_TYPES.mainMasterContainer) {
      return { shapes: shapesOf(document, record.start, record.end) }
    }
  }
  return null
}
