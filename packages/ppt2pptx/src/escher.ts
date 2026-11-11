/** Escher (MS-ODRAW) reading for the PPT stage-3 surfaces: shape geometry,
 *  the picture store, and the master's drawing. Records use the same 8-byte
 *  header as MS-PPT but live inside PPDrawing; this reader is deliberately
 *  tolerant (a malformed child ends that shape's parse, not the conversion). */

import { inflateSync } from 'node:zlib'
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
  clientData: 0xf011,
  secondaryOpt: 0xf121,
  tertiaryOpt: 0xf122,
} as const

export const OPT = {
  /** MS-ODRAW 2.3.18.5: opid MUST be 0x0004, FixedPoint 16.16 degrees */
  rotation: 0x0004,
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
  /** clockwise degrees from OPT.rotation */
  rotation: number
  /** OEPlaceholderAtom.placeholderId (0 = none), or null */
  placeholderId: number | null
  /** OfficeArtFSPGR coordinate space of a group shape */
  groupBounds: EscherRect | null
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
  ext: 'png' | 'jpg' | 'emf' | 'wmf'
  contentType: 'image/png' | 'image/jpeg' | 'image/x-emf' | 'image/x-wmf'
}
const PLACEHOLDER_ATOM = 0x0bc3

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
    rotation: 0,
    placeholderId: null,
    groupBounds: null,
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
    } else if (child.type === ESCHER.fspgr) {
      shape.groupBounds = rectOf(bytes, child)
    } else if (child.type === ESCHER.clientData) {
      for (const atom of children(bytes, child.start, child.end)) {
        if (atom.type === PLACEHOLDER_ATOM && atom.end - atom.start >= 5) {
          shape.placeholderId = bytes[atom.start + 4] ?? 0
        }
      }
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
        if (pid === OPT.rotation) shape.rotation = (value | 0) / 65536
        if (pid === OPT.pib) shape.blipIndex = value
      }
    } else if (child.type === ESCHER.clientTextbox) {
      shape.text = shapeTextOf(bytes, child)
    }
  }
  return shape
}

/** Every leaf SpContainer under a drawing, with its groups flattened away. */
function mapRect(child: EscherRect, bounds: EscherRect, outer: EscherRect): EscherRect {
  const boundsWidth = bounds.right - bounds.left
  const boundsHeight = bounds.bottom - bounds.top
  if (boundsWidth <= 0 || boundsHeight <= 0) return child
  const scaleX = (outer.right - outer.left) / boundsWidth
  const scaleY = (outer.bottom - outer.top) / boundsHeight
  return {
    left: Math.round(outer.left + (child.left - bounds.left) * scaleX),
    top: Math.round(outer.top + (child.top - bounds.top) * scaleY),
    right: Math.round(outer.left + (child.right - bounds.left) * scaleX),
    bottom: Math.round(outer.top + (child.bottom - bounds.top) * scaleY),
  }
}

/** Every leaf SpContainer under a drawing. A group's first shape is the group
 *  itself: its FSPGR bounds describe the child coordinate space and its anchor
 *  the outer rectangle, so member anchors are mapped into slide coordinates. */
export function shapesOf(bytes: Uint8Array, from: number, to: number): PptShape[] {
  const out: PptShape[] = []
  const emit = (
    record: EscherRecord,
    bounds: EscherRect | null,
    outer: EscherRect | null,
  ): void => {
    if (out.length >= MAX_SHAPES) return
    const shape = parseShape(bytes, record)
    if (shape.rect && bounds && outer) shape.rect = mapRect(shape.rect, bounds, outer)
    if ((shape.flags & (FSP.group | FSP.patriarch | FSP.deleted)) === 0) out.push(shape)
  }
  const walk = (start: number, end: number, depth: number): void => {
    if (depth > MAX_DEPTH) return
    for (const record of children(bytes, start, end)) {
      if (record.type === ESCHER.spgrContainer) {
        const nested = [...children(bytes, record.start, record.end)]
        const members = nested.filter((r) => r.type === ESCHER.spContainer)
        const [first, ...rest] = members
        const group = first ? parseShape(bytes, first) : null
        // A real group's first shape is the group itself (FSP group/patriarch
        // flag); without that flag every SpContainer is an ordinary shape.
        const isGroup = group !== null && (group.flags & (FSP.group | FSP.patriarch)) !== 0
        const bounds = isGroup ? group.groupBounds : null
        const outer = isGroup ? group.rect : null
        const childrenShapes = isGroup ? rest : members
        for (const member of childrenShapes) emit(member, bounds, outer)
        for (const child of nested.filter((r) => r.type === ESCHER.spgrContainer)) {
          walk(child.start, child.end, depth + 1)
        }
        continue
      }
      if (record.type === ESCHER.spContainer) {
        emit(record, null, null)
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
const BLIP_EMF = 0xf01a
const BLIP_WMF = 0xf01b
const TWO_UID_INSTANCES: Record<number, number[]> = {
  [BLIP_PNG]: [0x6e1],
  0xf01d: [0x46b, 0x6e3],
  0xf02a: [0x46b, 0x6e3],
}

/** Decode one embedded OfficeArtBlip: UID bytes, then a tag byte, then data. */
function readBlip(bytes: Uint8Array, offset: number, limit: number): PptImage | null {
  const record = recordAt(bytes, offset, limit)
  if (!record) return null
  if (record.type === BLIP_EMF || record.type === BLIP_WMF) {
    // OfficeArtBlipEMF/WMF: UID (32 bytes for the two-UID instances), then a
    // 34-byte metafile header whose cbSave names the decompressed size.
    const twoUid =
      (record.type === BLIP_EMF && record.instance === 0x3d5) ||
      (record.type === BLIP_WMF && record.instance === 0x217)
    let dataStart = record.start + (twoUid ? 32 : 16)
    if (dataStart + 34 > record.end) return null
    const dv = viewOf(bytes)
    const cbSave = dv.getUint32(dataStart + 28, true)
    const compression = bytes[dataStart + 32] ?? 0
    dataStart += 34
    const available = Math.min(cbSave, record.end - dataStart)
    let payload = bytes.slice(dataStart, dataStart + available)
    if (compression === 0) {
      try {
        payload = new Uint8Array(inflateSync(payload))
      } catch {
        return null
      }
    }
    if (payload.length === 0) return null
    return record.type === BLIP_EMF
      ? { bytes: payload, ext: 'emf', contentType: 'image/x-emf' }
      : { bytes: payload, ext: 'wmf', contentType: 'image/x-wmf' }
  }
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
