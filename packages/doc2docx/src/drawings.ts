/**
 * Floating pictures (OfficeArt shapes anchored in the text).
 *
 * PlcSpaMom (FIB pair 40) and PlcSpaHdr (pair 41) list the FSPA anchors of
 * the body and the header document: the CP of the \x08 anchor character, the
 * shape id and the shape rectangle in twips with its relative-to / wrapping
 * flags. DggInfo (pair 50) holds the OfficeArt data: the blip store (FBSE
 * records pointing at picture data in the WordDocument stream) and one
 * drawing container per story whose shape containers name their picture via
 * the pib property. Shapes whose id matches a text box story become text
 * boxes; other non-picture shapes (autoshapes, lines) are skipped with one
 * warning. Groups contribute their picture and text box children, placed
 * through the group's coordinate space.
 */
import { inflateSync } from 'node:zlib'
import type { FibInfo } from './fib'
import type { HeaderFooterStory } from './headers'
import type { TextBoxStories, TextBoxStory } from './textboxes'

export type PictureMediaType =
  'image/png' | 'image/jpeg' | 'image/x-wmf' | 'image/x-emf' | 'image/bmp' | 'image/tiff'

export type PictureExtension = 'png' | 'jpeg' | 'wmf' | 'emf' | 'bmp' | 'tiff'

/** cropped-away part of each side of the source picture, in 1/1000 of a percent (a:srcRect) */
export interface PictureCrop {
  l: number
  t: number
  r: number
  b: number
}

export interface PictureMedia {
  mediaType: PictureMediaType
  extension: PictureExtension
  bytes: Uint8Array
}

export const PICTURE_EXTENSIONS: Record<PictureMediaType, PictureExtension> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/x-wmf': 'wmf',
  'image/x-emf': 'emf',
  'image/bmp': 'bmp',
  'image/tiff': 'tiff',
}

interface FloatingShapeBase {
  /** CP of the \x08 anchor in its story (header CPs are relative to the header document) */
  cp: number
  story: 'body' | 'header'
  spid: number
  leftTwips: number
  topTwips: number
  widthTwips: number
  heightTwips: number
  horizontalRelative: 'margin' | 'page' | 'column' | 'character'
  verticalRelative: 'margin' | 'page' | 'paragraph' | 'line'
  wrap: 'square' | 'topAndBottom' | 'none' | 'tight' | 'through'
  behindText: boolean
}

export interface FloatingPicture extends FloatingShapeBase {
  kind: 'picture'
  media: PictureMedia
  crop?: PictureCrop
}

export interface FloatingTextBox extends FloatingShapeBase {
  kind: 'textBox'
  geometry: 'rect' | 'roundRect' | 'ellipse'
  /** RRGGBB fill; undefined = no fill */
  fill?: string
  /** outline; undefined = no line */
  line?: { color: string; widthEmu: number }
  /** left, top, right, bottom text insets in EMU */
  insetsEmu: [number, number, number, number]
  verticalAnchor: 't' | 'ctr' | 'b'
  /** rotated text (msotxflBtoT -> vert270, the top-to-bottom flows -> vert); undefined = horizontal */
  textDirection?: 'vert' | 'vert270'
  autoFit: boolean
  content: HeaderFooterStory
}

export type FloatingShape = FloatingPicture | FloatingTextBox

export interface FloatingPicturesResult {
  body: FloatingShape[]
  header: FloatingShape[]
  warnings: string[]
}

export interface DrawingSource {
  fib: FibInfo
  wordDocument: Uint8Array
  table: Uint8Array
}

const FSPA_BYTES = 26
const MAX_SHAPES = 65536

/* OfficeArt record types */
const REC_DGG_CONTAINER = 0xf000
const REC_BSTORE_CONTAINER = 0xf001
const REC_DG_CONTAINER = 0xf002
const REC_SPGR_CONTAINER = 0xf003
const REC_SP_CONTAINER = 0xf004
const REC_FBSE = 0xf007
const REC_FSPGR = 0xf009
const REC_FSP = 0xf00a
const REC_OPT = 0xf00b
const REC_CHILD_ANCHOR = 0xf00f
const REC_SECONDARY_OPT = 0xf121
const REC_TERTIARY_OPT = 0xf122
const FBSE_BYTES = 36
const OPT_PIB = 0x0104
/** cropFromTop/Bottom/Left/Right: 16.16 fixed-point fractions of the source picture */
const OPT_CROP_FROM_TOP = 0x0100
const OPT_CROP_FROM_BOTTOM = 0x0101
const OPT_CROP_FROM_LEFT = 0x0102
const OPT_CROP_FROM_RIGHT = 0x0103
const OPT_TXID = 0x0080
const OPT_DX_TEXT_LEFT = 0x0081
const OPT_DY_TEXT_TOP = 0x0082
const OPT_DX_TEXT_RIGHT = 0x0083
const OPT_DY_TEXT_BOTTOM = 0x0084
const OPT_ANCHOR_TEXT = 0x0087
const OPT_TEXT_FLOW = 0x0088
const OPT_TEXT_BOOLEANS = 0x00bf
const OPT_FILL_COLOR = 0x0181
const OPT_FILL_BOOLEANS = 0x01bf
const OPT_LINE_COLOR = 0x01c0
const OPT_LINE_WIDTH = 0x01cb
const OPT_LINE_BOOLEANS = 0x01ff
const OPT_POS_REL_H = 0x0390
const OPT_POS_REL_V = 0x0392
const DEFAULT_LINE_WIDTH_EMU = 9525
const DEFAULT_INSETS_EMU: [number, number, number, number] = [91440, 45720, 91440, 45720]
/** msosptRoundRectangle / msosptEllipse; everything else (incl. msosptTextBox 202) is a rectangle */
const SHAPE_GEOMETRY: Record<number, FloatingTextBox['geometry']> = { 2: 'roundRect', 3: 'ellipse' }
const FSP_GROUP = 0x0001
const FSP_PATRIARCH = 0x0004
const FSP_DELETED = 0x0008

interface BlipKind {
  mediaType: PictureMediaType | undefined
  metafile: boolean
  twoUidInstances: number[]
}

/** OfficeArtBlip record types (MS-ODRAW 2.2.23-2.2.30) */
const BLIP_KINDS: Record<number, BlipKind> = {
  0xf01a: { mediaType: 'image/x-emf', metafile: true, twoUidInstances: [0x3d5] },
  0xf01b: { mediaType: 'image/x-wmf', metafile: true, twoUidInstances: [0x217] },
  0xf01c: { mediaType: undefined, metafile: true, twoUidInstances: [0x543] },
  0xf01d: { mediaType: 'image/jpeg', metafile: false, twoUidInstances: [0x46b, 0x6e3] },
  0xf01e: { mediaType: 'image/png', metafile: false, twoUidInstances: [0x6e1] },
  0xf01f: { mediaType: 'image/bmp', metafile: false, twoUidInstances: [0x7a9] },
  0xf029: { mediaType: 'image/tiff', metafile: false, twoUidInstances: [0x6e5] },
  0xf02a: { mediaType: 'image/jpeg', metafile: false, twoUidInstances: [0x46b, 0x6e3] },
}

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

interface RecordHeader {
  version: number
  instance: number
  type: number
  length: number
  /** offset of the record payload */
  start: number
  /** offset one past the payload (clamped to the buffer) */
  end: number
}

function readRecord(bytes: Uint8Array, offset: number, limit: number): RecordHeader | undefined {
  if (offset + 8 > limit) return undefined
  const dv = view(bytes)
  const verInstance = dv.getUint16(offset, true)
  const type = dv.getUint16(offset + 2, true)
  const length = dv.getUint32(offset + 4, true)
  return {
    version: verInstance & 0x0f,
    instance: verInstance >> 4,
    type,
    length,
    start: offset + 8,
    end: Math.min(offset + 8 + length, limit),
  }
}

function* records(bytes: Uint8Array, start: number, limit: number): Generator<RecordHeader> {
  let offset = start
  while (offset + 8 <= limit) {
    const record = readRecord(bytes, offset, limit)
    if (!record) return
    yield record
    offset = record.start + record.length
  }
}

/** Is there an OfficeArt blip record header at this offset? */
export function isBlipRecordAt(bytes: Uint8Array, offset: number): boolean {
  if (offset + 8 > bytes.length) return false
  const dv = view(bytes)
  const type = dv.getUint16(offset + 2, true)
  return BLIP_KINDS[type] !== undefined && (dv.getUint16(offset, true) & 0x0f) === 0
}

/** a DIB blip is a BITMAPINFOHEADER + pixels; a 14-byte file header makes it a .bmp */
function dibToBmp(dib: Uint8Array): Uint8Array | undefined {
  if (dib.length < 40) return undefined
  const dv = view(dib)
  const headerSize = dv.getUint32(0, true)
  const bitCount = dv.getUint16(14, true)
  const compression = dv.getUint32(16, true)
  const colorsUsed = dv.getUint32(32, true)
  let colorTable = 0
  if (bitCount <= 8) colorTable = (colorsUsed || 1 << bitCount) * 4
  if (compression === 3 && headerSize === 40) colorTable += 12
  const out = new Uint8Array(14 + dib.length)
  out[0] = 0x42
  out[1] = 0x4d
  const outView = view(out)
  outView.setUint32(2, out.length, true)
  outView.setUint32(10, 14 + headerSize + colorTable, true)
  out.set(dib, 14)
  return out
}

/**
 * Decode one OfficeArtBlip record. Returns undefined (with a warning) for
 * unsupported kinds (PICT) and corrupt or truncated data.
 */
export function readBlip(
  bytes: Uint8Array,
  offset: number,
  warnings: string[],
): PictureMedia | undefined {
  const record = readRecord(bytes, offset, bytes.length)
  if (!record) {
    warnings.push('a picture record lies outside its stream and was omitted')
    return undefined
  }
  const kind = BLIP_KINDS[record.type]
  if (!kind) {
    warnings.push('a picture uses an unknown OfficeArt record and was omitted')
    return undefined
  }
  if (!kind.mediaType) {
    warnings.push('a picture uses an unsupported format (PICT) and was omitted')
    return undefined
  }
  const truncated = record.start + record.length > bytes.length
  const uidBytes = kind.twoUidInstances.includes(record.instance) ? 32 : 16
  let dataStart = record.start + uidBytes
  let payload: Uint8Array
  if (kind.metafile) {
    if (dataStart + 34 > record.end) {
      warnings.push('a metafile picture has a truncated header and was omitted')
      return undefined
    }
    const dv = view(bytes)
    const cbSave = dv.getUint32(dataStart + 28, true)
    const compression = bytes[dataStart + 32]!
    dataStart += 34
    const available = Math.min(cbSave, record.end - dataStart)
    const raw = bytes.slice(dataStart, dataStart + available)
    if (compression === 0) {
      try {
        payload = new Uint8Array(inflateSync(raw))
      } catch {
        warnings.push('a compressed metafile picture could not be decompressed and was omitted')
        return undefined
      }
    } else payload = raw
    if (payload.length === 0) {
      warnings.push('a metafile picture is empty and was omitted')
      return undefined
    }
  } else {
    dataStart += 1
    if (dataStart >= record.end) {
      warnings.push('a bitmap picture has no image data and was omitted')
      return undefined
    }
    payload = bytes.slice(dataStart, record.end)
    if (truncated) {
      // a cut JPEG still decodes up to the break; other bitmaps do not
      if (kind.mediaType !== 'image/jpeg') {
        warnings.push('a picture record is truncated and was omitted')
        return undefined
      }
      warnings.push('a JPEG picture is truncated; the available data was kept')
    }
    if (kind.mediaType === 'image/bmp') {
      const bmp = dibToBmp(payload)
      if (!bmp) {
        warnings.push('a DIB picture has a truncated bitmap header and was omitted')
        return undefined
      }
      payload = bmp
    }
  }
  return {
    mediaType: kind.mediaType,
    extension: PICTURE_EXTENSIONS[kind.mediaType],
    bytes: payload,
  }
}

interface ShapeAnchor {
  cp: number
  spid: number
  left: number
  top: number
  right: number
  bottom: number
  bx: number
  by: number
  wr: number
  belowText: boolean
}

function parsePlcSpa(
  table: Uint8Array,
  fc: number,
  lcb: number,
  what: string,
  warnings: string[],
): ShapeAnchor[] {
  if (lcb === 0) return []
  if (
    fc > table.length ||
    lcb > table.length - fc ||
    lcb < 4 ||
    (lcb - 4) % (4 + FSPA_BYTES) !== 0
  ) {
    warnings.push('the ' + what + ' shape table is invalid; its floating shapes were omitted')
    return []
  }
  const count = (lcb - 4) / (4 + FSPA_BYTES)
  if (count > MAX_SHAPES) {
    warnings.push('the ' + what + ' shape table declares too many shapes; they were omitted')
    return []
  }
  const dv = view(table)
  const anchors: ShapeAnchor[] = []
  for (let index = 0; index < count; index++) {
    const cp = dv.getUint32(fc + index * 4, true)
    const base = fc + (count + 1) * 4 + index * FSPA_BYTES
    const flags = dv.getUint16(base + 20, true)
    anchors.push({
      cp,
      spid: dv.getUint32(base, true),
      left: dv.getInt32(base + 4, true),
      top: dv.getInt32(base + 8, true),
      right: dv.getInt32(base + 12, true),
      bottom: dv.getInt32(base + 16, true),
      bx: (flags >> 1) & 0x03,
      by: (flags >> 3) & 0x03,
      wr: (flags >> 5) & 0x0f,
      belowText: (flags & 0x4000) !== 0,
    })
  }
  return anchors
}

interface Rect {
  left: number
  top: number
  right: number
  bottom: number
}

interface Shape {
  spid: number
  flags: number
  /** FSP instance: the shape type */
  shapeType: number
  /** OfficeArt property table (pid -> value) */
  props: Map<number, number>
  pib?: number
  childAnchor?: Rect
  groupBounds?: Rect
  children: Shape[]
}

interface BlipStoreEntry {
  /** stream the blip lives in and the record offset there */
  source: 'table' | 'wordDocument'
  offset: number
}

interface OfficeArt {
  blips: BlipStoreEntry[]
  /** dgglbl (0 body, 1 header) -> top-level shapes of that drawing */
  drawings: Map<number, Shape[]>
}

function readRect(bytes: Uint8Array, offset: number): Rect {
  const dv = view(bytes)
  return {
    left: dv.getInt32(offset, true),
    top: dv.getInt32(offset + 4, true),
    right: dv.getInt32(offset + 8, true),
    bottom: dv.getInt32(offset + 12, true),
  }
}

/** a 16.16 fixed fraction -> 1/1000 percent; cropping away everything or more is ignored */
/** 16.16 fixed-point fraction -> thousandths of a percent; negative sides (padding) are kept, as Word does */
function cropFraction(value: number | undefined): number {
  if (value === undefined) return 0
  const signed = value | 0
  if (signed <= -0x10000 || signed >= 0x10000) return 0
  return Math.round((signed * 100000) / 0x10000)
}

function shapeCrop(props: Map<number, number>): PictureCrop | undefined {
  const crop: PictureCrop = {
    l: cropFraction(props.get(OPT_CROP_FROM_LEFT)),
    t: cropFraction(props.get(OPT_CROP_FROM_TOP)),
    r: cropFraction(props.get(OPT_CROP_FROM_RIGHT)),
    b: cropFraction(props.get(OPT_CROP_FROM_BOTTOM)),
  }
  if (crop.l + crop.r >= 100000 || crop.t + crop.b >= 100000) return undefined
  return crop.l || crop.t || crop.r || crop.b ? crop : undefined
}

/**
 * Crop of the inline picture whose OfficeArt shape container starts at or
 * shortly after `offset` (a PICF may carry a picture name before it). Word
 * stores inline crops in the shape's OPT, not in the PICF twip fields.
 */
export function inlineShapeCrop(
  bytes: Uint8Array,
  offset: number,
  limit: number,
): PictureCrop | undefined {
  const dv = view(bytes)
  for (let at = offset; at + 8 <= Math.min(limit, offset + 512); at++) {
    if (dv.getUint16(at + 2, true) !== REC_SP_CONTAINER || (dv.getUint16(at, true) & 0x0f) !== 0x0f)
      continue
    const record = readRecord(bytes, at, limit)
    if (!record) return undefined
    return shapeCrop(parseSpContainer(bytes, record).props)
  }
  return undefined
}

function parseSpContainer(bytes: Uint8Array, record: RecordHeader): Shape {
  const shape: Shape = { spid: 0, flags: 0, shapeType: 0, props: new Map(), children: [] }
  const dv = view(bytes)
  for (const child of records(bytes, record.start, record.end)) {
    if (child.type === REC_FSP && child.end - child.start >= 8) {
      shape.spid = dv.getUint32(child.start, true)
      shape.flags = dv.getUint32(child.start + 4, true)
      shape.shapeType = child.instance
    } else if (child.type === REC_FSPGR && child.end - child.start >= 16) {
      shape.groupBounds = readRect(bytes, child.start)
    } else if (child.type === REC_CHILD_ANCHOR && child.end - child.start >= 16) {
      shape.childAnchor = readRect(bytes, child.start)
    } else if (
      child.type === REC_OPT ||
      child.type === REC_SECONDARY_OPT ||
      child.type === REC_TERTIARY_OPT
    ) {
      for (let index = 0; index < child.instance; index++) {
        const at = child.start + index * 6
        if (at + 6 > child.end) break
        const pid = dv.getUint16(at, true) & 0x3fff
        const value = dv.getUint32(at + 2, true)
        shape.props.set(pid, value)
        if (pid === OPT_PIB) shape.pib = value
      }
    }
  }
  return shape
}

function parseSpgrContainer(bytes: Uint8Array, record: RecordHeader): Shape[] {
  const shapes: Shape[] = []
  for (const child of records(bytes, record.start, record.end)) {
    if (child.type === REC_SP_CONTAINER) shapes.push(parseSpContainer(bytes, child))
    else if (child.type === REC_SPGR_CONTAINER) {
      const group = parseSpgrContainer(bytes, child)
      // the first shape of a group container is the group itself; the rest are its members
      const [self, ...members] = group
      if (self) {
        self.children = members
        shapes.push(self)
      }
    }
  }
  return shapes
}

function parseOfficeArt(
  table: Uint8Array,
  fib: FibInfo,
  warnings: string[],
): OfficeArt | undefined {
  const { fcDggInfo: fc, lcbDggInfo: lcb } = fib
  if (lcb === 0) return undefined
  if (fc > table.length || lcb > table.length - fc) {
    warnings.push('the drawing data lies outside the table stream; floating pictures were omitted')
    return undefined
  }
  const limit = fc + lcb
  const blips: BlipStoreEntry[] = []
  const drawings = new Map<number, Shape[]>()
  const dv = view(table)
  let offset = fc
  const dgg = readRecord(table, offset, limit)
  if (!dgg || dgg.type !== REC_DGG_CONTAINER) {
    warnings.push('the drawing data has no drawing group container; floating pictures were omitted')
    return undefined
  }
  for (const child of records(table, dgg.start, dgg.end)) {
    if (child.type !== REC_BSTORE_CONTAINER) continue
    for (const entry of records(table, child.start, child.end)) {
      if (entry.type !== REC_FBSE) {
        if (BLIP_KINDS[entry.type]) blips.push({ source: 'table', offset: entry.start - 8 })
        continue
      }
      if (entry.end - entry.start < FBSE_BYTES) {
        blips.push({ source: 'table', offset: -1 })
        continue
      }
      const nameBytes = table[entry.start + 33]!
      const embedded = entry.start + FBSE_BYTES + nameBytes
      if (embedded + 8 <= entry.end && isBlipRecordAt(table, embedded))
        blips.push({ source: 'table', offset: embedded })
      else blips.push({ source: 'wordDocument', offset: dv.getUint32(entry.start + 28, true) })
    }
  }
  offset = dgg.start + dgg.length
  while (offset + 9 <= limit) {
    const label = table[offset]!
    const dg = readRecord(table, offset + 1, limit)
    if (!dg || dg.type !== REC_DG_CONTAINER) break
    const shapes: Shape[] = []
    for (const child of records(table, dg.start, dg.end)) {
      if (child.type === REC_SPGR_CONTAINER) shapes.push(...parseSpgrContainer(table, child))
      else if (child.type === REC_SP_CONTAINER) shapes.push(parseSpContainer(table, child))
    }
    drawings.set(label, shapes)
    offset = dg.start + dg.length
  }
  return { blips, drawings }
}

function indexShapes(shapes: Shape[], into: Map<number, Shape>): void {
  for (const shape of shapes) {
    if (shape.flags & FSP_PATRIARCH) {
      indexShapes(shape.children, into)
      continue
    }
    into.set(shape.spid, shape)
    indexShapes(shape.children, into)
  }
}

function mapChildRect(child: Rect, bounds: Rect, outer: Rect): Rect | undefined {
  const boundsWidth = bounds.right - bounds.left
  const boundsHeight = bounds.bottom - bounds.top
  if (boundsWidth <= 0 || boundsHeight <= 0) return undefined
  const scaleX = (outer.right - outer.left) / boundsWidth
  const scaleY = (outer.bottom - outer.top) / boundsHeight
  return {
    left: Math.round(outer.left + (child.left - bounds.left) * scaleX),
    top: Math.round(outer.top + (child.top - bounds.top) * scaleY),
    right: Math.round(outer.left + (child.right - bounds.left) * scaleX),
    bottom: Math.round(outer.top + (child.bottom - bounds.top) * scaleY),
  }
}

/** every leaf shape inside a shape with its rectangle in twips (groups are flattened) */
function leavesOf(shape: Shape, rect: Rect, out: Array<{ shape: Shape; rect: Rect }>): void {
  if (shape.flags & FSP_DELETED) return
  if (shape.flags & FSP_GROUP) {
    if (!shape.groupBounds) return
    for (const child of shape.children) {
      if (!child.childAnchor) continue
      const childRect = mapChildRect(child.childAnchor, shape.groupBounds, rect)
      if (childRect) leavesOf(child, childRect, out)
    }
    return
  }
  out.push({ shape, rect })
}

/** OfficeArt COLORREF -> RRGGBB; scheme/system/palette indexed colors have no literal value */
function colorHex(value: number): string | undefined {
  if (value & 0x19000000) return undefined
  const hex = (channel: number): string => channel.toString(16).padStart(2, '0')
  return (hex(value & 0xff) + hex((value >> 8) & 0xff) + hex((value >> 16) & 0xff)).toUpperCase()
}

/** a boolean property group: the "use" bit (16 above) says whether the bit was written */
function optBoolean(
  props: Map<number, number>,
  pid: number,
  bit: number,
  fallback: boolean,
): boolean {
  const value = props.get(pid)
  if (value === undefined || (value & (bit << 16)) === 0) return fallback
  return (value & bit) !== 0
}

function textBoxStyle(
  shape: Shape,
): Omit<FloatingTextBox, keyof FloatingShapeBase | 'kind' | 'content'> {
  const { props } = shape
  const filled = optBoolean(props, OPT_FILL_BOOLEANS, 0x10, true)
  const lined = optBoolean(props, OPT_LINE_BOOLEANS, 0x08, true)
  const inset = (pid: number, index: number): number => {
    const value = props.get(pid)
    return value === undefined || value > 0x7fffffff ? DEFAULT_INSETS_EMU[index]! : value
  }
  const anchorText = props.get(OPT_ANCHOR_TEXT) ?? 0
  const textFlow = props.get(OPT_TEXT_FLOW) ?? 0
  const textDirection =
    textFlow === 2
      ? 'vert270'
      : textFlow === 1 || textFlow === 3 || textFlow === 5
        ? 'vert'
        : undefined
  return {
    geometry: SHAPE_GEOMETRY[shape.shapeType] ?? 'rect',
    ...(filled ? { fill: colorHex(props.get(OPT_FILL_COLOR) ?? 0xffffff) ?? 'FFFFFF' } : {}),
    ...(lined
      ? {
          line: {
            color: colorHex(props.get(OPT_LINE_COLOR) ?? 0) ?? '000000',
            widthEmu: props.get(OPT_LINE_WIDTH) ?? DEFAULT_LINE_WIDTH_EMU,
          },
        }
      : {}),
    insetsEmu: [
      inset(OPT_DX_TEXT_LEFT, 0),
      inset(OPT_DY_TEXT_TOP, 1),
      inset(OPT_DX_TEXT_RIGHT, 2),
      inset(OPT_DY_TEXT_BOTTOM, 3),
    ],
    verticalAnchor: anchorText % 3 === 1 ? 'ctr' : anchorText % 3 === 2 ? 'b' : 't',
    ...(textDirection ? { textDirection } : {}),
    autoFit: optBoolean(props, OPT_TEXT_BOOLEANS, 0x02, false),
  }
}

const HORIZONTAL_RELATIVE: FloatingPicture['horizontalRelative'][] = ['margin', 'page', 'column']
const VERTICAL_RELATIVE: FloatingPicture['verticalRelative'][] = ['margin', 'page', 'paragraph']
/** MSOPRH and MSOPRV: the shape's own origin, which overrides the FSPA's bx and by. */
const POS_REL_H: FloatingPicture['horizontalRelative'][] = ['margin', 'page', 'column', 'character']
const POS_REL_V: FloatingPicture['verticalRelative'][] = ['margin', 'page', 'paragraph', 'line']
const WRAPS: FloatingPicture['wrap'][] = [
  'square',
  'topAndBottom',
  'square',
  'none',
  'tight',
  'through',
]

/**
 * Hands each text box story to its shapes: the text goes to the shape FTXBXS.lid
 * names (the chain start); shapes that only reach the story through their
 * lTxid story index are continuations unless that start shape is missing.
 */
class StoryClaims {
  private readonly bySpid = new Map<number, TextBoxStory>()
  private readonly byIndex = new Map<number, TextBoxStory>()
  private readonly claimed = new Set<TextBoxStory>()

  constructor(
    stories: TextBoxStory[],
    private readonly hasShape: (spid: number) => boolean,
  ) {
    for (const story of stories) {
      this.bySpid.set(story.spid, story)
      this.byIndex.set(story.index, story)
    }
  }

  claim(shape: Shape): { story: TextBoxStory; continuation: boolean } | undefined {
    const own = this.bySpid.get(shape.spid)
    const txid = shape.props.get(OPT_TXID)
    const story =
      own ?? (txid !== undefined && txid > 0 ? this.byIndex.get((txid >>> 16) - 1) : undefined)
    if (!story) return undefined
    const continuation = this.claimed.has(story) || (!own && this.hasShape(story.spid))
    if (!continuation) this.claimed.add(story)
    return { story, continuation }
  }

  unclaimed(): TextBoxStory[] {
    return [...this.byIndex.values()].filter((story) => !this.claimed.has(story))
  }
}

/**
 * Read the floating pictures and text boxes of the body and the header
 * document. Tables that are missing or malformed yield warnings, never errors.
 * Without text box stories every non-picture shape is skipped.
 */
export function parseFloatingPictures(
  source: DrawingSource,
  textBoxes?: TextBoxStories,
): FloatingPicturesResult {
  const warnings: string[] = [...(textBoxes?.warnings ?? [])]
  const { fib, table, wordDocument } = source
  const bodyAnchors = parsePlcSpa(table, fib.fcPlcSpaMom, fib.lcbPlcSpaMom, 'body', warnings)
  const headerAnchors = parsePlcSpa(table, fib.fcPlcSpaHdr, fib.lcbPlcSpaHdr, 'header', warnings)
  const storyCount = (textBoxes?.body.length ?? 0) + (textBoxes?.header.length ?? 0)
  const art =
    bodyAnchors.length === 0 && headerAnchors.length === 0
      ? undefined
      : parseOfficeArt(table, fib, warnings)
  if (!art) {
    if (storyCount > 0) warnings.push('a text box story has no shape anchor; its text is omitted')
    return { body: [], header: [], warnings }
  }
  const mediaCache = new Map<number, PictureMedia | undefined>()
  const mediaFor = (pib: number): PictureMedia | undefined => {
    if (mediaCache.has(pib)) return mediaCache.get(pib)
    const entry = art.blips[pib - 1]
    let media: PictureMedia | undefined
    if (!entry || entry.offset < 0)
      warnings.push('a floating picture references a missing image and was omitted')
    else {
      const bytes = entry.source === 'table' ? table : wordDocument
      media = readBlip(bytes, entry.offset, warnings)
    }
    mediaCache.set(pib, media)
    return media
  }
  let skippedShapes = false
  const convert = (
    anchors: ShapeAnchor[],
    shapes: Shape[],
    story: FloatingShape['story'],
    stories: TextBoxStory[],
  ): FloatingShape[] => {
    const byId = new Map<number, Shape>()
    indexShapes(shapes, byId)
    const claims = new StoryClaims(stories, (spid) => byId.has(spid))
    const out: FloatingShape[] = []
    for (const anchor of anchors) {
      const shape = byId.get(anchor.spid)
      if (!shape) {
        warnings.push('a floating shape anchor has no shape data and was omitted')
        continue
      }
      const found: Array<{ shape: Shape; rect: Rect }> = []
      leavesOf(shape, anchor, found)
      for (const leaf of found) {
        const isPicture = leaf.shape.pib !== undefined && leaf.shape.pib > 0
        const claim = claims.claim(leaf.shape)
        if (!isPicture && !claim) {
          skippedShapes = true
          continue
        }
        // an embedded object drawn as a picture also owns a (usually empty) story
        if (isPicture && claim && claim.story.content.text !== '')
          warnings.push('a picture shape also carries text; that text is omitted')
        const width = leaf.rect.right - leaf.rect.left
        const height = leaf.rect.bottom - leaf.rect.top
        if (width <= 0 || height <= 0) {
          warnings.push('a floating shape has no size and was omitted')
          continue
        }
        const base: FloatingShapeBase = {
          cp: anchor.cp,
          story,
          spid: leaf.shape.spid,
          leftTwips: leaf.rect.left,
          topTwips: leaf.rect.top,
          widthTwips: width,
          heightTwips: height,
          horizontalRelative:
            POS_REL_H[shape.props.get(OPT_POS_REL_H) ?? -1] ??
            HORIZONTAL_RELATIVE[anchor.bx] ??
            'column',
          verticalRelative:
            POS_REL_V[shape.props.get(OPT_POS_REL_V) ?? -1] ??
            VERTICAL_RELATIVE[anchor.by] ??
            'paragraph',
          wrap: WRAPS[anchor.wr] ?? 'square',
          behindText: anchor.belowText,
        }
        if (claim && !isPicture) {
          if (claim.continuation)
            warnings.push(
              'a linked text box continues another box; its text stays in the first box',
            )
          out.push({
            ...base,
            kind: 'textBox',
            ...textBoxStyle(leaf.shape),
            content: claim.continuation ? { text: '', paragraphs: [''] } : claim.story.content,
          })
          continue
        }
        const media = mediaFor(leaf.shape.pib!)
        if (!media) continue
        const crop = shapeCrop(leaf.shape.props)
        out.push({ ...base, kind: 'picture', media, ...(crop ? { crop } : {}) })
      }
    }
    if (claims.unclaimed().length > 0)
      warnings.push('a text box story has no shape anchor; its text is omitted')
    return out
  }
  const body = convert(bodyAnchors, art.drawings.get(0) ?? [], 'body', textBoxes?.body ?? [])
  const header = convert(
    headerAnchors,
    art.drawings.get(1) ?? [],
    'header',
    textBoxes?.header ?? [],
  )
  if (skippedShapes)
    warnings.push('floating shapes other than pictures and text boxes are not converted')
  return { body, header, warnings: [...new Set(warnings)] }
}
