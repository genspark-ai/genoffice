import { inflateSync } from 'node:zlib'

import { type PptRecord, walkRecordsLenient } from './records'

/**
 * The OfficeArt (Escher) layer of a .ppt: the shape tree under every
 * PPDrawing, the property table of each shape, and the picture store the
 * `pib` properties index into. The record framing is the same 8-byte header
 * MS-PPT uses, so the PPT record walker reads these too.
 */

export const ESCHER = {
  dggContainer: 0xf000,
  bStoreContainer: 0xf001,
  dgContainer: 0xf002,
  spgrContainer: 0xf003,
  spContainer: 0xf004,
  fbse: 0xf007,
  spgr: 0xf009,
  sp: 0xf00a,
  opt: 0xf00b,
  clientTextbox: 0xf00d,
  childAnchor: 0xf00f,
  clientAnchor: 0xf010,
  clientData: 0xf011,
  secondaryOpt: 0xf121,
  tertiaryOpt: 0xf122,
} as const

export const OPT = {
  rotation: 0x0004,
  dxTextLeft: 0x0081,
  dyTextTop: 0x0082,
  dxTextRight: 0x0083,
  dyTextBottom: 0x0084,
  wrapText: 0x0085,
  anchorText: 0x0087,
  textFlow: 0x0088,
  textBooleans: 0x00bf,
  cropFromTop: 0x0100,
  cropFromBottom: 0x0101,
  cropFromLeft: 0x0102,
  cropFromRight: 0x0103,
  pib: 0x0104,
  pictureTransparent: 0x0107,
  geoLeft: 0x0140,
  geoTop: 0x0141,
  geoRight: 0x0142,
  geoBottom: 0x0143,
  pVertices: 0x0145,
  pSegmentInfo: 0x0146,
  adjustValue: 0x0147,
  fillType: 0x0180,
  fillColor: 0x0181,
  fillOpacity: 0x0182,
  fillBackColor: 0x0183,
  fillBlip: 0x0186,
  fillAngle: 0x018b,
  fillFocus: 0x018c,
  fillToLeft: 0x018d,
  fillToTop: 0x018e,
  fillToRight: 0x018f,
  fillToBottom: 0x0190,
  fillShadeColors: 0x0197,
  fillBooleans: 0x01bf,
  lineColor: 0x01c0,
  lineWidth: 0x01cb,
  lineStyle: 0x01cd,
  lineDashing: 0x01ce,
  lineStartArrowhead: 0x01d0,
  lineEndArrowhead: 0x01d1,
  lineBooleans: 0x01ff,
  shadowBooleans: 0x023f,
  metroBlob: 0x03a9,
  groupShapeBooleans: 0x03bf,
} as const

/** FSP flags */
export const FSP = {
  group: 0x0001,
  child: 0x0002,
  patriarch: 0x0004,
  deleted: 0x0008,
  flipH: 0x0040,
  flipV: 0x0080,
  connector: 0x0100,
  haveAnchor: 0x0200,
  background: 0x0400,
  haveSpt: 0x0800,
} as const

export interface Rect {
  left: number
  top: number
  right: number
  bottom: number
}

export interface EscherShape {
  spid: number
  flags: number
  /** FSP instance: the msosp shape type */
  shapeType: number
  props: Map<number, number>
  /** complex property payloads (fComplex set): pid → bytes */
  complex: Map<number, Uint8Array>
  /** the slide-space anchor, master units (576/in) */
  anchor?: Rect
  /** the group-space anchor of a member shape */
  childAnchor?: Rect
  /** the coordinate space a group's members are anchored in */
  groupBounds?: Rect
  /** the OfficeArtClientTextbox record, when the shape carries one */
  textbox?: PptRecord
  /** the OfficeArtClientData record (OEPlaceholderAtom and friends) */
  clientData?: PptRecord
  children: EscherShape[]
}

function viewOf(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function readRect32(stream: Uint8Array, at: number): Rect {
  const v = viewOf(stream)
  return {
    left: v.getInt32(at, true),
    top: v.getInt32(at + 4, true),
    right: v.getInt32(at + 8, true),
    bottom: v.getInt32(at + 12, true),
  }
}

/**
 * PowerPoint's OfficeArtClientAnchor: 8 bytes of int16 (top, left, right,
 * bottom — note the order) or 16 bytes of int32 in the same order, master units.
 */
function readClientAnchor(stream: Uint8Array, record: PptRecord): Rect | undefined {
  const v = viewOf(stream)
  if (record.length >= 16) {
    return {
      top: v.getInt32(record.offset, true),
      left: v.getInt32(record.offset + 4, true),
      right: v.getInt32(record.offset + 8, true),
      bottom: v.getInt32(record.offset + 12, true),
    }
  }
  if (record.length >= 8) {
    return {
      top: v.getInt16(record.offset, true),
      left: v.getInt16(record.offset + 2, true),
      right: v.getInt16(record.offset + 4, true),
      bottom: v.getInt16(record.offset + 6, true),
    }
  }
  return undefined
}

/** The property table: `instance` fixed 6-byte entries, then the complex payloads in order. */
function readOpt(stream: Uint8Array, record: PptRecord, shape: EscherShape): void {
  const v = viewOf(stream)
  const end = record.offset + record.length
  let complexAt = record.offset + record.instance * 6
  const pending: Array<{ pid: number; length: number }> = []
  for (let index = 0; index < record.instance; index++) {
    const at = record.offset + index * 6
    if (at + 6 > end) break
    const head = v.getUint16(at, true)
    const pid = head & 0x3fff
    const value = v.getUint32(at + 2, true)
    if (head & 0x8000) pending.push({ pid, length: value })
    else shape.props.set(pid, value)
  }
  for (const { pid, length } of pending) {
    // PowerPoint writes some arrays (0xFFF0 element size) with a length that
    // leaves out the 6-byte array header; the bytes are there all the same
    let take = length
    if (complexAt + 6 <= end) {
      const count = v.getUint16(complexAt, true)
      let size = v.getUint16(complexAt + 4, true)
      if (size === 0xfff0) size = 4
      if (size > 0 && count * size === length && complexAt + length + 6 <= end) take = length + 6
    }
    if (complexAt + take > end) break
    shape.complex.set(pid, stream.subarray(complexAt, complexAt + take))
    complexAt += take
  }
}

export function parseSpContainer(stream: Uint8Array, record: PptRecord): EscherShape {
  const shape: EscherShape = {
    spid: 0,
    flags: 0,
    shapeType: 0,
    props: new Map(),
    complex: new Map(),
    children: [],
  }
  const v = viewOf(stream)
  for (const child of walkRecordsLenient(stream, record.offset, record.offset + record.length)) {
    switch (child.type) {
      case ESCHER.sp:
        if (child.length >= 8) {
          shape.spid = v.getUint32(child.offset, true)
          shape.flags = v.getUint32(child.offset + 4, true)
          shape.shapeType = child.instance
        }
        break
      case ESCHER.spgr:
        if (child.length >= 16) shape.groupBounds = readRect32(stream, child.offset)
        break
      case ESCHER.childAnchor:
        if (child.length >= 16) shape.childAnchor = readRect32(stream, child.offset)
        break
      case ESCHER.clientAnchor:
        shape.anchor = readClientAnchor(stream, child)
        break
      case ESCHER.opt:
      case ESCHER.secondaryOpt:
      case ESCHER.tertiaryOpt:
        readOpt(stream, child, shape)
        break
      case ESCHER.clientTextbox:
        shape.textbox = child
        break
      case ESCHER.clientData:
        shape.clientData = child
        break
    }
  }
  return shape
}

/** A hostile stream can chain spgrContainer records 8 bytes a level, so the
 *  descent is capped like the pptx engine's group parse (MAX_GROUP_DEPTH=64):
 *  deeper nesting degrades to an empty group instead of exhausting the stack. */
const MAX_SPGR_DEPTH = 64

/** The shapes of an SpgrContainer; a nested group's first shape is the group itself. */
export function parseSpgrContainer(
  stream: Uint8Array,
  record: PptRecord,
  depth = 0,
): EscherShape[] {
  const shapes: EscherShape[] = []
  if (depth >= MAX_SPGR_DEPTH) return shapes
  for (const child of walkRecordsLenient(stream, record.offset, record.offset + record.length)) {
    if (child.type === ESCHER.spContainer) shapes.push(parseSpContainer(stream, child))
    else if (child.type === ESCHER.spgrContainer) {
      const [self, ...members] = parseSpgrContainer(stream, child, depth + 1)
      if (self) {
        self.children = members
        shapes.push(self)
      }
    }
  }
  return shapes
}

/**
 * The top-level shapes of a drawing (a PPDrawing's OfficeArtDgContainer), in
 * z-order. The patriarch group (the first shape of the top SpgrContainer) is
 * unwrapped; a background shape sits directly in the DgContainer.
 */
export function parseDrawing(stream: Uint8Array, dgContainer: PptRecord): EscherShape[] {
  const shapes: EscherShape[] = []
  for (const child of walkRecordsLenient(
    stream,
    dgContainer.offset,
    dgContainer.offset + dgContainer.length,
  )) {
    if (child.type === ESCHER.spgrContainer) {
      for (const shape of parseSpgrContainer(stream, child)) {
        if (shape.flags & FSP.patriarch) shapes.push(...shape.children)
        else shapes.push(shape)
      }
    } else if (child.type === ESCHER.spContainer) {
      shapes.push(parseSpContainer(stream, child))
    }
  }
  return shapes
}

/* ---------- pictures ---------- */

export type PictureMediaType =
  'image/png' | 'image/jpeg' | 'image/x-emf' | 'image/x-wmf' | 'image/bmp' | 'image/tiff'
export const PICTURE_EXTENSIONS: Record<PictureMediaType, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
  'image/bmp': 'bmp',
  'image/tiff': 'tiff',
}

export interface PictureMedia {
  mediaType: PictureMediaType
  extension: string
  bytes: Uint8Array
}

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

function isBlipRecordAt(bytes: Uint8Array, offset: number): boolean {
  if (offset < 0 || offset + 8 > bytes.length) return false
  const v = viewOf(bytes)
  return (
    BLIP_KINDS[v.getUint16(offset + 2, true)] !== undefined &&
    (v.getUint16(offset, true) & 0xf) === 0
  )
}

/** a DIB blip is a BITMAPINFOHEADER + pixels; a 14-byte file header makes it a .bmp */
function dibToBmp(dib: Uint8Array): Uint8Array | undefined {
  if (dib.length < 40) return undefined
  const v = viewOf(dib)
  const headerSize = v.getUint32(0, true)
  const bitCount = v.getUint16(14, true)
  const compression = v.getUint32(16, true)
  const colorsUsed = v.getUint32(32, true)
  let colorTable = 0
  if (bitCount <= 8) colorTable = (colorsUsed || 1 << bitCount) * 4
  if (compression === 3 && headerSize === 40) colorTable += 12
  const out = new Uint8Array(14 + dib.length)
  out[0] = 0x42
  out[1] = 0x4d
  const ov = viewOf(out)
  ov.setUint32(2, out.length, true)
  ov.setUint32(10, 14 + headerSize + colorTable, true)
  out.set(dib, 14)
  return out
}

/** Decode one OfficeArtBlip record at `offset` of `bytes` (undefined for PICT or damaged data). */
export function readBlip(bytes: Uint8Array, offset: number): PictureMedia | undefined {
  if (offset < 0 || offset + 8 > bytes.length) return undefined
  const v = viewOf(bytes)
  const head = v.getUint16(offset, true)
  const type = v.getUint16(offset + 2, true)
  const length = v.getUint32(offset + 4, true)
  const start = offset + 8
  const end = Math.min(start + length, bytes.length)
  const kind = BLIP_KINDS[type]
  if (!kind || !kind.mediaType) return undefined
  const instance = head >> 4
  let dataStart = start + (kind.twoUidInstances.includes(instance) ? 32 : 16)
  let payload: Uint8Array
  if (kind.metafile) {
    if (dataStart + 34 > end) return undefined
    const cbSave = v.getUint32(dataStart + 28, true)
    const compression = bytes[dataStart + 32]!
    dataStart += 34
    const raw = bytes.slice(dataStart, dataStart + Math.min(cbSave, end - dataStart))
    if (compression === 0) {
      try {
        payload = new Uint8Array(inflateSync(raw))
      } catch {
        return undefined
      }
    } else payload = raw
    if (payload.length === 0) return undefined
  } else {
    dataStart += 1
    if (dataStart >= end) return undefined
    payload = bytes.slice(dataStart, end)
    if (kind.mediaType === 'image/bmp') {
      const bmp = dibToBmp(payload)
      if (!bmp) return undefined
      payload = bmp
    }
  }
  return {
    mediaType: kind.mediaType,
    extension: PICTURE_EXTENSIONS[kind.mediaType],
    bytes: payload,
  }
}

/**
 * The picture store. The DocumentContainer's PPDrawingGroup holds one FBSE
 * per picture whose foDelay is the blip's offset in the Pictures stream; a
 * shape's `pib` is the 1-based index into that list. Decks without an FBSE
 * list (or with foDelay garbage) fall back to the blips in stream order.
 */
export class PictureStore {
  private readonly offsets: number[] = []
  private readonly cache = new Map<number, PictureMedia | null>()

  constructor(
    private readonly pictures: Uint8Array | null,
    dggContainer: { stream: Uint8Array; record: PptRecord } | null,
  ) {
    if (dggContainer) {
      const { stream, record } = dggContainer
      const v = viewOf(stream)
      for (const child of walkRecordsLenient(
        stream,
        record.offset,
        record.offset + record.length,
      )) {
        if (child.type !== ESCHER.bStoreContainer) continue
        for (const entry of walkRecordsLenient(stream, child.offset, child.offset + child.length)) {
          if (entry.type !== ESCHER.fbse || entry.length < 36) {
            this.offsets.push(-1)
            continue
          }
          this.offsets.push(v.getUint32(entry.offset + 28, true))
        }
      }
    }
    if (this.offsets.length === 0 && pictures) {
      let at = 0
      while (at + 8 <= pictures.length) {
        if (!isBlipRecordAt(pictures, at)) break
        this.offsets.push(at)
        at += 8 + viewOf(pictures).getUint32(at + 4, true)
      }
    }
  }

  get size(): number {
    return this.offsets.length
  }

  /** the picture a `pib` names (1-based), decoded once */
  get(pib: number): PictureMedia | null {
    if (!this.pictures || pib <= 0 || pib > this.offsets.length) return null
    const cached = this.cache.get(pib)
    if (cached !== undefined) return cached
    const offset = this.offsets[pib - 1]!
    const media = isBlipRecordAt(this.pictures, offset)
      ? (readBlip(this.pictures, offset) ?? null)
      : null
    this.cache.set(pib, media)
    return media
  }
}

/* ---------- colors ---------- */

export interface ColorScheme {
  /** 8 RRGGBB strings: background, text, shadow, title, fill, accent, accent+hyperlink, accent+followed */
  colors: string[]
}

export const DEFAULT_SCHEME: ColorScheme = {
  colors: ['FFFFFF', '000000', '808080', '000000', '00CC99', '3333CC', 'CCCCFF', 'B2B2B2'],
}

const hex2 = (channel: number): string => channel.toString(16).padStart(2, '0').toUpperCase()

/** an OfficeArt COLORREF (0xBBGGRR + flag byte) → RRGGBB against a scheme */
/** the shape's own colour properties a system colour reference (fSysIndex) points back at */
export interface SysColorContext {
  fill?: number
  back?: number
  line?: number
}

/**
 * an OfficeArt COLORREF (0xBBGGRR + flag byte) → RRGGBB against a scheme. A system colour
 * (flag 0x10) names one of the shape's own colours — 0xF0 fill, 0xF5 fill back, 0xF2 line,
 * 0xF1 line-or-fill, 0xF7 fill-then-line, 0xF3 the scheme shadow — and may darken/lighten it (modifier byte 1,
 * parameter byte 2); a master background "navy, darkened 46%" is written exactly that way.
 */
export function officeArtColor(
  value: number,
  scheme: ColorScheme,
  context: SysColorContext = {},
  depth = 0,
): string | undefined {
  const flags = (value >>> 24) & 0xff
  if (flags & 0x08) return scheme.colors[value & 0xff] ?? scheme.colors[(value & 0xff) % 8]
  if (flags & 0x10) {
    if (depth > 2) return undefined
    const which = value & 0xff
    const modifier = (value >>> 8) & 0xff
    const param = (value >>> 16) & 0xff
    let source: number | undefined
    let slot: number
    if (which === 0xf0 || which === 0xf7 || which === 0xf4) [source, slot] = [context.fill, 4]
    else if (which === 0xf5) [source, slot] = [context.back, 4]
    else if (which === 0xf2 || which === 0xf6) [source, slot] = [context.line, 1]
    else if (which === 0xf1) [source, slot] = [context.line ?? context.fill, 4]
    else if (which === 0xf3) slot = 2
    else return undefined // Windows system colours (0x00-0x1D) have no document value
    const base =
      source === undefined
        ? scheme.colors[slot]
        : officeArtColor(source, scheme, context, depth + 1)
    if (!base) return undefined
    return modifyColor(base, modifier, param)
  }
  if (flags & 0x01) return undefined // palette index: no literal value
  return hex2(value & 0xff) + hex2((value >> 8) & 0xff) + hex2((value >> 16) & 0xff)
}

/** the system-colour modifiers: 1 darken, 2 lighten, 3 add gray, 4 subtract gray, 5 reverse-subtract gray */
function modifyColor(hex: string, modifier: number, param: number): string {
  if (modifier < 1 || modifier > 5) return hex
  const channels = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16))
  const out = channels.map((c) => {
    switch (modifier) {
      case 1:
        return (c * param) / 255
      case 2:
        return 255 - ((255 - c) * param) / 255
      case 3:
        return c + param
      case 4:
        return c - param
      default:
        return param - c
    }
  })
  return out.map((c) => hex2(Math.max(0, Math.min(255, Math.round(c))))).join('')
}

/** a ColorIndexStruct (red, green, blue, index) → RRGGBB; index 0xFE means the literal RGB */
export function colorIndexColor(value: number, scheme: ColorScheme): string | undefined {
  const index = (value >>> 24) & 0xff
  if (index === 0xfe)
    return hex2(value & 0xff) + hex2((value >> 8) & 0xff) + hex2((value >> 16) & 0xff)
  if (index === 0xff) return undefined
  return scheme.colors[index] ?? scheme.colors[index % 8]
}

/* ---------- geometry ---------- */

/** msosp shape type → OOXML preset name (the preset names descend from the msosp names) */
export const PRESET_BY_SHAPE_TYPE: Record<number, string> = {
  1: 'rect',
  2: 'roundRect',
  3: 'ellipse',
  4: 'diamond',
  5: 'triangle',
  6: 'rtTriangle',
  7: 'parallelogram',
  8: 'trapezoid',
  9: 'hexagon',
  10: 'octagon',
  11: 'plus',
  12: 'star5',
  13: 'rightArrow',
  14: 'rightArrow',
  15: 'homePlate',
  16: 'cube',
  17: 'wedgeRoundRectCallout',
  18: 'star16',
  19: 'arc',
  20: 'line',
  21: 'plaque',
  22: 'can',
  23: 'donut',
  32: 'straightConnector1',
  33: 'bentConnector2',
  34: 'bentConnector3',
  35: 'bentConnector4',
  36: 'bentConnector5',
  37: 'curvedConnector2',
  38: 'curvedConnector3',
  39: 'curvedConnector4',
  40: 'curvedConnector5',
  41: 'callout1',
  42: 'callout2',
  43: 'callout3',
  44: 'accentCallout1',
  45: 'accentCallout2',
  46: 'accentCallout3',
  47: 'borderCallout1',
  48: 'borderCallout2',
  49: 'borderCallout3',
  50: 'accentBorderCallout1',
  51: 'accentBorderCallout2',
  52: 'accentBorderCallout3',
  53: 'ribbon',
  54: 'ribbon2',
  55: 'chevron',
  56: 'pentagon',
  57: 'noSmoking',
  58: 'star8',
  59: 'star16',
  60: 'star32',
  61: 'wedgeRectCallout',
  62: 'wedgeRoundRectCallout',
  63: 'wedgeEllipseCallout',
  64: 'wave',
  65: 'foldedCorner',
  66: 'leftArrow',
  67: 'downArrow',
  68: 'upArrow',
  69: 'leftRightArrow',
  70: 'upDownArrow',
  71: 'irregularSeal1',
  72: 'irregularSeal2',
  73: 'lightningBolt',
  74: 'heart',
  75: 'rect',
  76: 'quadArrow',
  77: 'leftArrowCallout',
  78: 'rightArrowCallout',
  79: 'upArrowCallout',
  80: 'downArrowCallout',
  81: 'leftRightArrowCallout',
  82: 'upDownArrowCallout',
  83: 'quadArrowCallout',
  84: 'bevel',
  85: 'leftBracket',
  86: 'rightBracket',
  87: 'leftBrace',
  88: 'rightBrace',
  89: 'leftUpArrow',
  90: 'bentUpArrow',
  91: 'bentArrow',
  92: 'star24',
  93: 'stripedRightArrow',
  94: 'notchedRightArrow',
  95: 'blockArc',
  96: 'smileyFace',
  97: 'verticalScroll',
  98: 'horizontalScroll',
  99: 'circularArrow',
  101: 'uturnArrow',
  102: 'curvedRightArrow',
  103: 'curvedLeftArrow',
  104: 'curvedUpArrow',
  105: 'curvedDownArrow',
  106: 'cloudCallout',
  107: 'ellipseRibbon',
  108: 'ellipseRibbon2',
  109: 'flowChartProcess',
  110: 'flowChartDecision',
  111: 'flowChartInputOutput',
  112: 'flowChartPredefinedProcess',
  113: 'flowChartInternalStorage',
  114: 'flowChartDocument',
  115: 'flowChartMultidocument',
  116: 'flowChartTerminator',
  117: 'flowChartPreparation',
  118: 'flowChartManualInput',
  119: 'flowChartManualOperation',
  120: 'flowChartConnector',
  121: 'flowChartPunchedCard',
  122: 'flowChartPunchedTape',
  123: 'flowChartSummingJunction',
  124: 'flowChartOr',
  125: 'flowChartCollate',
  126: 'flowChartSort',
  127: 'flowChartExtract',
  128: 'flowChartMerge',
  129: 'flowChartOfflineStorage',
  130: 'flowChartOnlineStorage',
  131: 'flowChartMagneticTape',
  132: 'flowChartMagneticDisk',
  133: 'flowChartMagneticDrum',
  134: 'flowChartDisplay',
  135: 'flowChartDelay',
  176: 'flowChartAlternateProcess',
  177: 'flowChartOffpageConnector',
  178: 'callout1',
  179: 'accentCallout1',
  180: 'borderCallout1',
  181: 'accentBorderCallout1',
  182: 'leftRightUpArrow',
  183: 'sun',
  184: 'moon',
  185: 'bracketPair',
  186: 'bracePair',
  187: 'star4',
  188: 'doubleWave',
  189: 'actionButtonBlank',
  190: 'actionButtonHome',
  191: 'actionButtonHelp',
  192: 'actionButtonInformation',
  193: 'actionButtonForwardNext',
  194: 'actionButtonBackPrevious',
  195: 'actionButtonEnd',
  196: 'actionButtonBeginning',
  197: 'actionButtonReturn',
  198: 'actionButtonDocument',
  199: 'actionButtonSound',
  200: 'actionButtonMovie',
  202: 'rect',
}

export interface PathCommand {
  kind: 'move' | 'line' | 'curve' | 'close'
  points: Array<[number, number]>
}

export interface CustomPath {
  width: number
  height: number
  commands: PathCommand[]
}

/** elements of an OfficeArt complex array property: (count, allocated, elementSize) then data */
function readArray(bytes: Uint8Array): { count: number; size: number; data: Uint8Array } | null {
  if (bytes.length < 6) return null
  const v = viewOf(bytes)
  const count = v.getUint16(0, true)
  let size = v.getUint16(4, true)
  if (size === 0xfff0) size = 4
  if (size === 0 || count * size > bytes.length - 6) return null
  return { count, size, data: bytes.subarray(6, 6 + count * size) }
}

/**
 * A freeform's path from pVertices + pSegmentInfo, in the shape's own
 * coordinate space (geoRight × geoBottom). Segment words: the top three bits
 * are the kind (0 lineTo, 1 curveTo, 2 moveTo, 3 close, 4 end, 5 escape),
 * the low 13 bits the count.
 */
export function freeformPath(shape: EscherShape): CustomPath | null {
  const vertexBytes = shape.complex.get(OPT.pVertices)
  if (!vertexBytes) return null
  const vertices = readArray(vertexBytes)
  if (!vertices || vertices.count === 0) return null
  const v = viewOf(vertices.data)
  const points: Array<[number, number]> = []
  for (let index = 0; index < vertices.count; index++) {
    const at = index * vertices.size
    if (vertices.size >= 8) points.push([v.getInt32(at, true), v.getInt32(at + 4, true)])
    else points.push([v.getInt16(at, true), v.getInt16(at + 2, true)])
  }
  const width = shape.props.get(OPT.geoRight) ?? 21600
  const height = shape.props.get(OPT.geoBottom) ?? 21600
  const commands: PathCommand[] = []
  const segmentBytes = shape.complex.get(OPT.pSegmentInfo)
  const segments = segmentBytes ? readArray(segmentBytes) : null
  if (!segments) {
    commands.push({ kind: 'move', points: [points[0]!] })
    commands.push({ kind: 'line', points: points.slice(1) })
    commands.push({ kind: 'close', points: [] })
    return { width, height, commands }
  }
  const sv = viewOf(segments.data)
  let cursor = 0
  let open = false
  for (let index = 0; index < segments.count; index++) {
    const word = sv.getUint16(index * 2, true)
    const kind = word >> 13
    const count = word & 0x1fff
    if (kind === 4) break
    if (kind === 2) {
      if (cursor < points.length) {
        commands.push({ kind: 'move', points: [points[cursor]!] })
        cursor++
        open = true
      }
      continue
    }
    if (kind === 3) {
      if (open) commands.push({ kind: 'close', points: [] })
      open = false
      continue
    }
    if (kind === 0 || kind === 1) {
      const per = kind === 1 ? 3 : 1
      const take = Math.min(count * per, points.length - cursor)
      if (take <= 0) continue
      if (!open) {
        // a path that starts without a moveTo starts at its first vertex
        commands.push({ kind: 'move', points: [points[Math.max(cursor - 1, 0)]!] })
        open = true
      }
      commands.push({
        kind: kind === 1 ? 'curve' : 'line',
        points: points.slice(cursor, cursor + take),
      })
      cursor += take
      continue
    }
    // escape (5) and anything unknown: skip its count
  }
  if (commands.length === 0) return null
  return { width: width || 21600, height: height || 21600, commands }
}

/** An arc's ellipse and end points in its own geometry space (geoLeft..geoRight × geoTop..geoBottom). */
export interface ArcGeometry {
  box: Rect
  ellipse: Rect
  start: [number, number]
  end: [number, number]
}

/**
 * msosptArc stores its path in pVertices as the VML `wr` command — ellipse left/top,
 * right/bottom, then the start and end of a clockwise sweep — in unsigned 16-bit
 * coordinates (43200 overflows a signed one). Without vertices it is the top-right quarter.
 */
export function arcGeometry(shape: EscherShape): ArcGeometry {
  const box: Rect = {
    left: shape.props.get(OPT.geoLeft) ?? 0,
    top: shape.props.get(OPT.geoTop) ?? 0,
    right: shape.props.get(OPT.geoRight) ?? 21600,
    bottom: shape.props.get(OPT.geoBottom) ?? 21600,
  }
  const vertexBytes = shape.complex.get(OPT.pVertices)
  const vertices = vertexBytes ? readArray(vertexBytes) : null
  if (vertices && vertices.count >= 4) {
    const v = viewOf(vertices.data)
    const point = (index: number): [number, number] => {
      const at = index * vertices.size
      return vertices.size >= 8
        ? [v.getInt32(at, true), v.getInt32(at + 4, true)]
        : [v.getUint16(at, true), v.getUint16(at + 2, true)]
    }
    const [left, top] = point(0)
    const [right, bottom] = point(1)
    if (right !== left && bottom !== top) {
      return { box, ellipse: { left, top, right, bottom }, start: point(2), end: point(3) }
    }
  }
  return {
    box,
    ellipse: { left: -21600, top: 0, right: 21600, bottom: 43200 },
    start: [0, 0],
    end: [21600, 21600],
  }
}

/* ---------- property helpers ---------- */

/** an OfficeArt boolean property set: the high word says which low bits are in use */
export function boolProp(
  props: Map<number, number>,
  pid: number,
  bit: number,
): boolean | undefined {
  const value = props.get(pid)
  if (value === undefined) return undefined
  const useBit = bit << 16
  if (!(value & useBit)) return undefined
  return (value & bit) !== 0
}

/** a 16.16 fixed-point property as a float */
export function fixedProp(value: number): number {
  return (value | 0) / 65536
}
