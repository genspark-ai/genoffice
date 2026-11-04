/**
 * Stage 3c-3: inline pictures from the legacy .doc Data stream.
 *
 * A picture is a \x01 character outside a field range whose CHPX run carries
 * sprmCFSpec (0x0855 = 1) and sprmCPicLocation (0x6A03, a signed 32-bit offset
 * into the CFB 'Data' stream). At that offset the stream holds a PICF header
 * (lcb, cbHeader, mm, xExt/yExt, dxaGoal/dyaGoal, mx/my) followed by the
 * OfficeArt inline shape whose blip store carries the picture as an
 * OfficeArtBlip record (PNG, JPEG, DIB, WMF or EMF; drawings.ts decodes it).
 * Pictures without a blip record fall back to a raw PNG/JPEG signature scan.
 * Corrupt picture data degrades per picture: a warning, and the picture is
 * skipped; only a PICF whose header is structurally invalid raises bad_picture.
 */
import { LegacyDocError, isLegacyDocError } from './errors'
import {
  inlineShapeCrop,
  isBlipRecordAt,
  readBlip,
  type PictureCrop,
  type PictureMedia,
  type PictureMediaType,
} from './drawings'
import { readFormattingDetailed, type FormattingSource, type RawCharacterRun } from './formatting'
import { iterateSprms, readSprmU8, readSprmU32 } from './sprm'

/** sprmCFSpec: the character is a special character (the picture placeholder) */
const SPRM_CF_SPEC = 0x0855
/** sprmCPicLocation: signed 32-bit offset of the picture in the Data stream */
const SPRM_C_PIC_LOCATION = 0x6a03
const PICTURE_CHAR = '\u0001'
const FIELD_INSTRUCTION = '\u0013'
const FIELD_SEPARATOR = '\u0014'
const FIELD_END = '\u0015'
/** PICF fields we read (MS-DOC 2.9.175 PICFAbstractType) */
const PICF_X_EXT = 0x08
const PICF_Y_EXT = 0x0a
const PICF_DXA_GOAL = 0x1c
const PICF_DYA_GOAL = 0x1e
const PICF_MX = 0x20
const PICF_MY = 0x22
const PICF_DXA_CROP_LEFT = 0x24
const PICF_DYA_CROP_TOP = 0x26
const PICF_DXA_CROP_RIGHT = 0x28
const PICF_DYA_CROP_BOTTOM = 0x2a
/** the size Word gives a picture dimension whose PICF goal is zero or negative */
const HAIRLINE_TWIPS = 11
/** 1 twip = 635 EMU */
export const EMU_PER_TWIP = 635

export interface ImageBlock {
  /** CP of the \x01 picture character in the body text */
  cp: number
  /** display width/height in twips (PICF dxaGoal/mx, then xExt, then dxaGoal) */
  widthTwips: number
  heightTwips: number
  crop?: PictureCrop
  mediaType: PictureMediaType
  /** exact image bytes (PNG chunk sequence, JPEG SOI..EOI range, metafile or .bmp) */
  bytes: Uint8Array
  /** non-fatal problems found while extracting this picture */
  warnings: string[]
}

/** the story to scan; defaults to the whole rawText (the body) */
export interface CpRange {
  startCp: number
  endCp: number
}

export interface InlineImageResult {
  images: ImageBlock[]
  warnings: string[]
}

function badPicture(detail: string): never {
  throw new LegacyDocError('bad_picture', detail)
}

function u16At(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint16(offset, true)
}

function i16At(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getInt16(offset, true)
}
function u32At(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(offset, true)
}

/** the CPs of body text picture characters outside field instructions (EMBED results are pictures) */
function pictureCharacterCps(rawText: string, range?: CpRange): number[] {
  const cps: number[] = []
  const fields: Array<'instruction' | 'result'> = []
  const end = Math.min(rawText.length, range?.endCp ?? rawText.length)
  for (let cp = range?.startCp ?? 0; cp < end; cp++) {
    const char = rawText[cp]
    if (char === FIELD_INSTRUCTION) fields.push('instruction')
    else if (char === FIELD_SEPARATOR) {
      if (fields.length > 0) fields[fields.length - 1] = 'result'
    } else if (char === FIELD_END) fields.pop()
    else if (char === PICTURE_CHAR && !fields.includes('instruction')) cps.push(cp)
  }
  return cps
}

function signed32(value: number): number {
  return value > 0x7fffffff ? value - 0x100000000 : value
}

/** the picture location and special-character flag of one CHPX run */
function pictureProperties(run: RawCharacterRun): {
  fSpec: boolean
  fcPic: number | undefined
} {
  let fSpec = false
  let fcPic: number | undefined
  for (const sprm of iterateSprms(run.sprms)) {
    if (sprm.opcode === SPRM_CF_SPEC) {
      const value = readSprmU8(sprm.operand)
      if (value !== undefined) fSpec = value !== 0
    } else if (sprm.opcode === SPRM_C_PIC_LOCATION) {
      const value = readSprmU32(sprm.operand)
      if (value !== undefined) fcPic = signed32(value)
    }
  }
  return { fSpec, fcPic }
}

/**
 * PICF crop twips against the goal size -> a:srcRect thousandths of a percent.
 * A side that is negative (padding) or a pair that removes the whole picture
 * is treated as uncropped.
 */
function picfCrop(data: Uint8Array, fcPic: number): PictureCrop | undefined {
  const dxaGoal = i16At(data, fcPic + PICF_DXA_GOAL)
  const dyaGoal = i16At(data, fcPic + PICF_DYA_GOAL)
  if (dxaGoal <= 0 || dyaGoal <= 0) return undefined
  const fraction = (twips: number, goal: number): number =>
    twips > 0 && twips < goal ? Math.round((twips * 100000) / goal) : 0
  const crop: PictureCrop = {
    l: fraction(i16At(data, fcPic + PICF_DXA_CROP_LEFT), dxaGoal),
    t: fraction(i16At(data, fcPic + PICF_DYA_CROP_TOP), dyaGoal),
    r: fraction(i16At(data, fcPic + PICF_DXA_CROP_RIGHT), dxaGoal),
    b: fraction(i16At(data, fcPic + PICF_DYA_CROP_BOTTOM), dyaGoal),
  }
  if (crop.l + crop.r >= 100000 || crop.t + crop.b >= 100000) return undefined
  return crop.l || crop.t || crop.r || crop.b ? crop : undefined
}

/** display size in twips: (dxaGoal - crop) * |mx| / 1000, then xExt, then dxaGoal */
function displaySize(
  data: Uint8Array,
  fcPic: number,
  crop: PictureCrop | undefined,
): { width: number; height: number } {
  const dxaGoal = i16At(data, fcPic + PICF_DXA_GOAL)
  const dyaGoal = i16At(data, fcPic + PICF_DYA_GOAL)
  // mx/my are stored as unsigned 16-bit scaling factors (48000 = 0.5 inch at 15 twips)
  const mx = u16At(data, fcPic + PICF_MX)
  const my = u16At(data, fcPic + PICF_MY)
  const xExt = i16At(data, fcPic + PICF_X_EXT)
  const yExt = i16At(data, fcPic + PICF_Y_EXT)
  const keptX = crop ? (100000 - crop.l - crop.r) / 100000 : 1
  const keptY = crop ? (100000 - crop.t - crop.b) / 100000 : 1
  let width = mx > 0 && dxaGoal > 0 ? Math.round((dxaGoal * keptX * mx) / 1000) : 0
  let height = my > 0 && dyaGoal > 0 ? Math.round((dyaGoal * keptY * my) / 1000) : 0
  // Word shows a negative goal as a hairline (11 twips), not as its magnitude
  if (width <= 0) width = xExt > 0 ? xExt : dxaGoal > 0 ? dxaGoal : HAIRLINE_TWIPS
  if (height <= 0) height = yExt > 0 ? yExt : dyaGoal > 0 ? dyaGoal : HAIRLINE_TWIPS
  return { width, height }
}

/** cut an exact PNG from the signature through the end of the IEND chunk */
function pngBytes(range: Uint8Array, start: number): Uint8Array | undefined {
  let offset = start + 8
  while (offset + 12 <= range.length) {
    const length =
      ((range[offset]! << 24) |
        (range[offset + 1]! << 16) |
        (range[offset + 2]! << 8) |
        range[offset + 3]!) >>>
      0
    const type = String.fromCharCode(
      range[offset + 4]!,
      range[offset + 5]!,
      range[offset + 6]!,
      range[offset + 7]!,
    )
    const next = offset + 12 + length
    if (next > range.length) return undefined
    if (type === 'IEND') return range.slice(start, next)
    offset = next
  }
  return undefined
}

/** cut a JPEG from SOI (FFD8FF) through the first EOI (FFD9) */
function jpegBytes(range: Uint8Array, start: number): Uint8Array | undefined {
  for (let offset = start + 2; offset + 1 < range.length; offset++) {
    if (range[offset] === 0xff && range[offset + 1] === 0xd9) return range.slice(start, offset + 2)
  }
  return undefined
}

function findSignature(range: Uint8Array, signature: number[]): number {
  outer: for (let offset = 0; offset + signature.length <= range.length; offset++) {
    for (let index = 0; index < signature.length; index++)
      if (range[offset + index] !== signature[index]) continue outer
    return offset
  }
  return -1
}

/** extract the picture out of one PICF range: its OfficeArt blip record, else a raw PNG/JPEG */
function extractImage(
  data: Uint8Array,
  fcPic: number,
  cbHeader: number,
  lcb: number,
  warnings: string[],
): PictureMedia | undefined {
  const range = data.subarray(fcPic, fcPic + lcb)
  // a byte pattern that merely looks like a blip header must not shadow a raw PNG/JPEG payload
  let blipWarnings: string[] = []
  for (let offset = cbHeader; offset + 8 <= range.length; offset++) {
    if (!isBlipRecordAt(range, offset)) continue
    const attempt: string[] = []
    const media = readBlip(range, offset, attempt)
    if (media) return media
    if (blipWarnings.length === 0) blipWarnings = attempt
  }
  const png = findSignature(range, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  if (png >= 0) {
    const bytes = pngBytes(range, png)
    if (!bytes) {
      warnings.push('a PNG picture is truncated (no complete IEND chunk) and was omitted')
      return undefined
    }
    return { mediaType: 'image/png', extension: 'png', bytes }
  }
  const jpeg = findSignature(range, [0xff, 0xd8, 0xff])
  if (jpeg >= 0) {
    const bytes = jpegBytes(range, jpeg)
    if (!bytes) {
      warnings.push('a JPEG picture is truncated (no EOI marker) and was omitted')
      return undefined
    }
    return { mediaType: 'image/jpeg', extension: 'jpeg', bytes }
  }
  if (blipWarnings.length > 0) warnings.push(...blipWarnings)
  else warnings.push('an embedded picture uses an unsupported format and was omitted')
  return undefined
}

/**
 * Read the inline pictures of the body story. Unsupported formats and a
 * missing Data stream are warnings; corrupt PICF / truncated image data raise
 * 'bad_picture'. Picture characters inside field ranges are ignored.
 */
export function parseInlineImages(source: FormattingSource, range?: CpRange): InlineImageResult {
  const warnings: string[] = []
  const cps = pictureCharacterCps(source.rawText, range)
  if (cps.length === 0) return { images: [], warnings: [] }
  let read
  try {
    read = readFormattingDetailed(source, [])
  } catch (error) {
    if (isLegacyDocError(error)) throw error
    throw error
  }
  const data = source.stream?.('Data')
  if (!data) {
    warnings.push('the document has pictures but no Data stream; images are omitted')
    return { images: [], warnings }
  }
  const images: ImageBlock[] = []
  for (const cp of cps) {
    const run = read.characterRuns.find((entry) => entry.cpStart <= cp && entry.cpEnd > cp)
    if (!run) {
      warnings.push('a picture character has no character properties and was omitted')
      continue
    }
    let properties
    try {
      properties = pictureProperties(run)
    } catch (error) {
      if (isLegacyDocError(error)) throw error
      throw error
    }
    if (!properties.fSpec || properties.fcPic === undefined) {
      warnings.push('a picture character has no picture data and was omitted')
      continue
    }
    const fcPic = properties.fcPic
    if (fcPic < 0 || fcPic + 8 > data.length)
      badPicture('the picture offset ' + fcPic + ' lies outside the Data stream')
    const lcb = u32At(data, fcPic)
    const cbHeader = u16At(data, fcPic + 4)
    if (lcb < 8 || cbHeader < 0x24 || cbHeader > lcb)
      badPicture(
        'the PICF header at Data offset ' +
          fcPic +
          ' is invalid (lcb ' +
          lcb +
          ', cbHeader ' +
          cbHeader +
          ')',
      )
    if (fcPic + lcb > data.length)
      badPicture(
        'the picture at Data offset ' +
          fcPic +
          ' needs ' +
          lcb +
          ' bytes but the stream has ' +
          (data.length - fcPic),
      )
    const picWarnings: string[] = []
    const extracted = extractImage(data, fcPic, cbHeader, lcb, picWarnings)
    if (!extracted) {
      warnings.push(...picWarnings)
      continue
    }
    // dxaGoal/mx already describe the shape-cropped picture; only PICF twip crops shrink it
    const shapeCrop = inlineShapeCrop(data, fcPic + cbHeader, fcPic + lcb)
    const crop = shapeCrop ?? picfCrop(data, fcPic)
    const size = displaySize(data, fcPic, shapeCrop ? undefined : crop)
    if (size.width <= 0 || size.height <= 0) {
      warnings.push('a picture declares no display size and was omitted')
      continue
    }
    images.push({
      cp,
      widthTwips: size.width,
      heightTwips: size.height,
      ...(crop ? { crop } : {}),
      mediaType: extracted.mediaType,
      bytes: extracted.bytes,
      warnings: picWarnings,
    })
  }
  return { images, warnings: [...new Set(warnings)] }
}

/** The image dimensions are in twips, the writer needs EMU. */
export function imageExtentEmu(image: ImageBlock): { cx: number; cy: number } {
  return { cx: image.widthTwips * EMU_PER_TWIP, cy: image.heightTwips * EMU_PER_TWIP }
}
