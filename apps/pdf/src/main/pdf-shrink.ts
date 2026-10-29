import { deflateSync, inflateSync } from 'node:zlib'
import { PDFDict, PDFDocument, PDFName, PDFNumber, PDFRawStream, PDFRef } from 'pdf-lib'

/**
 * Reduce File Size, the parts that are safe without re-rendering anything:
 * every plain FlateDecode stream is inflated and re-deflated at level 9
 * (producers commonly ship level 6, or store already-weak blobs), and Flate
 * images larger than `maxDimension` are downsampled to it with their /Width
 * //Height and /ColorSpace rewritten. Anything else — JPX/CCITT/JPEG streams,
 * predictor-encoded data, image masks — is left byte-identical: a best-effort
 * pass that can only make the file smaller, never broken.
 */

export interface ShrinkResult {
  streamsRecompressed: number
  imagesDownsampled: number
}

export const DEFAULT_MAX_DIMENSION = 2480 // ~300 DPI on A4

/** what a stream must look like before its pixels may be resampled */
interface DownsampleCandidate {
  width: number
  height: number
  channels: number
}

/**
 * An image qualifies only when its decoded bytes can be resampled and
 * re-encoded faithfully: oversized, 8 bits per component, a bare device
 * colour space (arrays like `[/ICCBased …]` or `[/Indexed …]` are not
 * renamable) and a plain Flate payload. An /SMask with a /Matte never
 * qualifies: its dimensions must keep matching the base image's.
 */
function downsampleCandidate(
  doc: PDFDocument,
  dict: PDFDict,
  maxDimension: number,
): DownsampleCandidate | null {
  if (Math.max(asInt(dict.get(PDFName.of('Width'))), asInt(dict.get(PDFName.of('Height')))) === 0)
    return null
  if (
    Math.max(asInt(dict.get(PDFName.of('Width'))), asInt(dict.get(PDFName.of('Height')))) <=
    maxDimension
  )
    return null
  if (asInt(dict.get(PDFName.of('BitsPerComponent')) ?? 8) !== 8) return null
  // lookupMaybe throws when the object resolves to another type, and array
  // colour spaces (most photo scans) do exactly that — lookup and test instead
  const cs = doc.context.lookup(dict.get(PDFName.of('ColorSpace')))
  const channels = cs instanceof PDFName ? channelCount(cs) : 0
  if (channels === 0) return null
  if (String(dict.get(PDFName.of('Filter'))) !== '/FlateDecode') return null
  if (dict.get(PDFName.of('Matte')) !== undefined) return null
  return {
    width: asInt(dict.get(PDFName.of('Width'))),
    height: asInt(dict.get(PDFName.of('Height'))),
    channels,
  }
}

/** bilinear-free nearest-column/row resample: fast, and for scans visually identical at print size */
export function resamplePixels(
  pixels: Uint8Array,
  width: number,
  height: number,
  channels: number,
  newWidth: number,
  newHeight: number,
): Uint8Array {
  const out = new Uint8Array(newWidth * newHeight * channels)
  for (let y = 0; y < newHeight; y++) {
    const sy = Math.min(height - 1, Math.floor(((y + 0.5) * height) / newHeight))
    for (let x = 0; x < newWidth; x++) {
      const sx = Math.min(width - 1, Math.floor(((x + 0.5) * width) / newWidth))
      const src = (sy * width + sx) * channels
      const dst = (y * newWidth + x) * channels
      for (let c = 0; c < channels; c++) out[dst + c] = pixels[src + c]!
    }
  }
  return out
}

function asInt(value: unknown): number {
  const n = Number(String(value ?? '').replace(/^\//, ''))
  return Number.isFinite(n) ? n : 0
}

function decodeFlate(contents: Uint8Array, dict: PDFDict): Uint8Array | null {
  // /DecodeParms (predictors) change byte semantics — those streams stay as they are
  if (dict.get(PDFName.of('DecodeParms')) !== undefined) return null
  try {
    return new Uint8Array(inflateSync(Buffer.from(contents)))
  } catch {
    return null
  }
}

function channelCount(name: PDFName | undefined): 1 | 3 | 4 | 0 {
  // Only bare device names, resolved through the context: a regex over the
  // stringified object matched /DeviceN [/DeviceCMYK …] and recolored the
  // image, while Indexed/Separation/ICCBased passed by accident.
  switch (name?.asString()) {
    case '/DeviceGray':
      return 1
    case '/DeviceRGB':
      return 3
    case '/DeviceCMYK':
      return 4
    default:
      return 0
  }
}

/**
 * Walk every indirect object and shrink what qualifies, in place. Callers save
 * afterwards and compare sizes for themselves.
 */
export function shrinkDocument(
  doc: PDFDocument,
  maxDimension = DEFAULT_MAX_DIMENSION,
): ShrinkResult {
  let streamsRecompressed = 0
  let imagesDownsampled = 0

  // /SMask pairs: the mask's dimensions must keep matching its base image's,
  // so the two may only shrink together — never one without the other. Record
  // the pairing first; enumeration order cannot decide it.
  const refKey = (ref: PDFRef): string => `${ref.objectNumber}:${ref.generationNumber}`
  const baseOfMask = new Map<string, PDFRawStream>()
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue
    if (String(obj.dict.get(PDFName.of('Subtype')) ?? '') !== '/Image') continue
    const mask = obj.dict.get(PDFName.of('SMask'))
    if (mask instanceof PDFRef) {
      const resolved = doc.context.lookup(mask)
      if (resolved instanceof PDFRawStream) baseOfMask.set(refKey(mask), obj)
    }
  }
  /** the other side of an /SMask pair, when it is an image that shrinks the same way */
  const pairShrinksToo = (other: PDFRawStream, mine: DownsampleCandidate): boolean => {
    const cand = downsampleCandidate(doc, other.dict, maxDimension)
    return cand !== null && cand.width === mine.width && cand.height === mine.height
  }

  for (const [ref, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue
    const dict: PDFDict = obj.dict
    const filter = String(dict.get(PDFName.of('Filter')) ?? '')
    const isImage = String(dict.get(PDFName.of('Subtype')) ?? '') === '/Image'

    if (isImage) {
      const cand = downsampleCandidate(doc, dict, maxDimension)
      if (!cand) continue
      // a base shrinks only when its mask shrinks with it, and a mask only
      // when its base does — a DCT base with a Flate mask must keep both
      const mask = dict.get(PDFName.of('SMask'))
      if (mask !== undefined) {
        const resolved = mask instanceof PDFRef ? doc.context.lookup(mask) : null
        if (!(resolved instanceof PDFRawStream) || !pairShrinksToo(resolved, cand)) continue
      }
      const base = baseOfMask.get(refKey(ref))
      if (base && !pairShrinksToo(base, cand)) continue
      const decoded = decodeFlate(obj.contents, dict)
      if (!decoded || decoded.length < cand.width * cand.height * cand.channels) continue
      const scale = maxDimension / Math.max(cand.width, cand.height)
      const newWidth = Math.max(1, Math.floor(cand.width * scale))
      const newHeight = Math.max(1, Math.floor(cand.height * scale))
      const pixels = resamplePixels(
        decoded,
        cand.width,
        cand.height,
        cand.channels,
        newWidth,
        newHeight,
      )
      // Clone the producer's dict and retitle only the geometry: a fresh
      // five-key dict dropped /SMask, /Mask, /Decode, /Interpolate, /Intent,
      // /OC, /StructParent and /Metadata — transparency turned opaque and
      // inverted CMYK/gray flipped. The color space is untouched: only bare
      // device names qualify, so the channels already describe this dict.
      const freshDict = dict.clone()
      freshDict.set(PDFName.of('Width'), PDFNumber.of(newWidth))
      freshDict.set(PDFName.of('Height'), PDFNumber.of(newHeight))
      const packed = deflateSync(pixels, { level: 9 })
      freshDict.set(PDFName.of('Length'), PDFNumber.of(packed.length))
      doc.context.assign(ref, PDFRawStream.of(freshDict, packed))
      imagesDownsampled++
      continue
    }

    if (filter === '/FlateDecode') {
      const decoded = decodeFlate(obj.contents, dict)
      if (!decoded) continue
      const packed = deflateSync(decoded, { level: 9 })
      if (packed.length >= obj.contents.length) continue
      doc.context.assign(ref, PDFRawStream.of(dict, packed))
      streamsRecompressed++
    }
  }
  return { streamsRecompressed, imagesDownsampled }
}
