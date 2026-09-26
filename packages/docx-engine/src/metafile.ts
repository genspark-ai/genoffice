import { decompressionStream, streamBytes } from './byte-stream'
import { convertEmfToDataUrl, convertWmfToDataUrl } from './vendor/emf-converter/index.mjs'

const EMF_MIMES = new Set(['image/emf', 'image/x-emf'])
// localized GDI facenames mapped to their CSS-resolvable names
const FONT_FAMILY_MAP = {
  游ゴシック: 'Yu Gothic', // yuu goshikku
  游明朝: 'Yu Mincho', // yuu minchou
  メイリオ: 'Meiryo',
  'ｍｓ ｐゴシック': 'MS PGothic',
  'ｍｓ ゴシック': 'MS Gothic',
  'ｍｓ ｕｉゴシック': 'MS UI Gothic',
  'ｍｓ ｐ明朝': 'MS PMincho',
  'ｍｓ 明朝': 'MS Mincho',
}
const WMF_MIMES = new Set(['image/wmf', 'image/x-wmf'])
// gzip-compressed metafiles (.emz/.wmz)
const EMZ_MIMES = new Set(['image/emz', 'image/x-emz'])
const WMZ_MIMES = new Set(['image/wmz', 'image/x-wmz'])

export function isMetafileMime(mime: string | undefined): mime is string {
  return (
    mime !== undefined &&
    (EMF_MIMES.has(mime) || WMF_MIMES.has(mime) || EMZ_MIMES.has(mime) || WMZ_MIMES.has(mime))
  )
}

function isGzip(bytes: Uint8Array): boolean {
  return bytes.length > 2 && bytes[0] === 0x1f && bytes[1] === 0x8b
}

/** Decompression bombs must not exhaust renderer memory during conversion */
export const MAX_METAFILE_GUNZIP_BYTES = 64 * 1024 * 1024

async function gunzip(bytes: Uint8Array): Promise<Uint8Array> {
  // Not `new Blob([bytes]).stream()`: jsdom, which the app tests run in, ships a
  // Blob with no stream() (#798), and the try/catch in metafileToDataUrl turned
  // that into a silent empty frame instead of a decoded metafile.
  const reader = streamBytes(bytes).pipeThrough(decompressionStream('gzip')).getReader()
  const chunks: Uint8Array[] = []
  let total = 0
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    total += value.byteLength
    if (total > MAX_METAFILE_GUNZIP_BYTES) {
      await reader.cancel()
      throw new Error(
        `metafile gunzip output exceeds ${MAX_METAFILE_GUNZIP_BYTES} bytes (possible zip bomb)`,
      )
    }
    chunks.push(value)
  }
  const out = new Uint8Array(total)
  let off = 0
  for (const chunk of chunks) {
    out.set(chunk, off)
    off += chunk.byteLength
  }
  return out
}

/** EMR_HEADER iType plus the ' EMF' signature at offset 40 */
function looksLikeEmf(bytes: Uint8Array): boolean {
  if (bytes.length < 44) return false
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return dv.getUint32(0, true) === 1 && dv.getUint32(40, true) === 0x464d4520
}

/** placeable-WMF magic, or a standard META_HEADER (type 1/2, HeaderSize 9) */
function looksLikeWmf(bytes: Uint8Array): boolean {
  if (bytes.length < 18) return false
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (dv.getUint32(0, true) === 0x9ac6cdd7) return true
  const type = dv.getUint16(0, true)
  return (type === 1 || type === 2) && dv.getUint16(2, true) === 9
}

/**
 * Render EMF/WMF (or gzipped EMZ/WMZ) bytes to a PNG data URL via the vendored
 * emf-converter. Returns null on parse failure or when no canvas API exists
 * (non-renderer environments), so callers keep their existing empty-frame
 * degrade. Failures are logged instead of silently swallowed.
 */
export interface MetafileRasterOptions {
  /** Longest raster side in device px; the converter scales down preserving the aspect ratio */
  maxSidePx?: number
}

export async function metafileToDataUrl(
  bytes: ArrayBuffer | Uint8Array,
  mime: string,
  raster: MetafileRasterOptions = {},
): Promise<string | null> {
  try {
    let u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    if (isGzip(u8)) u8 = await gunzip(u8)
    const buffer = u8.slice().buffer
    if (!isMetafileMime(mime)) return null
    // signature beats the declared mime: HWP-exported docx ship EMF bytes
    // under .wmf part names; mime only decides indeterminate bytes
    let isEmf: boolean
    if (looksLikeEmf(u8)) isEmf = true
    else if (looksLikeWmf(u8)) isEmf = false
    else isEmf = EMF_MIMES.has(mime) || EMZ_MIMES.has(mime)
    const opts = {
      dpiScale: 2,
      fontFamilyMap: FONT_FAMILY_MAP,
      ...(raster.maxSidePx ? { maxWidth: raster.maxSidePx, maxHeight: raster.maxSidePx } : {}),
    }
    const result = isEmf
      ? await convertEmfToDataUrl(buffer, opts)
      : await convertWmfToDataUrl(buffer, opts)
    if (result === null) {
      console.warn(`metafileToDataUrl: converter returned null (${mime}, ${u8.byteLength} bytes)`)
    }
    return result
  } catch (err) {
    console.warn(`metafileToDataUrl: conversion failed (${mime}):`, err)
    return null
  }
}
