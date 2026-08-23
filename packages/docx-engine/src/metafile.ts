import { convertEmfToDataUrl, convertWmfToDataUrl } from './vendor/emf-converter/index.mjs'

const EMF_MIMES = new Set(['image/emf', 'image/x-emf'])
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

async function gunzip(bytes: Uint8Array): Promise<Uint8Array> {
  // copy to a fresh ArrayBuffer-backed view (BlobPart rejects ArrayBufferLike)
  const stream = new Blob([new Uint8Array(bytes)])
    .stream()
    .pipeThrough(new DecompressionStream('gzip'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

/** EMR_HEADER iType plus the ' EMF' signature at offset 40 */
function looksLikeEmf(bytes: Uint8Array): boolean {
  if (bytes.length < 44) return false
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return dv.getUint32(0, true) === 1 && dv.getUint32(40, true) === 0x464d4520
}

/**
 * Render EMF/WMF (or gzipped EMZ/WMZ) bytes to a PNG data URL via the vendored
 * emf-converter. Returns null on parse failure or when no canvas API exists
 * (non-renderer environments), so callers keep their existing empty-frame
 * degrade. Failures are logged instead of silently swallowed.
 */
export async function metafileToDataUrl(
  bytes: ArrayBuffer | Uint8Array,
  mime: string,
): Promise<string | null> {
  try {
    let u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    if (isGzip(u8)) u8 = await gunzip(u8)
    const buffer = u8.slice().buffer
    let isEmf: boolean
    if (EMF_MIMES.has(mime)) isEmf = true
    else if (WMF_MIMES.has(mime)) isEmf = false
    else if (EMZ_MIMES.has(mime) || WMZ_MIMES.has(mime)) isEmf = looksLikeEmf(u8)
    else return null
    const result = isEmf
      ? await convertEmfToDataUrl(buffer, { dpiScale: 2 })
      : await convertWmfToDataUrl(buffer, { dpiScale: 2 })
    if (result === null) {
      console.warn(`metafileToDataUrl: converter returned null (${mime}, ${u8.byteLength} bytes)`)
    }
    return result
  } catch (err) {
    console.warn(`metafileToDataUrl: conversion failed (${mime}):`, err)
    return null
  }
}
