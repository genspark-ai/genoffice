/**
 * HWPX container access: an OCF-style zip (`mimetype` = application/hwp+zip,
 * Contents/content.hpf as the OPF manifest + spine, Contents/header.xml for the
 * shared style tables, Contents/sectionN.xml for the body, BinData/* for
 * embedded pictures). Every part is inflated through a size cap, the same
 * zip-bomb rule the ODF filters follow.
 */

import JSZip from 'jszip'
import type { JSZipObject } from 'jszip'
import { attr, descendants, parseXml, type XNode } from './xml'

/** one part may inflate to at most this many bytes */
export const MAX_HWPX_PART_BYTES = 64 * 1024 * 1024

export type HwpConvertErrorCode = 'unreadable' | 'not-hwpx' | 'hwp-binary' | 'no-body'

export class HwpConvertError extends Error {
  constructor(readonly code: HwpConvertErrorCode) {
    super(`hwpx import failed: ${code}`)
    this.name = 'HwpConvertError'
  }
}

export interface HwpxPackage {
  header: XNode[]
  /** section roots in spine order */
  sections: XNode[][]
  /** BinData bytes + declared media type by manifest id (hc:img binaryItemIDRef) */
  readBinary: (id: string) => Promise<{ bytes: Uint8Array; mediaType: string; path: string } | null>
}

const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0]

function declaredSize(entry: JSZipObject): number {
  return (
    (entry as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0
  )
}

interface ChunkStream {
  on(event: 'data', fn: (chunk: Uint8Array) => void): ChunkStream
  on(event: 'error', fn: (err: Error) => void): ChunkStream
  on(event: 'end', fn: () => void): ChunkStream
  pause(): ChunkStream
  resume(): ChunkStream
}

/** inflate while counting what actually comes out (a header can under-declare) */
function inflateCapped(entry: JSZipObject): Promise<Uint8Array> {
  if (declaredSize(entry) > MAX_HWPX_PART_BYTES) {
    return Promise.reject(new HwpConvertError('unreadable'))
  }
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = []
    let total = 0
    const stream = (
      entry as unknown as { internalStream(type: 'uint8array'): ChunkStream }
    ).internalStream('uint8array')
    stream.on('data', (chunk) => {
      total += chunk.length
      if (total > MAX_HWPX_PART_BYTES) {
        stream.pause()
        reject(new HwpConvertError('unreadable'))
        return
      }
      chunks.push(chunk)
    })
    stream.on('error', (err) => reject(err))
    stream.on('end', () => {
      const out = new Uint8Array(total)
      let offset = 0
      for (const chunk of chunks) {
        out.set(chunk, offset)
        offset += chunk.length
      }
      resolve(out)
    })
    stream.resume()
  })
}

async function readXmlPart(zip: JSZip, path: string): Promise<XNode[] | null> {
  const entry = zip.file(path)
  if (!entry) return null
  try {
    return parseXml(new TextDecoder().decode(await inflateCapped(entry)))
  } catch {
    return null
  }
}

/** true for an HWP 5.x binary (OLE compound file) — often served under a .hwpx name */
export function isHwpBinary(bytes: Uint8Array): boolean {
  return OLE_MAGIC.every((byte, index) => bytes[index] === byte)
}

export async function loadHwpxPackage(bytes: Uint8Array): Promise<HwpxPackage> {
  if (isHwpBinary(bytes)) throw new HwpConvertError('hwp-binary')
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch {
    throw new HwpConvertError('unreadable')
  }
  const mimetypeEntry = zip.file('mimetype')
  const mimetype = mimetypeEntry
    ? new TextDecoder().decode(await inflateCapped(mimetypeEntry)).trim()
    : ''
  if (mimetype && mimetype !== 'application/hwp+zip' && mimetype !== 'application/haansofthwpx') {
    throw new HwpConvertError('not-hwpx')
  }

  // manifest + spine (Contents/content.hpf); fall back to sectionN order
  const manifest = new Map<string, { href: string; mediaType: string }>()
  const spine: string[] = []
  const hpf = await readXmlPart(zip, 'Contents/content.hpf')
  if (hpf) {
    for (const item of descendants(hpf, 'opf:item')) {
      const id = attr(item, 'id')
      const href = attr(item, 'href')
      if (id && href) manifest.set(id, { href, mediaType: attr(item, 'media-type') ?? '' })
    }
    for (const ref of descendants(hpf, 'opf:itemref')) {
      const idref = attr(ref, 'idref')
      const href = idref ? manifest.get(idref)?.href : undefined
      if (href && /section\d+\.xml$/i.test(href)) spine.push(href)
    }
  }
  if (spine.length === 0) {
    const found = Object.keys(zip.files)
      .filter((name) => /^Contents\/section\d+\.xml$/i.test(name))
      .sort((a, b) => Number(a.match(/(\d+)\.xml$/i)![1]) - Number(b.match(/(\d+)\.xml$/i)![1]))
    spine.push(...found)
  }
  if (!mimetype && spine.length === 0) throw new HwpConvertError('not-hwpx')

  const header = await readXmlPart(zip, 'Contents/header.xml')
  if (!header) throw new HwpConvertError('unreadable')
  const sections: XNode[][] = []
  for (const href of spine) {
    const path =
      href.startsWith('Contents/') || href.startsWith('/') ? href.replace(/^\//, '') : href
    const root = (await readXmlPart(zip, path)) ?? (await readXmlPart(zip, `Contents/${path}`))
    if (root) sections.push(root)
  }
  if (sections.length === 0) throw new HwpConvertError('no-body')

  const readBinary: HwpxPackage['readBinary'] = async (id) => {
    const item = manifest.get(id)
    const candidates = item ? [item.href] : []
    // producers without a manifest entry still name the file after the id
    if (!item) {
      for (const name of Object.keys(zip.files)) {
        if (name.startsWith('BinData/') && name.slice(8).replace(/\.[^.]+$/, '') === id) {
          candidates.push(name)
        }
      }
    }
    for (const path of candidates) {
      const entry = zip.file(path) ?? zip.file(path.replace(/^\//, ''))
      if (!entry) continue
      try {
        return { bytes: await inflateCapped(entry), mediaType: item?.mediaType ?? '', path }
      } catch {
        return null
      }
    }
    return null
  }

  return { header, sections, readBinary }
}
