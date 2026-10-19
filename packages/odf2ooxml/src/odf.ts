/**
 * ODF package loading: the zip container, its two XML parts, and the body
 * lookup both converters share. A file is accepted by the shape of its
 * `office:body` (the mimetype entry lies too often to be trusted), and any
 * structural surprise — missing part, malformed XML, an entry too large to
 * inflate — returns null so the caller reports an unreadable file instead of
 * converting half of one.
 */

import JSZip from 'jszip'
import type { JSZipObject } from 'jszip'
import { XNode, firstChild, nodeChildren, nodeName, parseOdfXml } from './xml'

/** why a conversion gave up; the app maps every code to a message */
export class OdfConvertError extends Error {
  constructor(public readonly code: 'unreadable' | 'not-a-text-document' | 'not-a-presentation') {
    super(code)
    this.name = 'OdfConvertError'
  }
}

/**
 * The largest a single zip entry may inflate to. 64 MiB is the bound the rest
 * of the repo puts on decompression; a part past it is refused on the size the
 * archive declares, so a bomb costs a lookup rather than the heap.
 */
export const MAX_ODF_PART_BYTES = 64 * 1024 * 1024

/** JSZip keeps the central-directory size on the entry's lazy compressed data */
function declaredSize(entry: JSZipObject): number {
  return (
    (entry as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0
  )
}

function exceedsCap(entry: JSZipObject): boolean {
  return declaredSize(entry) > MAX_ODF_PART_BYTES
}

/**
 * Inflate an entry while counting what actually comes out: a local header can
 * declare less than the stream yields, so the declared-size gate alone would
 * let a mislabelled part through.
 */
interface ChunkStream {
  on(event: 'data', fn: (chunk: Uint8Array) => void): ChunkStream
  on(event: 'error', fn: (err: Error) => void): ChunkStream
  on(event: 'end', fn: () => void): ChunkStream
  pause(): ChunkStream
  resume(): ChunkStream
}

function inflateCapped(entry: JSZipObject): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = []
    let total = 0
    // JSZip's typings stop at async(); internalStream is the public chunked API
    const stream = (
      entry as unknown as { internalStream(type: 'uint8array'): ChunkStream }
    ).internalStream('uint8array')
    stream.on('data', (chunk) => {
      total += chunk.length
      if (total > MAX_ODF_PART_BYTES) {
        stream.pause()
        reject(new OdfConvertError('unreadable'))
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

async function inflateText(entry: JSZipObject): Promise<string | null> {
  if (exceedsCap(entry)) return null
  try {
    return new TextDecoder().decode(await inflateCapped(entry))
  } catch {
    return null
  }
}

export interface OdfPackage {
  contentRoots: XNode[]
  stylesRoots: XNode[] | null
  /** read one zip entry by its path as spelled in xlink:href */
  read: (path: string) => Promise<Uint8Array | null>
}

export async function loadOdfPackage(bytes: Uint8Array): Promise<OdfPackage | null> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch {
    return null
  }
  const contentEntry = zip.file('content.xml')
  if (!contentEntry) return null
  const content = await inflateText(contentEntry)
  if (!content) return null
  const contentRoots = parseOdfXml(content)
  if (!contentRoots) return null
  const stylesEntry = zip.file('styles.xml')
  let stylesXml: string | null = null
  if (stylesEntry) {
    stylesXml = await inflateText(stylesEntry)
    if (stylesXml === null) return null
  }
  const stylesRoots = stylesXml ? parseOdfXml(stylesXml) : null
  return {
    contentRoots,
    stylesRoots,
    read: async (path: string) => {
      const clean = path.replace(/^\.\//, '')
      const entry = zip.file(clean) ?? zip.file(decodeURIComponent(clean))
      if (!entry) return null
      if (exceedsCap(entry)) {
        // Unlike a missing entry, which only costs the caller this one image, a
        // part this large is a file the app cannot convert: say so instead of
        // inflating it or silently handing back a document with holes in it.
        throw new OdfConvertError('unreadable')
      }
      try {
        return await inflateCapped(entry)
      } catch (err) {
        if (err instanceof OdfConvertError) throw err
        return null
      }
    },
  }
}

/** the office:text / office:presentation element under office:body, or null */
export function odfBody(
  pack: OdfPackage,
  kind: 'office:text' | 'office:presentation',
): XNode | null {
  for (const root of pack.contentRoots) {
    const body = firstChild(root, 'office:body')
    if (!body) continue
    const element = firstChild(body, kind)
    if (element) return element
  }
  return null
}

/** the office:automatic-styles element of a document part (content.xml or styles.xml) */
export function automaticStyles(pack: OdfPackage): XNode[] {
  const out: XNode[] = []
  for (const root of [...pack.contentRoots, ...(pack.stylesRoots ?? [])]) {
    const node = firstChild(root, 'office:automatic-styles')
    if (node) out.push(node)
  }
  return out
}

/** element children only, with their names — the body walkers' main iteration */
export function childElements(node: XNode): Array<{ name: string; node: XNode }> {
  const out: Array<{ name: string; node: XNode }> = []
  for (const child of nodeChildren(node)) {
    const name = nodeName(child)
    if (name && name !== '#comment') out.push({ name, node: child })
  }
  return out
}
