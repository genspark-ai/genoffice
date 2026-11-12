import { inflateRawSync } from 'node:zlib'

/**
 * PowerPoint 2007+ saves a shape's DrawingML as a metroBlob (OPT 0x03a9), an
 * OPC package whose drs/shapexml.xml carries what the legacy records cannot
 * (a per-paragraph first-line indent); its text is blanked, so only geometry is read.
 */

function zipEntry(bytes: Uint8Array, name: string): Uint8Array | null {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  for (let at = bytes.length - 22; at >= 0 && at >= bytes.length - 22 - 0xffff; at--) {
    if (v.getUint32(at, true) === 0x06054b50) {
      eocd = at
      break
    }
  }
  if (eocd < 0) return null
  const count = v.getUint16(eocd + 10, true)
  let at = v.getUint32(eocd + 16, true)
  for (let i = 0; i < count && at + 46 <= bytes.length; i++) {
    if (v.getUint32(at, true) !== 0x02014b50) return null
    const method = v.getUint16(at + 10, true)
    const compressed = v.getUint32(at + 20, true)
    const nameLength = v.getUint16(at + 28, true)
    const extraLength = v.getUint16(at + 30, true)
    const commentLength = v.getUint16(at + 32, true)
    const local = v.getUint32(at + 42, true)
    const entryName = Buffer.from(bytes.subarray(at + 46, at + 46 + nameLength)).toString('utf8')
    if (entryName === name) {
      if (local + 30 > bytes.length) return null
      const dataAt = local + 30 + v.getUint16(local + 26, true) + v.getUint16(local + 28, true)
      const data = bytes.subarray(dataAt, dataAt + compressed)
      if (method === 0) return data
      if (method !== 8) return null
      try {
        return new Uint8Array(inflateRawSync(data))
      } catch {
        return null
      }
    }
    at += 46 + nameLength + extraLength + commentLength
  }
  return null
}

/** the per-paragraph marL / indent (EMU) of the blob's text body, when it has one */
export function metroParagraphIndents(
  blob: Uint8Array,
): Array<{ marL?: number; indent?: number }> | null {
  const xml = zipEntry(blob, 'drs/shapexml.xml')
  if (!xml) return null
  const text = Buffer.from(xml).toString('utf8')
  const body = /<p:txBody>([\s\S]*?)<\/p:txBody>/.exec(text)?.[1]
  if (!body) return null
  const out: Array<{ marL?: number; indent?: number }> = []
  for (const paragraph of body.split(/<a:p[\s>]/).slice(1)) {
    const pPr = /<a:pPr\b([^>]*)>/.exec(paragraph)?.[1] ?? ''
    const marL = /\bmarL="(-?\d+)"/.exec(pPr)?.[1]
    const indent = /\bindent="(-?\d+)"/.exec(pPr)?.[1]
    out.push({
      ...(marL !== undefined ? { marL: Number(marL) } : {}),
      ...(indent !== undefined ? { indent: Number(indent) } : {}),
    })
  }
  return out
}
