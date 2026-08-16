import JSZip from 'jszip'

const EOCD_SIG = 0x06054b50
const CENTRAL_SIG = 0x02014b50
const UNICODE_PATH_ID = 0x7075

/**
 * Word resolves docx parts by the zip header file names and ignores Info-ZIP
 * Unicode Path (0x7075) extra fields; JSZip honors them, so a crc-valid but
 * conflicting field can shadow word/document.xml with another entry's bytes
 * (POI's unicode-path corpus). Blank the field id in the central directory so
 * JSZip falls back to the header names. Returns the input when nothing to do.
 */
function neutralizeUnicodePathFields(bytes: Uint8Array): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  const stop = Math.max(0, bytes.length - 22 - 0xffff)
  for (let i = bytes.length - 22; i >= stop; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      eocd = i
      break
    }
  }
  if (eocd < 0) return bytes
  const count = view.getUint16(eocd + 10, true)
  const cdOffset = view.getUint32(eocd + 16, true)
  if (count === 0xffff || cdOffset === 0xffffffff) return bytes // zip64: leave as-is
  let out: Uint8Array | null = null
  let p = cdOffset
  for (let i = 0; i < count; i++) {
    if (p + 46 > bytes.length || view.getUint32(p, true) !== CENTRAL_SIG) return out ?? bytes
    const nameLen = view.getUint16(p + 28, true)
    const extraLen = view.getUint16(p + 30, true)
    const commentLen = view.getUint16(p + 32, true)
    let q = p + 46 + nameLen
    const extraEnd = Math.min(q + extraLen, bytes.length)
    while (q + 4 <= extraEnd) {
      const fieldId = view.getUint16(q, true)
      const fieldLen = view.getUint16(q + 2, true)
      if (fieldId === UNICODE_PATH_ID) {
        out ??= new Uint8Array(bytes)
        out[q] = 0xff
        out[q + 1] = 0xff
      }
      q += 4 + fieldLen
    }
    p += 46 + nameLen + extraLen + commentLen
  }
  return out ?? bytes
}

/** Load a docx/zip resolving part names the way Word does. */
export async function loadDocxZip(bytes: Uint8Array): Promise<JSZip> {
  return JSZip.loadAsync(neutralizeUnicodePathFields(bytes))
}
