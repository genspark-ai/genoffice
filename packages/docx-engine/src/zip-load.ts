import JSZip from 'jszip'
import { assertDeclaredSizesWithinLimits, assertZipInflatesWithinLimits } from '@genoffice/zip-gate'
import { needsOoxmlNormalization, normalizeOoxmlXml } from './ooxml-normalize'

// Re-exported for the engine's public API (index.ts) and the CLI's pre-open
// check: the implementations live in @genoffice/zip-gate so the pptx engine
// and the attachment parsers share one metered-inflation gate (see #759).
export {
  assertDeclaredSizesWithinLimits,
  assertZipInflatesWithinLimits,
  type DeclaredPart,
  type ZipLimits,
} from '@genoffice/zip-gate'
export { DEFAULT_ZIP_LIMITS as DOCX_ZIP_LIMITS } from '@genoffice/zip-gate'

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

/**
 * Cheap fast path: reject zip bombs using the uncompressed sizes the central
 * directory declares (JSZip keeps them in the lazy `_data` compressed object).
 *
 * Advisory only — it costs nothing when the archive tells the truth and nothing
 * at all when it lies. `assertZipInflatesWithinLimits` is the gate that holds
 * against a forged declaration.
 */
export function assertZipWithinLimits(zip: JSZip): void {
  const files = Object.values(zip.files).filter((f) => !f.dir)
  assertDeclaredSizesWithinLimits(
    files.map((file) => ({
      name: file.name,
      usize:
        (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0,
    })),
  )
}

// The gate parts carry the strict / non-canonical-prefix markers whenever the
// package needs normalizing (strict packages are strict in document.xml and
// their rels; prefix oddities live in document.xml).
const NORMALIZE_GATE_PARTS = ['word/document.xml', 'word/_rels/document.xml.rels', '_rels/.rels']

/**
 * ISO Strict OOXML and non-canonical namespace prefixes are normalized here,
 * at the single zip entry point, so parseDocx offsets and saveDocx patches
 * operate on identical part text (and saved packages come out transitional).
 */
async function normalizeOoxmlParts(zip: JSZip): Promise<void> {
  let needed = false
  for (const path of NORMALIZE_GATE_PARTS) {
    const file = zip.file(path)
    if (file && needsOoxmlNormalization(await file.async('string'))) {
      needed = true
      break
    }
  }
  if (!needed) return
  for (const file of Object.values(zip.files)) {
    if (file.dir || !/\.(xml|rels)$/i.test(file.name)) continue
    const xml = await file.async('string')
    if (needsOoxmlNormalization(xml)) zip.file(file.name, normalizeOoxmlXml(xml))
  }
}

/** Load a docx/zip resolving part names the way Word does. */
export async function loadDocxZip(bytes: Uint8Array): Promise<JSZip> {
  const prepared = neutralizeUnicodePathFields(bytes)
  // Metered inflation first: past this point no part is larger than a limit,
  // so the inflates normalizeOoxmlParts and every caller perform are bounded.
  await assertZipInflatesWithinLimits(prepared)
  const zip = await JSZip.loadAsync(prepared)
  assertZipWithinLimits(zip)
  await normalizeOoxmlParts(zip)
  return zip
}
