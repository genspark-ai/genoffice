/**
 * Detect a pasted base64 image (a whole data: URL or a bare base64 dump) so the
 * composer can turn it into a real image attachment instead of chat text. Users
 * paste these when a model or another tool asked them for "the logo as base64";
 * as text they bloat the request and push the model into typing base64 back
 * through tool arguments, which gateways drop mid-stream (r182).
 */

export interface PastedBase64Image {
  bytes: Uint8Array
  ext: 'png' | 'jpg' | 'gif' | 'webp'
}

/** below this a base64 blob is small enough to be legitimate chat text */
const MIN_BARE_BASE64_CHARS = 4096

const DATA_URL_RE = /^data:image\/(png|jpe?g|gif|webp);base64,([A-Za-z0-9+/=\s]+)$/i
const BARE_BASE64_RE = /^[A-Za-z0-9+/\s]+={0,2}\s*$/

function decode(base64: string): Uint8Array | null {
  try {
    const clean = base64.replace(/\s+/g, '')
    const bin = atob(clean)
    const bytes = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
    return bytes
  } catch {
    return null
  }
}

/** image type from the file signature; the paste never carries a trustworthy name */
function sniffExt(b: Uint8Array): PastedBase64Image['ext'] | null {
  if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png'
  if (b.length > 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpg'
  if (b.length > 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) return 'gif'
  if (
    b.length > 12 &&
    b[0] === 0x52 &&
    b[1] === 0x49 &&
    b[2] === 0x46 &&
    b[3] === 0x46 &&
    b[8] === 0x57 &&
    b[9] === 0x45 &&
    b[10] === 0x42 &&
    b[11] === 0x50
  )
    return 'webp'
  return null
}

/**
 * The whole pasted text as one base64 image, or null when it is ordinary text.
 * Bare base64 (no data: prefix) must be long enough to be unambiguous and must
 * decode to a recognizable image signature.
 */
export function pastedBase64Image(text: string): PastedBase64Image | null {
  const trimmed = text.trim()
  if (trimmed.length < 64) return null
  const dataUrl = DATA_URL_RE.exec(trimmed)
  if (dataUrl) {
    const bytes = decode(dataUrl[2]!)
    if (!bytes || bytes.length === 0) return null
    const ext = sniffExt(bytes)
    // trust the declared type when the signature is unknown: the preview will show the truth
    const declared = dataUrl[1]!.toLowerCase()
    return {
      bytes,
      ext: ext ?? (declared === 'jpeg' ? 'jpg' : (declared as PastedBase64Image['ext'])),
    }
  }
  if (trimmed.length < MIN_BARE_BASE64_CHARS || !BARE_BASE64_RE.test(trimmed)) return null
  const bytes = decode(trimmed)
  if (!bytes) return null
  const ext = sniffExt(bytes)
  return ext ? { bytes, ext } : null
}
