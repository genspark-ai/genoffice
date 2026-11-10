import { clipboard, nativeImage } from 'electron'

export const ELEMENT_CLIPBOARD_FORMAT = 'io.genoffice.slides.elements'

export type ElementClipboardIdentity = { token: string; senderId: number }

const MAX_TOKEN_LENGTH = 256
const MAX_PNG_BASE64_LENGTH = 40 * 1024 * 1024

export function isElementClipboardToken(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= MAX_TOKEN_LENGTH
}

function markerBufferMatches(token: string): boolean {
  try {
    return clipboard.readBuffer(ELEMENT_CLIPBOARD_FORMAT).equals(Buffer.from(token))
  } catch {
    return false
  }
}

function markerHtmlMatches(token: string): boolean {
  try {
    const encodedToken = encodeURIComponent(token)
    const images = clipboard
      .readHTML()
      .matchAll(/<img\b[^>]*\bdata-genoffice-slides-elements\s*=\s*"([^"]*)"[^>]*>/gi)
    return [...images].some((match) => match[1] === encodedToken)
  } catch {
    return false
  }
}

/** Whether the OS clipboard still belongs to the supplied element-copy token. */
export function elementClipboardMarkerMatches(token: string): boolean {
  return isElementClipboardToken(token) && (markerBufferMatches(token) || markerHtmlMatches(token))
}

/** Prevent a delayed renderer PNG from replacing a later copy or an external clipboard change. */
export function canWriteElementClipboardImage(
  current: ElementClipboardIdentity | null,
  senderId: number,
  token: unknown,
): token is string {
  return (
    isElementClipboardToken(token) &&
    current?.senderId === senderId &&
    current.token === token &&
    elementClipboardMarkerMatches(token)
  )
}

function decodePngBase64(pngBase64: unknown): Buffer | null {
  if (
    typeof pngBase64 !== 'string' ||
    pngBase64.length === 0 ||
    pngBase64.length > MAX_PNG_BASE64_LENGTH ||
    pngBase64.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(pngBase64)
  ) {
    return null
  }
  const bytes = Buffer.from(pngBase64, 'base64')
  return bytes.length ? bytes : null
}

/** Writes a standard image and an HTML marker together so external apps can paste the image. */
export function writeElementClipboardImage(token: string, pngBase64: unknown): boolean {
  if (!isElementClipboardToken(token)) return false
  const bytes = decodePngBase64(pngBase64)
  if (!bytes) return false
  try {
    const decoded = nativeImage.createFromBuffer(bytes)
    if (decoded.isEmpty()) return false
    const normalizedPng = decoded.toPNG()
    if (!normalizedPng.length) return false
    const html = `<img src="data:image/png;base64,${normalizedPng.toString('base64')}" data-genoffice-slides-elements="${encodeURIComponent(token)}">`
    if (!elementClipboardMarkerMatches(token)) return false
    clipboard.write({ image: decoded, html })
    return true
  } catch {
    return false
  }
}
