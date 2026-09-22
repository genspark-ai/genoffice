/**
 * w:altChunk import (display only). The chunk part stays untouched in the
 * package and is re-emitted verbatim on save; the host converts HTML / MHT
 * to a docx with the html2docx chain (needs a browser), which the parser
 * then reads like any other document. A .docx chunk is parsed directly.
 */
import type JSZip from 'jszip'
import type { RelInfo } from './parse-xml-text'

export type AltChunkHtmlConverter = (html: string) => Promise<Uint8Array | null>

let htmlConverter: AltChunkHtmlConverter | null = null

export function setAltChunkHtmlConverter(fn: AltChunkHtmlConverter | null): void {
  htmlConverter = fn
}

/** an HTML/MHT chunk can only expand where the host installed a converter */
export function hasAltChunkHtmlConverter(): boolean {
  return htmlConverter !== null
}

export type AltChunkKind = 'html' | 'mht' | 'docx'

const ALT_CHUNK_REL = /\/aFChunk$/
const MAX_CHUNK_BYTES = 64 * 1024 * 1024

export function altChunkKind(
  path: string,
  contentType: string | undefined,
  head: Uint8Array,
): AltChunkKind | null {
  if (head.length >= 4 && head[0] === 0x50 && head[1] === 0x4b && head[2] === 3 && head[3] === 4) {
    return 'docx'
  }
  const ct = contentType?.split(';')[0].trim().toLowerCase()
  if (ct === 'text/html' || ct === 'application/xhtml+xml') return 'html'
  if (ct === 'message/rfc822' || ct === 'multipart/related') return 'mht'
  if (ct === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    return 'docx'
  const ext = path.split('.').pop()?.toLowerCase() ?? ''
  if (ext === 'htm' || ext === 'html' || ext === 'xhtml') return 'html'
  if (ext === 'mht' || ext === 'mhtml' || ext === 'eml') return 'mht'
  if (ext === 'docx' || ext === 'docm' || ext === 'dotx') return 'docx'
  const text = latin1(head)
  if (/^(MIME-Version|From|Subject|Content-Type):/i.test(text.trimStart())) return 'mht'
  if (/<!doctype\s+html|<html[\s>]/i.test(text)) return 'html'
  return null
}

/** bytes as Latin-1 text (header / markup sniffing, transfer-decoding input) */
function latin1(bytes: Uint8Array, limit = bytes.length): string {
  let s = ''
  const n = Math.min(bytes.length, limit)
  for (let i = 0; i < n; i += 8192) {
    s += String.fromCharCode(...bytes.subarray(i, Math.min(i + 8192, n)))
  }
  return s
}

function decodeWithCharset(bytes: Uint8Array, charset: string | undefined): string {
  const label = (charset ?? '')
    .trim()
    .replace(/^["']|["']$/g, '')
    .toLowerCase()
  for (const candidate of [label, 'utf-8']) {
    if (!candidate) continue
    try {
      return new TextDecoder(candidate, { fatal: candidate === 'utf-8' }).decode(bytes)
    } catch {
      /* unknown label or invalid utf-8: try the next */
    }
  }
  return new TextDecoder('windows-1252').decode(bytes)
}

/** HTML bytes -> string: BOM, then <meta charset>, then utf-8 with a 1252 fallback */
export function decodeHtmlBytes(bytes: Uint8Array): string {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3))
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(bytes.subarray(2))
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(bytes.subarray(2))
  }
  const head = latin1(bytes, 4096)
  const meta =
    /<meta[^>]+charset\s*=\s*["']?\s*([\w-]+)/i.exec(head)?.[1] ??
    /<\?xml[^>]+encoding\s*=\s*["']([\w-]+)/i.exec(head)?.[1]
  return decodeWithCharset(bytes, meta)
}

interface MimePart {
  headers: Map<string, string>
  body: Uint8Array
}

function parseMimeHeaders(raw: string): Map<string, string> {
  const headers = new Map<string, string>()
  // folded continuation lines start with whitespace
  for (const line of raw.replace(/\r?\n[ \t]+/g, ' ').split(/\r?\n/)) {
    const colon = line.indexOf(':')
    if (colon <= 0) continue
    headers.set(line.slice(0, colon).trim().toLowerCase(), line.slice(colon + 1).trim())
  }
  return headers
}

function headerParam(value: string | undefined, name: string): string | undefined {
  if (!value) return undefined
  const m = new RegExp(`${name}\\s*=\\s*(?:"([^"]*)"|([^;\\s]+))`, 'i').exec(value)
  return m ? (m[1] ?? m[2]) : undefined
}

function splitHeadersBody(bytes: Uint8Array): { headers: Map<string, string>; body: Uint8Array } {
  const text = latin1(bytes, 65536)
  const m = /\r?\n\r?\n/.exec(text)
  if (!m) return { headers: parseMimeHeaders(text), body: new Uint8Array(0) }
  return {
    headers: parseMimeHeaders(text.slice(0, m.index)),
    body: bytes.subarray(m.index + m[0].length),
  }
}

function indexOfBytes(hay: Uint8Array, needle: Uint8Array, from: number): number {
  outer: for (let i = from; i <= hay.length - needle.length; i++) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer
    return i
  }
  return -1
}

function splitMultipart(body: Uint8Array, boundary: string): MimePart[] {
  const parts: MimePart[] = []
  const marker = new TextEncoder().encode(`--${boundary}`)
  let pos = indexOfBytes(body, marker, 0)
  while (pos !== -1) {
    let start = pos + marker.length
    if (body[start] === 0x2d && body[start + 1] === 0x2d) break
    if (body[start] === 0x0d) start++
    if (body[start] === 0x0a) start++
    const next = indexOfBytes(body, marker, start)
    let end = next === -1 ? body.length : next
    if (body[end - 1] === 0x0a) end--
    if (body[end - 1] === 0x0d) end--
    parts.push(splitHeadersBody(body.subarray(start, Math.max(start, end))))
    pos = next
  }
  return parts
}

export function decodeQuotedPrintable(text: string): Uint8Array {
  const out: number[] = []
  const src = text.replace(/=\r?\n/g, '')
  for (let i = 0; i < src.length; i++) {
    const c = src.charCodeAt(i)
    if (c === 0x3d && i + 2 < src.length) {
      const hex = src.slice(i + 1, i + 3)
      if (/^[0-9A-Fa-f]{2}$/.test(hex)) {
        out.push(parseInt(hex, 16))
        i += 2
        continue
      }
    }
    out.push(c & 0xff)
  }
  return Uint8Array.from(out)
}

function decodeBase64(text: string): Uint8Array {
  const clean = text.replace(/[^A-Za-z0-9+/=]/g, '')
  const bin = atob(clean)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}

function decodeTransfer(part: MimePart): Uint8Array {
  const enc = (part.headers.get('content-transfer-encoding') ?? '').toLowerCase()
  if (enc === 'quoted-printable') return decodeQuotedPrintable(latin1(part.body))
  if (enc === 'base64') return decodeBase64(latin1(part.body))
  return part.body
}

/**
 * MHT (Word "Single File Web Page") -> the HTML document string. Referenced
 * image parts are inlined as data URLs by Content-Location / Content-ID so
 * the converter sees them without the package.
 */
export function decodeMhtToHtml(bytes: Uint8Array): string | null {
  const top = splitHeadersBody(bytes)
  const type = top.headers.get('content-type') ?? ''
  const boundary = headerParam(type, 'boundary')
  const parts: MimePart[] = boundary ? splitMultipart(top.body, boundary) : [top]
  const htmlPart = parts.find((p) =>
    /^text\/html/i.test((p.headers.get('content-type') ?? '').trim()),
  )
  if (!htmlPart) return null
  const htmlBytes = decodeTransfer(htmlPart)
  const charset = headerParam(htmlPart.headers.get('content-type'), 'charset')
  const html = charset ? decodeWithCharset(htmlBytes, charset) : decodeHtmlBytes(htmlBytes)
  const baseLocation = htmlPart.headers.get('content-location') ?? ''
  // one map + one pass over the html: per-part regexes made this parts x html
  const inline = new Map<string, string>()
  for (const part of parts) {
    if (part === htmlPart) continue
    const ct = (part.headers.get('content-type') ?? '').split(';')[0].trim().toLowerCase()
    if (!ct.startsWith('image/')) continue
    const dataUrl = `data:${ct};base64,${btoa(latin1(decodeTransfer(part)))}`
    const location = part.headers.get('content-location')
    if (location) {
      inline.set(location.toLowerCase(), dataUrl)
      if (baseLocation) {
        const base = baseLocation.replace(/[^/]*$/, '')
        if (location.startsWith(base))
          inline.set(location.slice(base.length).toLowerCase(), dataUrl)
      }
    }
    const cid = part.headers.get('content-id')?.replace(/^<|>$/g, '')
    if (cid) inline.set(`cid:${cid}`.toLowerCase(), dataUrl)
  }
  if (inline.size === 0) return html
  return html.replace(
    /\b(src|href)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+))/gi,
    (m, attr: string, dq?: string, sq?: string, uq?: string) => {
      const dataUrl = inline.get((dq ?? sq ?? uq ?? '').toLowerCase())
      if (!dataUrl) return m
      const q = sq !== undefined ? "'" : '"'
      return `${attr}=${q}${dataUrl}${q}`
    },
  )
}

export function altChunkPartPath(rels: Map<string, RelInfo>, rId: string): string | null {
  const rel = rels.get(rId)
  if (!rel || rel.targetMode === 'External' || !ALT_CHUNK_REL.test(rel.type)) return null
  const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`
  return path.replace(/^word\/\.\.\//, '')
}

/**
 * Resolve the chunk to docx bytes the parser can read, or null when the part
 * is missing, of an unsupported type, or no HTML converter is installed.
 */
export async function altChunkToDocx(
  zip: JSZip,
  rels: Map<string, RelInfo>,
  rId: string,
  contentTypeOf: (path: string) => Promise<string | undefined>,
): Promise<Uint8Array | null> {
  const path = altChunkPartPath(rels, rId)
  const file = path ? zip.file(path) : null
  if (!path || !file) return null
  const bytes = await file.async('uint8array')
  if (bytes.length === 0 || bytes.length > MAX_CHUNK_BYTES) return null
  const kind = altChunkKind(path, await contentTypeOf(path), bytes.subarray(0, 4096))
  if (kind === 'docx') return bytes
  if (!kind || !htmlConverter) return null
  const html = kind === 'mht' ? decodeMhtToHtml(bytes) : decodeHtmlBytes(bytes)
  if (!html || !html.trim()) return null
  return htmlConverter(html)
}
