import { readFile } from 'node:fs/promises'
import { extname } from 'node:path'
import { docToText } from './doc'
import { docxToText } from './docx'
import { pdfToText } from './pdf'
import { pptToText } from './ppt'
import { pptxToText } from './pptx'
import { xlsxToText } from './xlsx'

export type ParsedFileKind = 'text' | 'image' | 'unsupported'

export interface ParsedFile {
  ok: boolean
  text?: string
  kind: ParsedFileKind
  mime?: string
  error?: string
}

/** No text extraction for images: callers read raw bytes and go multimodal (see @genoffice/ai-provider images support) */
const IMAGE_MIMES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
}

const TEXT_EXTS = new Set([
  'txt',
  'md',
  'markdown',
  'csv',
  'tsv',
  'json',
  'xml',
  'html',
  'htm',
  'log',
  'py',
])

const FATAL_UTF8_DECODER = new TextDecoder('utf-8', { fatal: true })
// windows-1252 (a latin-1 superset) maps every byte to a character, so the
// fallback below can never fail (full ICU, Node >= 14, like 'utf-16be' above)
const WINDOWS_1252_DECODER = new TextDecoder('windows-1252')

/**
 * Decode plain-text bytes honouring Unicode BOMs: UTF-8 (BOM stripped), UTF-16LE
 * and UTF-16BE are decoded. Bytes that are not valid UTF-8 (latin-1, GBK,
 * Shift-JIS, a stray invalid byte in otherwise-valid UTF-8, ...) fall back to a
 * windows-1252 decode instead of being rejected: callers index whatever survives
 * rather than drop the whole attachment, and a literal U+FFFD in valid UTF-8 is
 * never mistaken for a decode error (the fatal decoder tells them apart).
 * Returns null only for a BOM-declared UTF-32 file, whose declared encoding we
 * cannot decode at all.
 */
function decodeTextBytes(bytes: Buffer): string | null {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    if (bytes[2] === 0 && bytes[3] === 0) return null // UTF-32LE BOM
    return bytes.toString('utf16le', 2)
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    // Node buffers have no 'utf-16be'; TextDecoder (full ICU, Node >= 14) does
    return new TextDecoder('utf-16be').decode(bytes.subarray(2))
  }
  const bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
  const payload = bom ? bytes.subarray(3) : bytes
  try {
    return FATAL_UTF8_DECODER.decode(payload)
  } catch {
    return WINDOWS_1252_DECODER.decode(payload)
  }
}

/** parse an attachment into plain text (or flag it as image / unsupported) */
export async function parseFileToText(filePath: string): Promise<ParsedFile> {
  const ext = extname(filePath).slice(1).toLowerCase()
  const imageMime = IMAGE_MIMES[ext]
  if (imageMime) return { ok: true, kind: 'image', mime: imageMime }
  try {
    if (TEXT_EXTS.has(ext)) {
      const text = decodeTextBytes(await readFile(filePath))
      if (text === null) {
        return {
          ok: false,
          kind: 'text',
          error:
            'Unsupported text encoding: BOM-declared UTF-32 cannot be decoded (UTF-8 and UTF-16 are supported)',
        }
      }
      return { ok: true, kind: 'text', text }
    }
    switch (ext) {
      case 'doc':
        return { ok: true, kind: 'text', text: await docToText(await readFile(filePath)) }
      case 'docx':
        return { ok: true, kind: 'text', text: await docxToText(await readFile(filePath)) }
      case 'ppt':
        return { ok: true, kind: 'text', text: await pptToText(await readFile(filePath)) }
      case 'pptx':
        return { ok: true, kind: 'text', text: await pptxToText(await readFile(filePath)) }
      case 'xlsx':
      case 'xlsm':
        return { ok: true, kind: 'text', text: await xlsxToText(await readFile(filePath)) }
      case 'pdf':
        return { ok: true, kind: 'text', text: await pdfToText(await readFile(filePath)) }
    }
  } catch (e) {
    return { ok: false, kind: 'text', error: e instanceof Error ? e.message : String(e) }
  }
  return { ok: false, kind: 'unsupported', error: `Unsupported file type: .${ext || 'unknown'}` }
}
