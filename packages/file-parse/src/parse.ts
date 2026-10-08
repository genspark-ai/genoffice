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
// windows-1252 maps every byte, so this fallback can never fail
const WINDOWS_1252_DECODER = new TextDecoder('windows-1252')

/**
 * BOM-aware text decode; invalid UTF-8 falls back to windows-1252 so callers
 * index what survives. null only for a BOM-declared UTF-32 file.
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

type BinaryKind = 'zip' | 'pdf' | 'ole2' | 'unknown'

/** pdf.js scans the first 1024 bytes for the header, and leading junk before it
 *  (HTTP header remnants, whitespace) is common in the wild, so match the header
 *  where a real parser does rather than pinning it to bytes 0..3. */
const PDF_HEADER_WINDOW = 1024

function sniffBinary(bytes: Buffer): BinaryKind {
  if (bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b) return 'zip'
  if (bytes.subarray(0, PDF_HEADER_WINDOW).includes(Buffer.from('%PDF-'))) return 'pdf'
  if (
    bytes.length >= 4 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0
  )
    return 'ole2'
  return 'unknown'
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
    const bytes = await readFile(filePath)
    const sniffed = sniffBinary(bytes)
    if (ext === 'pdf' && sniffed !== 'pdf') {
      return {
        ok: false,
        kind: 'text',
        error: `Content mismatch: .pdf file has ${sniffed} magic bytes`,
      }
    }
    if (
      (ext === 'docx' || ext === 'pptx' || ext === 'xlsx' || ext === 'xlsm') &&
      sniffed !== 'zip'
    ) {
      return {
        ok: false,
        kind: 'text',
        error: `Content mismatch: .${ext} file has ${sniffed} magic bytes`,
      }
    }
    // a DOCX saved under a .doc name is common enough that it parses as OOXML
    if (
      (ext === 'doc' && sniffed !== 'ole2' && sniffed !== 'zip') ||
      (ext === 'ppt' && sniffed !== 'ole2')
    ) {
      return {
        ok: false,
        kind: 'text',
        error: `Content mismatch: .${ext} file has ${sniffed} magic bytes`,
      }
    }
    switch (ext) {
      case 'doc':
        return {
          ok: true,
          kind: 'text',
          text: sniffed === 'zip' ? await docxToText(bytes) : await docToText(bytes),
        }
      case 'docx':
        return { ok: true, kind: 'text', text: await docxToText(bytes) }
      case 'ppt':
        return { ok: true, kind: 'text', text: await pptToText(bytes) }
      case 'pptx':
        return { ok: true, kind: 'text', text: await pptxToText(bytes) }
      case 'xlsx':
      case 'xlsm':
        return { ok: true, kind: 'text', text: await xlsxToText(bytes) }
      case 'pdf':
        return { ok: true, kind: 'text', text: await pdfToText(bytes) }
    }
  } catch (e) {
    return { ok: false, kind: 'text', error: e instanceof Error ? e.message : String(e) }
  }
  return { ok: false, kind: 'unsupported', error: `Unsupported file type: .${ext || 'unknown'}` }
}
