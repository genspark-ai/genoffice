import { basename, extname } from 'node:path'
import { stat } from 'node:fs/promises'
import { parseFileToText } from '@genoffice/file-parse'

/**
 * The read-only half of what the mod assistant may do with a document: its text,
 * and nothing else. The extension allowlist is what keeps "read a document" from
 * becoming "read any file on the machine", and the byte cap keeps one call from
 * pulling a huge file into the model's context.
 *
 * There is deliberately no writer anywhere in this module, or in the registry it
 * ships with: the assistant cannot change a document in this build.
 */

const DOCUMENT_EXTS = new Set([
  'docx',
  'docm',
  'dotx',
  'dotm',
  'xlsx',
  'xlsm',
  'xltx',
  'pptx',
  'pptm',
  'potx',
  'pdf',
  'md',
  'markdown',
  'html',
  'htm',
  'txt',
  'csv',
  'rtf',
  'odt',
  'ods',
  'odp',
])

export const DOCUMENT_MAX_BYTES = 50 * 1024 * 1024
export const DOCUMENT_MAX_CHARS = 200_000

export interface ReadDocumentResult {
  name: string
  text: string
  chars: number
}

/** Parsed text of one document file; throws with a plain message a tool can show. */
/** Whether this file name is one the reader can open (the listing uses it too). */
export function isReadableDocument(name: string): boolean {
  return DOCUMENT_EXTS.has(extname(name).slice(1).toLowerCase())
}

export async function readDocumentText(path: string): Promise<ReadDocumentResult> {
  const ext = extname(path).slice(1).toLowerCase()
  if (!DOCUMENT_EXTS.has(ext)) {
    throw new Error(
      'not a document: ' + (ext ? '.' + ext + ' files are not readable here' : 'no file extension'),
    )
  }
  const info = await stat(path)
  if (!info.isFile()) throw new Error('not a file: ' + path)
  if (info.size > DOCUMENT_MAX_BYTES) {
    throw new Error(
      'file too large to read (' +
        Math.round(info.size / 1024 / 1024) +
        'MB, cap ' +
        DOCUMENT_MAX_BYTES / 1024 / 1024 +
        'MB)',
    )
  }
  const parsed = await parseFileToText(path)
  if (!parsed.ok || parsed.kind !== 'text' || parsed.text == null) {
    throw new Error(parsed.error ?? 'unreadable file')
  }
  const text = parsed.text.slice(0, DOCUMENT_MAX_CHARS)
  return { name: basename(path), text, chars: text.length }
}
