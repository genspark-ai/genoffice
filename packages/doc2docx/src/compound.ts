import {
  OLE2_HEADER_BYTES,
  OLE2_MAGIC,
  hasOle2Signature,
  openOle2Container as openSharedOle2Container,
  type Ole2FailureCode,
  type Ole2OpenResult,
} from '@genoffice/compound-file'

import { LegacyDocError } from './errors'

export { OLE2_HEADER_BYTES, OLE2_MAGIC, hasOle2Signature }
export type { Ole2FailureCode, Ole2OpenResult }

/**
 * Upper bounds for the legacy import. A filing-size .doc is a few megabytes;
 * the caps keep a corrupt or hostile file from exhausting memory before the
 * format checks can reject it.
 */
export const LEGACY_DOC_LIMITS = {
  /** whole input file */
  maxBytes: 256 * 1024 * 1024,
  /** one decompressed stream (WordDocument / 0Table / 1Table) */
  maxStreamBytes: 128 * 1024 * 1024,
} as const

export interface CompoundFileLimits {
  maxBytes: number
  maxStreamBytes: number
}

/** The size caps of the shared opener; each legacy importer picks its own. */
export type Ole2Limits = CompoundFileLimits

/** The top-level streams of an OLE2 container (OLE2 calls them storages/streams). */
export interface CompoundFile {
  /** raw bytes of a stream by name (e.g. "WordDocument"), or undefined when absent */
  stream(name: string): Uint8Array | undefined
  /** names of the streams that carry document data (producer markers filtered out) */
  streams(): string[]
}

export type Ole2Streams = CompoundFile

export const WORD_DOCUMENT_STREAM = 'WordDocument'

/**
 * Open an OLE2 compound file and expose its document streams. The generic
 * rules (signature, 512-byte header, input and stream caps, owned stream
 * bytes) live in @genoffice/compound-file, shared with the other legacy
 * importers; this wrapper defaults to the .doc caps.
 */
export function openOle2Container(
  bytes: Uint8Array,
  limits: Ole2Limits = LEGACY_DOC_LIMITS,
): Ole2OpenResult {
  return openSharedOle2Container(bytes, limits)
}

/**
 * Open a .doc's OLE2 container. Validates the CFB signature and header size,
 * caps the input and every stored stream, and requires the WordDocument
 * stream (the FIB + text of every Word 97-2003 document).
 */
export function openCompoundFile(
  bytes: Uint8Array,
  limits: CompoundFileLimits = LEGACY_DOC_LIMITS,
): CompoundFile {
  const opened = openOle2Container(bytes, limits)
  if (!opened.ok) throw new LegacyDocError(opened.code, opened.detail)
  if (!opened.container.stream(WORD_DOCUMENT_STREAM)) {
    throw new LegacyDocError(
      'missing_stream',
      `the OLE2 container has no ${WORD_DOCUMENT_STREAM} stream`,
    )
  }
  return opened.container
}
