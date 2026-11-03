import {
  CompoundFileError,
  openCompoundStreams,
  type CompoundFileLimits as SharedCompoundFileLimits,
} from '@genoffice/compound-file'

import { LegacyDocError } from './errors'

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

/** The top-level streams of an OLE2 container (OLE2 calls them storages/streams). */
export interface CompoundFile {
  /** raw bytes of a stream by name (e.g. "WordDocument"), or undefined when absent */
  stream(name: string): Uint8Array | undefined
  /** names of the streams that carry document data (producer markers filtered out) */
  streams(): string[]
}

const WORD_DOCUMENT_STREAM = 'WordDocument'

/**
 * Open a .doc's OLE2 container. The generic OLE2 rules (signature, header,
 * caps, owned streams) live in @genoffice/compound-file, shared with the .xls
 * reader; this wrapper maps its failures onto LegacyDocError and requires the
 * WordDocument stream (the FIB + text of every Word 97-2003 document).
 */
export function openCompoundFile(
  bytes: Uint8Array,
  limits: CompoundFileLimits = LEGACY_DOC_LIMITS,
): CompoundFile {
  let opened
  try {
    opened = openCompoundStreams(bytes, limits satisfies SharedCompoundFileLimits)
  } catch (error) {
    if (error instanceof CompoundFileError) {
      throw new LegacyDocError(error.code, error.detail ?? error.message)
    }
    throw error
  }
  if (!opened.stream(WORD_DOCUMENT_STREAM)) {
    throw new LegacyDocError(
      'missing_stream',
      'the OLE2 container has no ' + WORD_DOCUMENT_STREAM + ' stream',
    )
  }
  return {
    stream: (name) => opened.stream(name),
    streams: () => [...opened.names()],
  }
}
