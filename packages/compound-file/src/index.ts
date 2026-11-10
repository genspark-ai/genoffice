import * as CFB from 'cfb'

/**
 * The OLE2/CFB container opener shared by the legacy Office importers (the
 * .doc reader in @genoffice/doc2docx, the .ppt reader in @genoffice/ppt2pptx,
 * and any .xls reader that comes along).
 *
 * It validates the compound-file signature and the 512-byte header before
 * cfb parses anything, caps the input and every stream, and hands back owned
 * stream bytes so a later mutation of the container cannot change a parsed
 * part. Callers layer their own format checks (WordDocument / PowerPoint
 * Document / Workbook) and map the failure codes onto their own error model.
 */

/** An OLE2/CFB compound document starts with this 8-byte signature. */
export const OLE2_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const
/** A compound file is at least its 512-byte header. */
export const OLE2_HEADER_BYTES = 512

/** The size caps of the shared opener; each legacy importer picks its own. */
export interface Ole2Limits {
  /** whole input file */
  maxBytes: number
  /** one decompressed stream */
  maxStreamBytes: number
}

/** The top-level streams of an OLE2 container (OLE2 calls them storages/streams). */
export interface Ole2Streams {
  /** raw bytes of a stream by name (e.g. "WordDocument"), or undefined when absent */
  stream(name: string): Uint8Array | undefined
  /** names of the streams that carry document data (producer markers filtered out) */
  streams(): string[]
}

/** Stable failure codes of the opener; callers map them onto their own error type. */
export type Ole2FailureCode = 'not_cfb' | 'truncated' | 'limits_exceeded' | 'corrupt'

export type Ole2OpenResult =
  | { readonly ok: true; readonly container: Ole2Streams }
  | { readonly ok: false; readonly code: Ole2FailureCode; readonly detail: string }

function failure(code: Ole2FailureCode, detail: string): Ole2OpenResult {
  return { ok: false, code, detail }
}

export function hasOle2Signature(bytes: Uint8Array): boolean {
  if (bytes.length < OLE2_MAGIC.length) return false
  return OLE2_MAGIC.every((value, index) => bytes[index] === value)
}

/** the stream's own bytes, so a later mutation of the CFB container cannot change a parsed part */
function toOwnedBytes(content: unknown): Uint8Array {
  if (content instanceof Uint8Array) return new Uint8Array(content)
  if (Array.isArray(content)) return Uint8Array.from(content as number[])
  // cfb always stores Array<number> / Uint8Array for stream content
  return new Uint8Array(0)
}

/** stream name relative to the root entry ("Root Entry/WordDocument" -> "WordDocument") */
function streamNameOf(path: string, fallback: string): string {
  const relative = path.startsWith('Root Entry/')
    ? path.slice('Root Entry/'.length)
    : path.replace(/^\/+/, '')
  return relative || fallback
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

/** cfb seeds every new container with this producer marker; it carries no document data */
const cfbProducerStream = (name: string): boolean => name.startsWith('\u0001')

/**
 * Open an OLE2 compound file and expose its document streams.
 *
 * Validates the 8-byte signature and the 512-byte header, caps the input and
 * every stored stream, and skips cfb's producer-marker streams. A file that
 * passes this check is a real compound document, not an OOXML part renamed to
 * a legacy extension.
 */
export function openOle2Container(bytes: Uint8Array, limits: Ole2Limits): Ole2OpenResult {
  if (bytes.length === 0) return failure('not_cfb', 'the file is empty')
  if (bytes.length > limits.maxBytes) {
    return failure(
      'limits_exceeded',
      `the file is ${bytes.length} bytes, over the ${limits.maxBytes}-byte limit`,
    )
  }
  if (!hasOle2Signature(bytes)) {
    return failure('not_cfb', 'the file does not start with the OLE2 compound-document signature')
  }
  if (bytes.length < OLE2_HEADER_BYTES) {
    return failure(
      'truncated',
      `the file is ${bytes.length} bytes, shorter than the 512-byte OLE2 header`,
    )
  }

  let container: CFB.CFB$Container
  try {
    // cfb's own reader: 'buffer' input is parsed directly (it would otherwise
    // treat a Uint8Array as a base64 string)
    container = CFB.read(bytes, { type: 'buffer' })
  } catch (error) {
    return failure('corrupt', `the OLE2 container cannot be read (${errorText(error)})`)
  }

  const streams = new Map<string, Uint8Array>()
  const paths = container.FullPaths ?? []
  for (let index = 0; index < paths.length; index++) {
    const entry = container.FileIndex[index]
    if (!entry || entry.type !== 2) continue
    const name = streamNameOf(paths[index] ?? '', entry.name)
    if (!name || cfbProducerStream(name) || streams.has(name)) continue
    const content = toOwnedBytes(entry.content)
    if (content.length > limits.maxStreamBytes) {
      return failure(
        'limits_exceeded',
        `the ${name} stream is ${content.length} bytes, over the ${limits.maxStreamBytes}-byte limit`,
      )
    }
    streams.set(name, content)
  }

  return {
    ok: true,
    container: {
      stream: (name) => streams.get(name),
      streams: () => [...streams.keys()].sort(),
    },
  }
}
