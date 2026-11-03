import * as CFB from 'cfb'

/**
 * One OLE2/CFB opener for the legacy Office importers. Word and Excel both
 * store their data in compound files; the container rules (signature, 512-byte
 * header, caps, owned stream bytes) are format-independent, so doc2docx and
 * xls2xlsx share them instead of each carrying a weaker copy.
 */

export interface CompoundFileLimits {
  /** whole input file */
  maxBytes: number
  /** one decompressed stream */
  maxStreamBytes: number
}

export type CompoundFileErrorCode = 'not_cfb' | 'truncated' | 'limits_exceeded' | 'corrupt'

export class CompoundFileError extends Error {
  readonly code: CompoundFileErrorCode
  readonly detail?: string

  constructor(code: CompoundFileErrorCode, detail?: string) {
    super(detail ?? code)
    this.name = 'CompoundFileError'
    this.code = code
    this.detail = detail
  }
}

export interface CompoundStreams {
  /** raw bytes of a stream by its root-relative name, or undefined when absent */
  stream(name: string): Uint8Array | undefined
  /** names of the data streams (the cfb producer marker is filtered out), sorted */
  names(): readonly string[]
}

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
const OLE2_HEADER_BYTES = 512

function hasCfbMagic(bytes: Uint8Array): boolean {
  if (bytes.length < CFB_MAGIC.length) return false
  return CFB_MAGIC.every((value, index) => bytes[index] === value)
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

/** cfb seeds every new container with this producer marker; it carries no document data */
function isProducerStream(name: string): boolean {
  return name.startsWith('\u0001')
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

/**
 * Validate and open an OLE2/CFB container. The signature and header are checked
 * before cfb parses the file, every stream is capped while it is copied out, and
 * failures surface as a CompoundFileError so each format can map it onto its own
 * user-facing error model. The caller decides which streams a document needs.
 */
export function openCompoundStreams(
  bytes: Uint8Array,
  limits: CompoundFileLimits,
): CompoundStreams {
  if (bytes.length === 0) throw new CompoundFileError('not_cfb', 'the file is empty')
  if (bytes.length > limits.maxBytes) {
    throw new CompoundFileError(
      'limits_exceeded',
      `the file is ${bytes.length} bytes, over the ${limits.maxBytes}-byte limit`,
    )
  }
  if (!hasCfbMagic(bytes)) {
    throw new CompoundFileError(
      'not_cfb',
      'the file does not start with the OLE2 compound-document signature',
    )
  }
  if (bytes.length < OLE2_HEADER_BYTES) {
    throw new CompoundFileError(
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
    throw new CompoundFileError(
      'corrupt',
      `the OLE2 container cannot be read (${errorText(error)})`,
    )
  }

  const streams = new Map<string, Uint8Array>()
  const paths = container.FullPaths ?? []
  for (let index = 0; index < paths.length; index++) {
    const entry = container.FileIndex[index]
    if (!entry || entry.type !== 2) continue
    const name = streamNameOf(paths[index] ?? '', entry.name)
    if (!name || isProducerStream(name) || streams.has(name)) continue
    const content = toOwnedBytes(entry.content)
    if (content.length > limits.maxStreamBytes) {
      throw new CompoundFileError(
        'limits_exceeded',
        `the ${name} stream is ${content.length} bytes, over the ${limits.maxStreamBytes}-byte limit`,
      )
    }
    streams.set(name, content)
  }

  return {
    stream: (name) => streams.get(name),
    names: () => [...streams.keys()].sort(),
  }
}
