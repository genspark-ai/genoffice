import * as CFB from 'cfb'
import { LegacyPptError } from './errors'

/** Caps mirror the other legacy importers: fail fast, never walk gigabytes. */
export const MAX_INPUT_BYTES = 256 * 1024 * 1024
export const MAX_STREAM_BYTES = 128 * 1024 * 1024

/** Open a legacy .ppt container: OLE2/CFB validated, the presentation streams out. */
export function openPptContainer(bytes: Uint8Array): {
  document: Uint8Array
  currentUser: Uint8Array | null
} {
  if (bytes.length > MAX_INPUT_BYTES) {
    throw new LegacyPptError('limits_exceeded', `input ${bytes.length} bytes`)
  }
  if (bytes.length < 8 || bytes[0] !== 0xd0 || bytes[1] !== 0xcf) {
    throw new LegacyPptError('not_cfb')
  }
  let cfb: CFB.CFB$Container
  try {
    cfb = CFB.read(bytes as unknown as Buffer, { type: 'buffer' })
  } catch (err) {
    throw new LegacyPptError('not_cfb', err instanceof Error ? err.message : undefined)
  }
  const read = (name: string): Uint8Array | null => {
    const entry = CFB.find(cfb, name)
    // cfb hands the stream back as a Buffer for big streams and a plain
    // number array for small ones — both are array-likes of bytes
    const content = entry?.content as ArrayLike<number> | undefined
    const size = content?.length ?? 0
    if (size === 0) return null
    if (size > MAX_STREAM_BYTES) {
      throw new LegacyPptError('limits_exceeded', `stream ${name} ${size} bytes`)
    }
    const out = new Uint8Array(size)
    for (let i = 0; i < size; i++) out[i] = content![i]!
    return out
  }
  const document = read('PowerPoint Document')
  if (!document) throw new LegacyPptError('missing_stream')
  return { document, currentUser: read('Current User') }
}
