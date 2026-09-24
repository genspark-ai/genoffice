import { gzipSync } from 'node:zlib'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { MAX_METAFILE_GUNZIP_BYTES, metafileToDataUrl } from '../src/metafile'

// The real converter needs a canvas API this environment does not have. It is
// replaced by one that reports the size of what it received, so the assertions
// below see only whether the bytes reached it inflated.
vi.mock('../src/vendor/emf-converter/index.mjs', () => ({
  convertEmfToDataUrl: vi.fn(
    async (buffer: ArrayBuffer) => `data:image/png;base64,${buffer.byteLength}`,
  ),
  convertWmfToDataUrl: vi.fn(async () => null),
}))

/** Bytes that only have to survive the trip: the converter is mocked. */
const PAYLOAD = Buffer.from('an EMF in the real world; any bytes here')

/** jsdom's Blob has no stream(); this is what that looks like from in here. */
function withoutBlobStream(): () => void {
  const proto = Blob.prototype as unknown as { stream?: () => unknown }
  const descriptor = Object.getOwnPropertyDescriptor(proto, 'stream')
  Object.defineProperty(proto, 'stream', { value: undefined, configurable: true, writable: true })
  return () => {
    if (descriptor) Object.defineProperty(proto, 'stream', descriptor)
    else delete proto.stream
  }
}

describe('metafile gunzip (#798)', () => {
  let restore: () => void
  beforeEach(() => {
    restore = withoutBlobStream()
  })
  afterEach(() => {
    restore()
    vi.restoreAllMocks()
  })

  it('decodes a gzipped metafile without Blob.prototype.stream', async () => {
    // Before: gunzip threw "(intermediate value).stream is not a function",
    // metafileToDataUrl caught it and returned null, and the image silently
    // rendered as an empty frame instead of an unsupported-format error.
    const emz = gzipSync(PAYLOAD)
    await expect(metafileToDataUrl(emz, 'image/x-emz')).resolves.toBe(
      `data:image/png;base64,${PAYLOAD.length}`,
    )
  })

  it('still refuses a gzip bomb through the same path', async () => {
    // The 64 MB cap is the reason gunzip streams at all; it has to keep working
    // on the stream this builds by hand, in the environment the cap protects.
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const bomb = gzipSync(Buffer.alloc(MAX_METAFILE_GUNZIP_BYTES + 1024))
    await expect(metafileToDataUrl(bomb, 'image/x-emz')).resolves.toBeNull()
    expect(warn.mock.calls.flat().join(' ')).toContain('gunzip output exceeds')
  })
})
