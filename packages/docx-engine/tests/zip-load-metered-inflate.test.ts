import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { crc32, deflateRawSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { assertZipInflatesWithinLimits, loadDocxZip } from '../src/zip-load'
import { writeZip, type ZipEntry } from '../src/zip-splice'

const KB = 1024
const MB = 1024 * KB

/** Small budgets so the tests stay fast; the shape of the check is the point. */
const LIMITS = { maxParts: 100, maxPartBytes: 2 * MB, maxTotalBytes: 8 * MB } as const

function entry(name: string, data: Buffer, usize: number, method = 8): ZipEntry {
  return {
    name,
    nameBytes: Buffer.from(name),
    flags: 0,
    method,
    time: 0,
    date: 0,
    crc: crc32(data),
    csize: data.length,
    usize,
    verMade: 20,
    verNeed: 20,
    intAttr: 0,
    extAttr: 0,
    dataOffset: 0, // writeZip lays the entries out and records the real offsets
  }
}

/** A single part holding `realBytes` of NULs, declaring whatever it likes. */
function onePartArchive(name: string, realBytes: number, declares: number, method = 8): Buffer {
  const truth = Buffer.alloc(realBytes)
  const payload = method === 8 ? deflateRawSync(truth) : truth
  return writeZip([{ meta: entry(name, payload, declares, method), data: payload }])
}

describe('assertZipInflatesWithinLimits', () => {
  it('refuses a part that declares less than it inflates to', async () => {
    // The #759 case: 64 MB behind a 300-byte declaration. Every declared number
    // is inside the limits, so only inflating can find out.
    const archive = onePartArchive('word/document.xml', 64 * MB, 300)
    await expect(assertZipInflatesWithinLimits(archive, LIMITS)).rejects.toThrow(
      /part word\/document\.xml declares 300 uncompressed bytes but inflates past that/,
    )
  })

  it('catches the lie at the declared size, not at the payload', async () => {
    const archive = onePartArchive('word/document.xml', 64 * MB, 300)
    const before = process.memoryUsage().rss
    await expect(assertZipInflatesWithinLimits(archive, LIMITS)).rejects.toThrow(/inflates past/)
    // The budget is one byte past the claim, so nothing near 64 MB is allocated.
    expect(process.memoryUsage().rss - before).toBeLessThan(16 * MB)
  })

  it('bounds the package as a whole, not just each part', async () => {
    // Six honest 1.5 MB parts: each is under the 2 MB per-part cap, so only the
    // 8 MB total can stop it — five fit, the sixth does not.
    const parts = Array.from({ length: 6 }, (_, i) => {
      const payload = deflateRawSync(Buffer.alloc(1536 * KB))
      return { meta: entry(`p${i}.xml`, payload, 1536 * KB), data: payload }
    })
    await expect(assertZipInflatesWithinLimits(writeZip(parts), LIMITS)).rejects.toThrow(
      /total uncompressed size \d+ exceeds the 8388608 limit/,
    )
  })

  it('still rejects an oversized declaration without inflating anything', async () => {
    // 4 MB of real payload declaring 100 MB: the per-part check fires on the
    // claim alone, which is what keeps honest bombs at ~0 ms.
    const archive = onePartArchive('word/document.xml', 4 * MB, 100 * MB)
    await expect(assertZipInflatesWithinLimits(archive, LIMITS)).rejects.toThrow(
      /declares 104857600 uncompressed bytes \(limit 2097152\)/,
    )
  })

  it('counts stored parts against their declared length too', async () => {
    const archive = onePartArchive('word/document.xml', 3 * MB, 0, 0) // stored, declares nothing
    await expect(assertZipInflatesWithinLimits(archive, LIMITS)).rejects.toThrow(
      /declares 0 uncompressed bytes but inflates past that/,
    )
  })

  it('refuses compression methods the gate cannot meter', async () => {
    const archive = onePartArchive('word/document.xml', KB, KB, 9)
    await expect(assertZipInflatesWithinLimits(archive, LIMITS)).rejects.toThrow(
      /unsupported compression method 9/,
    )
  })

  it('accepts an archive that declares more than it delivers', async () => {
    // Padding high is conservative, not a bomb: the total adds the verified
    // length, so an over-claiming part cannot starve the budget of its siblings.
    const truth = Buffer.alloc(64 * KB)
    const payload = deflateRawSync(truth)
    const archive = writeZip([{ meta: entry('word/document.xml', payload, 1 * MB), data: payload }])
    await expect(assertZipInflatesWithinLimits(archive, LIMITS)).resolves.toBeUndefined()
  })

  it('accepts an honest archive and still loads it', async () => {
    const truth = Buffer.from('<?xml version="1.0"?><w:document/>')
    const payload = deflateRawSync(truth)
    const archive = writeZip([
      { meta: entry('word/document.xml', payload, truth.length), data: payload },
    ])
    await expect(assertZipInflatesWithinLimits(archive, LIMITS)).resolves.toBeUndefined()
    const zip = await loadDocxZip(new Uint8Array(archive))
    expect(await zip.file('word/document.xml')?.async('string')).toContain('<w:document/>')
  })
})

describe('the gate runs where Blob#stream is missing', () => {
  it('inflates without Blob.prototype.stream', async () => {
    // jsdom — the environment apps/docs' tests run in — ships a Blob with no
    // stream(). Inflating through it failed 389 docs tests in CI on #781 while
    // every Node-environment suite passed, so the gate must not depend on it.
    const proto = Blob.prototype as unknown as { stream?: () => unknown; [k: string]: unknown }
    const descriptor = Object.getOwnPropertyDescriptor(proto, 'stream')
    Object.defineProperty(proto, 'stream', { value: undefined, configurable: true, writable: true })
    try {
      const truth = Buffer.from('<?xml version="1.0"?><w:document/>')
      const payload = deflateRawSync(truth)
      const honest = writeZip([
        { meta: entry('word/document.xml', payload, truth.length), data: payload },
      ])
      await expect(assertZipInflatesWithinLimits(honest, LIMITS)).resolves.toBeUndefined()

      // and it still catches a lie through the same path
      await expect(
        assertZipInflatesWithinLimits(onePartArchive('word/document.xml', 64 * MB, 300), LIMITS),
      ).rejects.toThrow(/inflates past/)
    } finally {
      if (descriptor) Object.defineProperty(proto, 'stream', descriptor)
      else delete proto.stream
    }
  })
})

describe('zip-load stays bundleable into the renderer', () => {
  it('imports no Node builtins and nothing from zip-splice', () => {
    // The Docs renderer bundles zip-load through parseDocx. zip-splice is built
    // on node:fs and node:zlib and survives tree-shaking only while nothing
    // calls it — one live import there, or one node: builtin here, fails
    // `electron-vite build` on Vite's browser-external stubs. e2e caught exactly
    // that; this keeps it caught.
    const source = readFileSync(join(__dirname, '../src/zip-load.ts'), 'utf8')
    expect(source).not.toMatch(/from\s+'node:/)
    expect(source).not.toMatch(/from '\.\/zip-splice'/)
  })
})
