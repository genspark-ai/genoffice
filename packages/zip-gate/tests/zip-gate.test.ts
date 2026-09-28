import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  assertDeclaredSizesWithinLimits,
  assertZipInflatesWithinLimits,
  DEFAULT_ZIP_LIMITS,
  type DeclaredPart,
} from '../src/index'

const KB = 1024
const MB = 1024 * KB

/** Small budgets so the tests stay fast; the shape of the check is the point. */
const LIMITS = { maxParts: 100, maxPartBytes: 2 * MB, maxTotalBytes: 8 * MB } as const

const NAME = 'word/document.xml'

/** A valid one-part archive whose real payload is `realBytes`, with the
 *  central directory's declared uncompressed size rewritten to `declares`. */
async function onePart(
  realBytes: number,
  declares: number,
  opts?: { method?: 'store'; name?: string },
): Promise<Buffer> {
  const zip = new JSZip()
  const name = opts?.name ?? NAME
  const content =
    opts?.method === 'store'
      ? zip.file(name, Buffer.alloc(realBytes), { compression: 'STORE' })
      : zip.file(name, Buffer.alloc(realBytes))
  void content
  const bytes = (await zip.generateAsync({ type: 'nodebuffer' })) as Buffer
  // rewrite the *file* entry's declared size (JSZip also emits an auto-created
  // directory entry whose central record comes first)
  const centralName = bytes.lastIndexOf(Buffer.from(name))
  bytes.writeUInt32LE(declares, centralName - 46 + 24)
  return bytes
}

describe('assertZipInflatesWithinLimits', () => {
  it('refuses a part that declares less than it inflates to', async () => {
    // The #759 case: megabytes behind a 300-byte declaration. Every declared
    // number is inside the limits, so only inflating can find out.
    const archive = await onePart(1024 * KB, 300)
    await expect(assertZipInflatesWithinLimits(archive, LIMITS)).rejects.toThrow(
      new RegExp(
        `part ${NAME.replace(/\//g, '\\/')} declares 300 uncompressed bytes but inflates past that`,
      ),
    )
  })

  it('catches the lie at the declared size, not at the payload', async () => {
    const archive = await onePart(1024 * KB, 300)
    const before = process.memoryUsage().rss
    await expect(assertZipInflatesWithinLimits(archive, LIMITS)).rejects.toThrow(/inflates past/)
    // The budget is one byte past the claim, so nothing near the payload size
    // is ever allocated.
    expect(process.memoryUsage().rss - before).toBeLessThan(16 * MB)
  })

  it('accepts an honest part, and a stored part that tells the truth', async () => {
    await expect(
      assertZipInflatesWithinLimits(await onePart(300, 300), LIMITS),
    ).resolves.toBeUndefined()
    await expect(
      assertZipInflatesWithinLimits(
        await onePart(64 * KB, 64 * KB, { method: 'store', name: 'word/media/image1.png' }),
        LIMITS,
      ),
    ).resolves.toBeUndefined()
  })

  it('refuses a stored part that holds more bytes than it declares', async () => {
    await expect(
      assertZipInflatesWithinLimits(await onePart(64 * KB, 300, { method: 'store' }), LIMITS),
    ).rejects.toThrow(/inflates past that/)
  })

  it('rejects oversized declarations before inflating anything', async () => {
    const big = await onePart(300, LIMITS.maxPartBytes + 1)
    await expect(assertZipInflatesWithinLimits(big, LIMITS)).rejects.toThrow(/limit 2097152/)
  })

  it('rejects excess totals from verified lengths, not claims', async () => {
    // Two stored 3 MB parts honestly declaring 3 MB each: the verified total
    // busts the 4 MB budget even though no single part is over maxPartBytes.
    const zip = new JSZip()
    zip.file('a.bin', Buffer.alloc(3 * MB), { compression: 'STORE' })
    zip.file('b.bin', Buffer.alloc(3 * MB), { compression: 'STORE' })
    const bytes = (await zip.generateAsync({ type: 'nodebuffer' })) as Buffer
    await expect(
      assertZipInflatesWithinLimits(new Uint8Array(bytes), {
        ...LIMITS,
        maxPartBytes: 4 * MB,
        maxTotalBytes: 5 * MB,
      }),
    ).rejects.toThrow(/total uncompressed size/)
  })

  it('bounds every entry against the archive and refuses unsupported forms', async () => {
    // A declared csize past the end of the file must never reach a reader
    // that allocates from declarations.
    const good = await onePart(300, 300)
    const CENTRAL_SIG = Buffer.from([0x50, 0x4b, 0x01, 0x02])
    const centralOfPart = good.lastIndexOf(Buffer.from(NAME)) - 46
    const oversized = Buffer.from(good)
    oversized.writeUInt32LE(0xffffff00, centralOfPart + 20)
    await expect(assertZipInflatesWithinLimits(oversized, LIMITS)).rejects.toThrow(
      /outside the .*-byte archive/,
    )

    // Encrypted flag in the central directory.
    const encrypted = Buffer.from(good)
    encrypted.writeUInt16LE(0x1, centralOfPart + 8)
    await expect(assertZipInflatesWithinLimits(encrypted, LIMITS)).rejects.toThrow(/encrypted/)

    // Garbage that is not a zip at all.
    await expect(assertZipInflatesWithinLimits(Buffer.from('not a zip'), LIMITS)).rejects.toThrow(
      /end of central directory/,
    )
  })

  it('keeps the default policy shared with the engines', () => {
    expect(DEFAULT_ZIP_LIMITS.maxParts).toBe(10000)
    expect(DEFAULT_ZIP_LIMITS.maxPartBytes).toBe(512 * 1024 * 1024)
    expect(DEFAULT_ZIP_LIMITS.maxTotalBytes).toBe(1.5 * 1024 * 1024 * 1024)
  })
})

describe('assertDeclaredSizesWithinLimits', () => {
  it('rejects too many parts by declared count alone', () => {
    const parts: DeclaredPart[] = Array.from({ length: LIMITS.maxParts + 1 }, (_, i) => ({
      name: `p/${i}`,
      usize: 1,
    }))
    expect(() => assertDeclaredSizesWithinLimits(parts, LIMITS)).toThrow(/parts exceeds/)
    expect(() =>
      assertDeclaredSizesWithinLimits([{ name: 'a.xml', usize: 10 }], LIMITS),
    ).not.toThrow()
  })
})
