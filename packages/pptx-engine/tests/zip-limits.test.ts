import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PackageArchive, PPTX_ZIP_LIMITS, assertZipWithinLimits } from '../src/index'
import { resolveTarget } from '../src/zip'

const here = dirname(fileURLToPath(import.meta.url))

function fakeZip(parts: Array<{ name: string; size: number }>): JSZip {
  const files = Object.fromEntries(
    parts.map((p) => [p.name, { name: p.name, dir: false, _data: { uncompressedSize: p.size } }]),
  )
  return { files } as unknown as JSZip
}

describe('pptx zip limits', () => {
  it('opens a normal deck', async () => {
    const bytes = readFileSync(join(here, 'fixtures/01_standard_business.pptx'))
    const pkg = await PackageArchive.open(new Uint8Array(bytes))
    expect(pkg.has('ppt/presentation.xml')).toBe(true)
  })

  it('normalizes encoded and backslash relationship targets', () => {
    expect(resolveTarget('ppt/presentation.xml', 'slides/slide%201.xml')).toBe(
      'ppt/slides/slide 1.xml',
    )
    expect(resolveTarget('ppt/slides/slide1.xml', '..\\media\\image%201.png')).toBe(
      'ppt/media/image 1.png',
    )
    expect(resolveTarget('ppt/presentation.xml', '/ppt/slides/slide1.xml')).toBe(
      'ppt/slides/slide1.xml',
    )
    expect(resolveTarget('ppt/slides/slide1.xml', '\\ppt\\slides\\slide2.xml')).toBe(
      'ppt/slides/slide2.xml',
    )
  })

  it('rejects too many parts before inflating', async () => {
    const zip = new JSZip()
    for (let i = 0; i <= PPTX_ZIP_LIMITS.maxParts; i++) zip.file(`p/${i}.xml`, '<a/>')
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    await expect(PackageArchive.open(bytes)).rejects.toThrow(/parts exceeds/)
  })

  it('rejects a lying declaration by inflating past the claim (GH #759)', async () => {
    // 1.5 MB of real payload behind a 300-byte declaration: the advisory pass
    // reads the lie as truth, so only the metered gate can catch it.
    const zip = new JSZip()
    zip.file('ppt/presentation.xml', '<p/>')
    zip.file('ppt/slides/slide1.xml', Buffer.alloc(1536 * 1024))
    const bytes = (await zip.generateAsync({ type: 'uint8array' })) as Uint8Array
    const buf = Buffer.from(bytes)
    const centralName = buf.lastIndexOf(Buffer.from('ppt/slides/slide1.xml'))
    buf.writeUInt32LE(300, centralName - 46 + 24)
    await expect(PackageArchive.open(new Uint8Array(buf))).rejects.toThrow(
      /ppt\/slides\/slide1\.xml declares 300 uncompressed bytes but inflates past that/,
    )
  })

  it('rejects oversized parts and totals by declared size', () => {
    expect(() =>
      assertZipWithinLimits(fakeZip([{ name: 'a.bin', size: PPTX_ZIP_LIMITS.maxPartBytes + 1 }])),
    ).toThrow(/declares/)
    const half = PPTX_ZIP_LIMITS.maxPartBytes
    expect(() =>
      assertZipWithinLimits(
        fakeZip([
          { name: 'a.bin', size: half },
          { name: 'b.bin', size: half },
          { name: 'c.bin', size: half },
          { name: 'd.bin', size: half },
        ]),
      ),
    ).toThrow(/total uncompressed/)
    expect(() => assertZipWithinLimits(fakeZip([{ name: 'a.xml', size: 10 }]))).not.toThrow()
  })
})
