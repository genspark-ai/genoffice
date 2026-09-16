import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PackageArchive, PPTX_ZIP_LIMITS, assertZipWithinLimits } from '../src/index'

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

  it('rejects too many parts before inflating', async () => {
    const zip = new JSZip()
    for (let i = 0; i <= PPTX_ZIP_LIMITS.maxParts; i++) zip.file(`p/${i}.xml`, '<a/>')
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    await expect(PackageArchive.open(bytes)).rejects.toThrow(/parts exceeds/)
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
