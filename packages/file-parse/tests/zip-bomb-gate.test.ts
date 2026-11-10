import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { pptxToText } from '../src/pptx'
import { xlsxToText } from '../src/xlsx'

/** a structurally valid zip whose part declares 600MB uncompressed: the gate
 *  must reject it before anything is inflated (GH #759) */
async function bytesWithHugeDeclaration(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>',
  )
  const bytes = (await zip.generateAsync({ type: 'nodebuffer' })) as Buffer
  // rewrite the *file* entry's declared uncompressed size — the zip also holds
  // an auto-created 'xl/' dir entry whose central record comes first
  const centralName = bytes.lastIndexOf(Buffer.from('xl/workbook.xml'))
  bytes.writeUInt32LE(600_000_000, centralName - 46 + 24)
  return new Uint8Array(bytes)
}

/** same shape, but the part lies: 1.5 MB of real payload declared as 300 bytes.
 *  Only the metered gate can catch this — the advisory pass reads the lie. */
async function bytesWithLyingDeclaration(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>',
  )
  zip.file('xl/worksheets/sheet1.xml', Buffer.alloc(1536 * 1024))
  const bytes = (await zip.generateAsync({ type: 'nodebuffer' })) as Buffer
  const centralName = bytes.lastIndexOf(Buffer.from('xl/worksheets/sheet1.xml'))
  bytes.writeUInt32LE(300, centralName - 46 + 24)
  return new Uint8Array(bytes)
}

describe('xlsx/pptx zip bomb gate', () => {
  it('rejects an xlsx whose parts declare more than the per-part limit', async () => {
    await expect(xlsxToText(await bytesWithHugeDeclaration())).rejects.toThrow(
      /declares 600000000 uncompressed bytes/,
    )
  })

  it('rejects a pptx the same way', async () => {
    await expect(pptxToText(await bytesWithHugeDeclaration())).rejects.toThrow(
      /declares 600000000 uncompressed bytes/,
    )
  })

  it('rejects a lying declaration by inflating past the claim, not by reading it', async () => {
    await expect(xlsxToText(await bytesWithLyingDeclaration())).rejects.toThrow(
      /inflates past that/,
    )
  })

  it('rejects a lying pptx declaration the same way', async () => {
    await expect(pptxToText(await bytesWithLyingDeclaration())).rejects.toThrow(
      /inflates past that/,
    )
  })
})
