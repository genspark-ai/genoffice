/** Undecodable media (e.g. TIFF) must never be dropped from the saved package. */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { openPptx, savePptx, createBlankPptx } from '../src/index'

async function deckWithTiffMedia(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await createBlankPptx())
  // Payload content doesn't matter for preservation — only the bytes surviving does
  const tiffBytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 1, 2, 3, 4, 5, 6, 7, 8])
  zip.file('ppt/media/image1.tiff', tiffBytes)
  zip.file('ppt/media/image2.tif', tiffBytes)
  return zip.generateAsync({ type: 'nodebuffer' })
}

describe('media preservation on save', () => {
  it('keeps .tif/.tiff parts byte-for-byte through open -> save', async () => {
    const opened = await openPptx(await deckWithTiffMedia())
    const out = await JSZip.loadAsync(await savePptx(opened))
    const tiff = out.file('ppt/media/image1.tiff')
    const tif = out.file('ppt/media/image2.tif')
    expect(tiff).not.toBeNull()
    expect(tif).not.toBeNull()
    const bytes = await tiff!.async('uint8array')
    expect(Array.from(bytes.slice(0, 4))).toEqual([0x49, 0x49, 0x2a, 0x00])
  })

  it('keeps them after an edit marks the slide dirty', async () => {
    const opened = await openPptx(await deckWithTiffMedia())
    const slide = opened.deck.slides[0]!
    slide.structureDirty = true
    const out = await JSZip.loadAsync(await savePptx(opened))
    expect(out.file('ppt/media/image1.tiff')).not.toBeNull()
    expect(out.file('ppt/media/image2.tif')).not.toBeNull()
  })
})
