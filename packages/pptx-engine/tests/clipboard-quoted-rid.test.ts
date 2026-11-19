/**
 * The clipboard scans the element's XML for the relationships it points at.
 * That scan only matched double-quoted attributes, so an element written with
 * single quotes copied with `rels: []` and no bundled parts: the paste kept
 * the source slide's rId, which does not exist on the target slide, and the
 * picture came out broken.
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import {
  addPicture,
  copyElementData,
  createBlankPptx,
  openPptx,
  savePptx,
  type PictureElement,
} from '../src/index'

const OFF = { x: 914400, y: 914400, cx: 1828800, cy: 914400 }

// 1x1 red PNG
const PNG_1PX = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

describe('copying an element whose rels are single-quoted', () => {
  it('still finds the picture relationship and bundles its bytes', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const pic = addPicture(opened, slide, { bytes: PNG_1PX, ext: 'png', offset: { ...OFF } })!

    // Rewrite the saved slide with single-quoted attributes, the way some
    // producers serialize (and as the other relationship scanners already allow).
    const zip = await JSZip.loadAsync(await savePptx(opened))
    const original = await zip.file(slide.path)!.async('string')
    const quoted = original.replace(/r:embed="(rId\d+)"/g, "r:embed='$1'")
    expect(quoted).not.toBe(original)
    expect(quoted).toContain("r:embed='")
    zip.file(slide.path, quoted)

    const reopened = await openPptx(await zip.generateAsync({ type: 'uint8array' }))
    const reopenedSlide = reopened.deck.slides[0]!
    const reopenedPic = reopenedSlide.elements.find(
      (element) => element.type === 'picture',
    ) as PictureElement

    const clip = copyElementData(reopened, reopenedSlide, reopenedPic)
    expect(clip.xml).toContain("r:embed='")
    expect(clip.rels).toHaveLength(1)
    expect(clip.rels[0]!.rid).toBe(/r:embed='(rId\d+)'/.exec(clip.xml)?.[1] ?? '')
    expect(clip.parts?.[pic.mediaRef!]).toBeDefined()
  })
})
