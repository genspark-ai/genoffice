import { describe, expect, it } from 'vitest'
import {
  addChart,
  addElement,
  addMedia,
  addModel3d,
  addPicture,
  createBlankPptx,
  deleteElement,
  deleteSlide,
  duplicateSlide,
  groupElements,
  openPptx,
  replacePictureBytes,
  resetSlideBackground,
  savePptx,
  setElementFill,
  setElementImageFill,
  setSlideBackground,
  setSlideBackgroundImage,
} from '../src/index'
import type { GroupElement, PictureElement } from '../src/types'
import { relsPathFor } from '../src/zip'

const PNG_1PX = new Uint8Array(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+XwV2AAAAAElFTkSuQmCC',
    'base64',
  ),
)
const GIF_1PX = new Uint8Array(
  Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64'),
)
const OFF = { x: 914400, y: 914400, cx: 3657600, cy: 2057400 }

const mediaParts = (opened: Awaited<ReturnType<typeof openPptx>>) =>
  [...opened.archive.entries.keys()].filter((path) => path.startsWith('ppt/media/'))

describe('deleted resource cleanup', () => {
  it('removes an unshared picture relationship and media part', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const picture = addPicture(opened, slide, { bytes: PNG_1PX, ext: 'png', offset: OFF })!
    const mediaPath = picture.mediaRef

    expect(opened.archive.entries.has(mediaPath)).toBe(true)
    expect(deleteElement(opened, slide, picture.id)).toBe(true)
    expect(opened.archive.entries.has(mediaPath)).toBe(false)
    expect(opened.archive.readText(relsPathFor(slide.path))).not.toContain('/image"')

    const reopened = await openPptx(await savePptx(opened))
    expect(mediaParts(reopened)).toEqual([])
  })

  it('keeps a media part until the final slide reference is deleted', async () => {
    const opened = await openPptx(await createBlankPptx())
    const first = opened.deck.slides[0]!
    const picture = addPicture(opened, first, { bytes: PNG_1PX, ext: 'png', offset: OFF })!
    const mediaPath = picture.mediaRef
    const second = duplicateSlide(opened, 0)!
    const duplicatePicture = second.elements.find((element) => element.type === 'picture')!

    expect(deleteElement(opened, first, picture.id)).toBe(true)
    expect(opened.archive.entries.has(mediaPath)).toBe(true)

    expect(deleteElement(opened, second, duplicatePicture.id)).toBe(true)
    expect(opened.archive.entries.has(mediaPath)).toBe(false)
  })

  it('removes both embedded video relationships, bytes, and poster on element deletion', async () => {
    const opened = await openPptx(await createBlankPptx())
    const added = addMedia(opened, 0, {
      kind: 'video',
      bytes: new Uint8Array([0, 0, 0, 24, 0x66, 0x74, 0x79, 0x70]),
      ext: 'mp4',
      offset: OFF,
    })!

    expect(mediaParts(opened).length).toBe(2)
    expect(deleteElement(opened, added.slide, added.elementId)).toBe(true)
    expect(mediaParts(opened)).toEqual([])
    const rels = opened.archive.readText(relsPathFor(added.slide.path))!
    expect(rels).not.toContain('/video"')
    expect(rels).not.toContain('/media"')
    expect(rels).not.toContain('/image"')
  })

  it('removes an unshared chart part and its content-type override', async () => {
    const opened = await openPptx(await createBlankPptx())
    const added = addChart(opened, 0, {
      kind: 'bar',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'Sales', values: [10, 20] }],
      offset: OFF,
    })!
    const chartPath = [...opened.archive.entries.keys()].find((path) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(path),
    )!

    expect(deleteElement(opened, added.slide, added.elementId)).toBe(true)
    expect(opened.archive.entries.has(chartPath)).toBe(false)
    expect(opened.archive.readText('[Content_Types].xml')).not.toContain(`PartName="/${chartPath}"`)
  })

  it('collects a shared 3D model after deleting its final live placeholder', async () => {
    const opened = await openPptx(await createBlankPptx())
    const added = addModel3d(opened, 0, {
      bytes: new Uint8Array([0x67, 0x6c, 0x54, 0x46]),
      ext: 'glb',
      offset: OFF,
    })!
    const modelPath = mediaParts(opened).find((path) => path.endsWith('.glb'))!
    const second = duplicateSlide(opened, 0)!
    const duplicateModel = second.elements.find((element) =>
      element.descr?.startsWith('aislides-3d:'),
    )!

    expect(deleteElement(opened, added.slide, added.elementId)).toBe(true)
    expect(opened.archive.entries.has(modelPath)).toBe(true)

    expect(deleteElement(opened, second, duplicateModel.id)).toBe(true)
    expect(opened.archive.entries.has(modelPath)).toBe(false)
  })

  it('collects resources owned by a deleted slide', async () => {
    const opened = await openPptx(await createBlankPptx())
    duplicateSlide(opened, 0)
    addMedia(opened, 0, {
      kind: 'video',
      bytes: new Uint8Array([0, 0, 0, 24, 0x66, 0x74, 0x79, 0x70]),
      ext: 'mp4',
      offset: OFF,
    })

    expect(mediaParts(opened).length).toBe(2)
    expect(deleteSlide(opened, 0)).toBe(true)
    expect(mediaParts(opened)).toEqual([])
    expect(opened.deck.slides).toHaveLength(1)
  })
})

describe('superseded picture resource cleanup', () => {
  it('removes the old picture relationship and media after replacement', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const picture = addPicture(opened, slide, { bytes: PNG_1PX, ext: 'png', offset: OFF })!
    const oldMedia = picture.mediaRef
    const oldRid = /r:embed="([^"]+)"/.exec(picture.anchor.originalXml)![1]!

    expect(replacePictureBytes(opened, slide, picture.id, GIF_1PX, 'gif')).toBe(true)
    expect(opened.archive.entries.has(oldMedia)).toBe(false)
    expect(opened.archive.readText(relsPathFor(slide.path))).not.toContain(`Id="${oldRid}"`)
    expect(opened.archive.entries.has((picture as PictureElement).mediaRef)).toBe(true)
  })

  it('keeps replaced picture media while another slide still uses it', async () => {
    const opened = await openPptx(await createBlankPptx())
    const first = opened.deck.slides[0]!
    const picture = addPicture(opened, first, { bytes: PNG_1PX, ext: 'png', offset: OFF })!
    const oldMedia = picture.mediaRef
    const second = duplicateSlide(opened, 0)!
    const duplicate = second.elements.find((element) => element.type === 'picture')!

    expect(replacePictureBytes(opened, first, picture.id, GIF_1PX, 'gif')).toBe(true)
    expect(opened.archive.entries.has(oldMedia)).toBe(true)

    expect(replacePictureBytes(opened, second, duplicate.id, GIF_1PX, 'gif')).toBe(true)
    expect(opened.archive.entries.has(oldMedia)).toBe(false)
  })
})

describe('superseded slide background cleanup', () => {
  it('collects an old image background when replaced by a new image', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const oldMedia = setSlideBackgroundImage(opened, slide, { bytes: PNG_1PX, ext: 'png' })!
    const oldRid = /r:embed="([^"]+)"/.exec(slide.bodyPrefix)![1]!

    const newMedia = setSlideBackgroundImage(opened, slide, { bytes: GIF_1PX, ext: 'gif' })!
    expect(newMedia).not.toBe(oldMedia)
    expect(opened.archive.entries.has(oldMedia)).toBe(false)
    expect(opened.archive.entries.has(newMedia)).toBe(true)
    expect(opened.archive.readText(relsPathFor(slide.path))).not.toContain(`Id="${oldRid}"`)
  })

  it('collects image backgrounds replaced by solid, gradient, or reset', async () => {
    const transitions = [
      (opened: Awaited<ReturnType<typeof openPptx>>) =>
        setSlideBackground(opened, opened.deck.slides[0]!, '#112233'),
      (opened: Awaited<ReturnType<typeof openPptx>>) =>
        setSlideBackground(opened, opened.deck.slides[0]!, {
          stops: [
            { pos: 0, color: '#112233' },
            { pos: 1, color: '#445566' },
          ],
          angle: 5400000,
        }),
      (opened: Awaited<ReturnType<typeof openPptx>>) =>
        resetSlideBackground(opened, opened.deck.slides[0]!),
    ]

    for (const transition of transitions) {
      const opened = await openPptx(await createBlankPptx())
      const slide = opened.deck.slides[0]!
      const oldMedia = setSlideBackgroundImage(opened, slide, { bytes: PNG_1PX, ext: 'png' })!
      transition(opened)
      expect(opened.archive.entries.has(oldMedia)).toBe(false)
      expect(opened.archive.readText(relsPathFor(slide.path))).not.toContain('/image"')
    }
  })

  it('preserves shared background media until the final slide drops it', async () => {
    const opened = await openPptx(await createBlankPptx())
    const first = opened.deck.slides[0]!
    const media = setSlideBackgroundImage(opened, first, { bytes: PNG_1PX, ext: 'png' })!
    const second = duplicateSlide(opened, 0)!

    setSlideBackground(opened, first, '#112233')
    expect(opened.archive.entries.has(media)).toBe(true)

    resetSlideBackground(opened, second)
    expect(opened.archive.entries.has(media)).toBe(false)
  })
})

describe('superseded shape picture-fill cleanup', () => {
  function addShape(opened: Awaited<ReturnType<typeof openPptx>>, x = 0) {
    return addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { ...OFF, x },
      fillColor: '#112233',
    })
  }

  it('collects an old picture fill when replaced by a new image', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const shape = addShape(opened)
    const oldMedia = setElementImageFill(opened, slide, shape.id, {
      bytes: PNG_1PX,
      ext: 'png',
    })!

    const newMedia = setElementImageFill(opened, slide, shape.id, {
      bytes: GIF_1PX,
      ext: 'gif',
    })!
    expect(newMedia).not.toBe(oldMedia)
    expect(opened.archive.entries.has(oldMedia)).toBe(false)
    expect(opened.archive.entries.has(newMedia)).toBe(true)
  })

  it('collects picture fills replaced by solid, gradient, or none', async () => {
    const fills = [
      '#112233',
      {
        stops: [
          { pos: 0, color: '#112233' },
          { pos: 1, color: '#445566' },
        ],
        angle: 5400000,
      },
      'none',
    ]

    for (const fill of fills) {
      const opened = await openPptx(await createBlankPptx())
      const slide = opened.deck.slides[0]!
      const shape = addShape(opened)
      const oldMedia = setElementImageFill(opened, slide, shape.id, {
        bytes: PNG_1PX,
        ext: 'png',
      })!
      expect(setElementFill(opened, slide, shape.id, fill)).toBe(true)
      expect(opened.archive.entries.has(oldMedia)).toBe(false)
      expect(opened.archive.readText(relsPathFor(slide.path))).not.toContain('/image"')
    }
  })

  it('preserves shared fill media until the final shape drops it', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const first = addShape(opened)
    const second = addShape(opened, 4000000)
    const media = setElementImageFill(opened, slide, first.id, {
      bytes: PNG_1PX,
      ext: 'png',
    })!
    setElementImageFill(opened, slide, second.id, { mediaPath: media })

    expect(setElementFill(opened, slide, first.id, '#112233')).toBe(true)
    expect(opened.archive.entries.has(media)).toBe(true)

    expect(setElementFill(opened, slide, second.id, 'none')).toBe(true)
    expect(opened.archive.entries.has(media)).toBe(false)
  })

  it('cleans a superseded picture fill on a group child', async () => {
    const opened = await openPptx(await createBlankPptx())
    const first = addShape(opened)
    const second = addShape(opened, 4000000)
    const grouped = groupElements(opened, 0, [first.id, second.id])!
    const slide = grouped.slide
    const group = slide.elements.find((element) => element.id === grouped.groupId) as GroupElement
    const child = group.children.find((element) => element.type === 'shape')!
    const media = setElementImageFill(
      opened,
      slide,
      child.id,
      { bytes: PNG_1PX, ext: 'png' },
      { groupId: group.id },
    )!

    expect(
      setElementFill(opened, slide, child.id, '#112233', {
        groupId: group.id,
      }),
    ).toBe(true)
    expect(opened.archive.entries.has(media)).toBe(false)
  })
})
