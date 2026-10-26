import { describe, expect, it } from 'vitest'
import {
  collectPageImages,
  EMPTY_PAGE_IMAGE_PLAN,
  MAX_IMAGE_BASE64,
  MAX_IMAGES,
  MAX_TOTAL_BASE64,
  pageImageAnnotation,
  type PageImageBridge,
} from '../src/renderer/ai/page-images'
import type { PageImageBytes, PageImageRef } from '../src/shared/ipc'

const ref = (pageIndex: number, i: number): PageImageRef => ({
  pageIndex,
  rect: [100 + i * 50, 700, 140 + i * 50, 740],
  aboveText: false,
})

const png = (b64: string): PageImageBytes => ({ mime: 'image/png', base64: b64 })

/** Bridge mock: the IPC contract the collector drives (list → per-ref bytes) */
const bridge = (over: Partial<PageImageBridge> = {}): PageImageBridge => ({
  listImages: () => Promise.resolve([]),
  imageBytes: () => Promise.resolve(null),
  isDeleted: () => false,
  ...over,
})

describe('collectPageImages', () => {
  it('collects every image and numbers ordinals per page', async () => {
    const seen: PageImageRef[] = []
    const b = bridge({
      listImages: () => Promise.resolve([ref(0, 0), ref(0, 1), ref(2, 0)]),
      imageBytes: (r) => {
        seen.push(r)
        return Promise.resolve(png(`A${r.pageIndex}_${r.rect[0]}`))
      },
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toHaveLength(3)
    // the rect that reached the extractor matches the listed ref (coord → pixels chain)
    expect(seen.map((r) => r.rect)).toEqual([[100, 700, 140, 740], [150, 700, 190, 740], [100, 700, 140, 740]])
    expect(plan.ordinalsByPage.get(0)).toEqual([1, 2])
    expect(plan.ordinalsByPage.get(2)).toEqual([3])
    expect(plan.skipped).toBe(0)
  })

  it('passes extracted bytes through as AgentImage (base64 + mime)', async () => {
    const b = bridge({
      listImages: () => Promise.resolve([ref(1, 0)]),
      imageBytes: () => Promise.resolve({ mime: 'image/jpeg', base64: '/9j/4AAQ' }),
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toEqual([{ base64: '/9j/4AAQ', mime: 'image/jpeg' }])
  })

  it('drops undecodable images and counts them as skipped', async () => {
    const b = bridge({
      listImages: () => Promise.resolve([ref(0, 0), ref(0, 1)]),
      imageBytes: (r) => Promise.resolve(r.rect[0] === 100 ? png('A') : null),
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toHaveLength(1)
    expect(plan.skipped).toBe(1)
  })

  it('truncates at the 8-image cap and reports the remainder', async () => {
    const refs = Array.from({ length: 11 }, (_, i) => ref(0, i))
    const b = bridge({
      listImages: () => Promise.resolve(refs),
      imageBytes: (r) => Promise.resolve(png(`A${r.rect[0]}`)),
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toHaveLength(MAX_IMAGES)
    expect(plan.skipped).toBe(3)
  })

  it('skips a single image over the 5MB base64 cap', async () => {
    const b = bridge({
      listImages: () => Promise.resolve([ref(0, 0)]),
      imageBytes: () => Promise.resolve(png('a'.repeat(MAX_IMAGE_BASE64 + 1))),
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toHaveLength(0)
    expect(plan.skipped).toBe(1)
  })

  it('stops when the 20MB total budget would be exceeded', async () => {
    // 5 images of ~4.4MB: the 5th would push the total past 20MB
    const size = Math.floor(MAX_TOTAL_BASE64 / 4.5)
    const refs = Array.from({ length: 5 }, (_, i) => ref(0, i))
    const b = bridge({
      listImages: () => Promise.resolve(refs),
      imageBytes: (r) => Promise.resolve(png(`${'a'.repeat(size)}${r.rect[0]}`)),
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toHaveLength(4)
    expect(plan.skipped).toBe(1)
    expect(plan.images.every((img) => img.base64.length <= MAX_IMAGE_BASE64)).toBe(true)
  })

  it('deduplicates identical content and counts the duplicate as skipped', async () => {
    const b = bridge({
      listImages: () => Promise.resolve([ref(0, 0), ref(1, 0)]),
      imageBytes: () => Promise.resolve(png('same')),
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toEqual([{ base64: 'same', mime: 'image/png' }])
    expect(plan.ordinalsByPage.size).toBe(1)
    expect(plan.ordinalsByPage.get(0)).toEqual([1])
    expect(plan.skipped).toBe(1)
  })

  it('skips images on deleted pages before the cap counts them', async () => {
    const b = bridge({
      listImages: () => Promise.resolve([ref(0, 0), ref(1, 0)]),
      imageBytes: () => Promise.resolve(png('A')),
      isDeleted: (i) => i === 0,
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toHaveLength(1)
    expect(plan.ordinalsByPage.get(1)).toEqual([1])
    expect(plan.skipped).toBe(0)
  })

  it('rejects mimes outside the allowlist', async () => {
    const b = bridge({
      listImages: () => Promise.resolve([ref(0, 0)]),
      imageBytes: () => Promise.resolve({ mime: 'image/tiff', base64: 'A' } as PageImageBytes),
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toHaveLength(0)
    expect(plan.skipped).toBe(1)
  })

  it('degrades to the empty plan when listing throws', async () => {
    const plan = await collectPageImages(
      bridge({ listImages: () => Promise.reject(new Error('ipc down')) }),
    )
    expect(plan).toEqual(EMPTY_PAGE_IMAGE_PLAN)
  })

  it('degrades to the empty plan when one byte read throws', async () => {
    const b = bridge({
      listImages: () => Promise.resolve([ref(0, 0)]),
      imageBytes: () => Promise.reject(new Error('pdfium busy')),
    })
    const plan = await collectPageImages(b)
    expect(plan.images).toHaveLength(0)
    expect(plan.skipped).toBe(1)
  })
})

describe('pageImageAnnotation', () => {
  it('renders one line per page with the attachment ordinals', () => {
    const plan = {
      images: [{ base64: 'a', mime: 'image/png' }],
      ordinalsByPage: new Map([
        [0, [1]],
        [2, [2]],
      ]),
      skipped: 0,
    }
    expect(pageImageAnnotation(plan)).toBe(
      'page 1 has 1 picture(s) — attached as image #1\npage 3 has 1 picture(s) — attached as image #2',
    )
  })

  it('joins multiple ordinals of one page with commas', () => {
    const plan = {
      images: [
        { base64: 'a', mime: 'image/png' },
        { base64: 'b', mime: 'image/png' },
        { base64: 'c', mime: 'image/png' },
      ],
      ordinalsByPage: new Map([[4, [1, 2, 3]]]),
      skipped: 0,
    }
    expect(pageImageAnnotation(plan)).toBe(
      'page 5 has 3 picture(s) — attached as image #1, #2, #3',
    )
  })

  it('appends the overflow line when the budget skipped images', () => {
    const plan = {
      images: [{ base64: 'a', mime: 'image/png' }],
      ordinalsByPage: new Map([[0, [1]]]),
      skipped: 5,
    }
    expect(pageImageAnnotation(plan)).toBe(
      'page 1 has 1 picture(s) — attached as image #1\n(5 more pictures not attached)',
    )
  })

  it('returns an empty string for an empty plan (non-vision runs stay unchanged)', () => {
    expect(pageImageAnnotation(EMPTY_PAGE_IMAGE_PLAN)).toBe('')
  })
})
