// Lazy media URLs for Slides: the shape the render tree carries, how main parses it
// back, and which parts qualify. (The protocol handler itself needs electron and is
// covered by the app run, like the docs app's equivalent.)
import { describe, expect, it } from 'vitest'
import {
  LAZY_MEDIA_MIN_PART_BYTES,
  parseSlidesMediaUrl,
  shouldServeLazily,
  slidesMediaBase,
  slidesMediaUrl,
} from '../src/shared/media-url'

describe('slides media URLs', () => {
  it('round-trips a part path through the URL', () => {
    const base = slidesMediaBase('deck3')
    const url = slidesMediaUrl(base, 'ppt/media/image12.png')
    expect(url).toBe('genoffice-slides-media://deck3/ppt%2Fmedia%2Fimage12.png')
    expect(parseSlidesMediaUrl(url)).toEqual({ key: 'deck3', mediaRef: 'ppt/media/image12.png' })
  })

  it('keeps spaces and non-ascii part names intact', () => {
    const url = slidesMediaUrl(slidesMediaBase('d1'), 'ppt/media/图 1.png')
    expect(parseSlidesMediaUrl(url)?.mediaRef).toBe('ppt/media/图 1.png')
  })

  it('rejects anything that is not one of our URLs', () => {
    expect(parseSlidesMediaUrl('genoffice-docx-media://deck3/word/media/image1.png')).toBeNull()
    expect(parseSlidesMediaUrl('data:image/png;base64,AAAA')).toBeNull()
    expect(parseSlidesMediaUrl('genoffice-slides-media://deck3/')).toBeNull()
    expect(parseSlidesMediaUrl('genoffice-slides-media://deck3')).toBeNull()
  })
})

describe('shouldServeLazily', () => {
  it('serves rasters at or above the threshold over the protocol', () => {
    expect(shouldServeLazily('image/png', LAZY_MEDIA_MIN_PART_BYTES)).toBe(true)
    expect(shouldServeLazily('image/jpeg', LAZY_MEDIA_MIN_PART_BYTES * 4)).toBe(true)
    expect(shouldServeLazily('image/webp', LAZY_MEDIA_MIN_PART_BYTES + 1)).toBe(true)
  })

  it('inlines parts below the threshold, where base64 costs nothing worth a round trip', () => {
    expect(shouldServeLazily('image/png', LAZY_MEDIA_MIN_PART_BYTES - 1)).toBe(false)
    expect(shouldServeLazily('image/jpeg', 1024)).toBe(false)
  })

  it('inlines SVG and metafiles whatever their size', () => {
    // SVG is retinted against the slide's theme before display, so it keeps the
    // per-slide inline path; EMF/WMF are rasterized on the renderer side.
    expect(shouldServeLazily('image/svg+xml', LAZY_MEDIA_MIN_PART_BYTES * 10)).toBe(false)
    expect(shouldServeLazily('image/x-emf', LAZY_MEDIA_MIN_PART_BYTES * 10)).toBe(false)
    expect(shouldServeLazily('image/wmf', LAZY_MEDIA_MIN_PART_BYTES * 10)).toBe(false)
  })
})
