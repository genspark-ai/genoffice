import { describe, expect, it } from 'vitest'
import {
  copyImageDisplaySize,
  MAX_COPY_IMAGE_DATA_URL,
  validCopyImageDataUrl,
} from '../src/main/copy-image-guard'

describe('copy-image guards', () => {
  it('accepts raster data URLs and lazy-media ids', () => {
    expect(validCopyImageDataUrl('data:image/png;base64,AAAA')).toBe(true)
    expect(validCopyImageDataUrl('data:image/jpeg;base64,AAAA')).toBe(true)
    expect(validCopyImageDataUrl('lazy-media-id-123')).toBe(true)
  })

  it('rejects oversized, mistyped, and non-string payloads', () => {
    expect(MAX_COPY_IMAGE_DATA_URL).toBe(64 * 1024 * 1024)
    expect(validCopyImageDataUrl('x'.repeat(MAX_COPY_IMAGE_DATA_URL + 1))).toBe(false)
    expect(validCopyImageDataUrl('data:image/png;base64,' + 'A'.repeat(40 * 1024 * 1024))).toBe(
      true,
    )
    expect(validCopyImageDataUrl('data:text/html,<script>alert(1)</script>')).toBe(false)
    expect(validCopyImageDataUrl('data:image/svg+xml;base64,AAAA')).toBe(false)
    expect(validCopyImageDataUrl(42)).toBe(false)
    expect(validCopyImageDataUrl(null)).toBe(false)
  })

  it('clamps display-size overrides to finite values', () => {
    expect(copyImageDisplaySize(JSON.stringify({ imageWidthPx: 200 }))).toEqual({ width: 200 })
    expect(copyImageDisplaySize(JSON.stringify({ imageWidthPx: Infinity }))).toEqual({})
    expect(copyImageDisplaySize(JSON.stringify({ imageHeightPx: 1e12 }))).toEqual({
      height: 100000,
    })
    expect(copyImageDisplaySize('not json')).toEqual({})
    expect(copyImageDisplaySize('x'.repeat(3000))).toEqual({})
  })
})
