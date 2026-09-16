import { describe, expect, it } from 'vitest'

import { decodeDataUrl, isSavableImageUrl, suggestImageFileName } from '../src/index'

describe('isSavableImageUrl', () => {
  it('accepts data, http(s) and app asset schemes', () => {
    expect(isSavableImageUrl('data:image/png;base64,AAAA')).toBe(true)
    expect(isSavableImageUrl('https://example.com/a.png')).toBe(true)
    expect(isSavableImageUrl('md-asset:///Users/me/doc/assets/image.png')).toBe(true)
  })
  it('refuses local files, renderer blobs and non-URLs', () => {
    expect(isSavableImageUrl('file:///etc/passwd')).toBe(false)
    expect(isSavableImageUrl('blob:genoffice-app://docs/1234')).toBe(false)
    expect(isSavableImageUrl('assets/image.png')).toBe(false)
  })
})

describe('suggestImageFileName', () => {
  it('keeps the URL file name when it has an image extension', () => {
    expect(suggestImageFileName('md-asset:///d/assets/photo%201.JPG', 'image/jpeg')).toBe(
      'photo 1.JPG',
    )
    expect(suggestImageFileName('https://x.test/img/chart.webp?v=2', null)).toBe('chart.webp')
  })
  it('falls back to the MIME type, then png', () => {
    expect(suggestImageFileName('data:image/jpeg;base64,AAAA', 'image/jpeg')).toBe('image.jpg')
    expect(suggestImageFileName('https://x.test/render', 'image/svg+xml; charset=utf-8')).toBe(
      'image.svg',
    )
    expect(suggestImageFileName('https://x.test/render', 'application/octet-stream')).toBe(
      'image.png',
    )
  })
})

describe('decodeDataUrl', () => {
  it('decodes base64 and percent-encoded payloads', () => {
    const b64 = decodeDataUrl('data:image/png;base64,aGVsbG8=')
    expect(b64?.mime).toBe('image/png')
    expect(b64?.bytes.toString()).toBe('hello')
    const plain = decodeDataUrl('data:image/svg+xml,%3Csvg%2F%3E')
    expect(plain?.bytes.toString()).toBe('<svg/>')
  })
  it('rejects malformed input', () => {
    expect(decodeDataUrl('data:nope')).toBeNull()
  })
})
