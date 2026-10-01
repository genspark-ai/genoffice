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

  it('parses every parameter run a real data URL can carry', () => {
    // the fix narrows the parameter body to [^;,]*, so the accepted set must
    // not shrink: mime, the ;base64 flag and extra parameters all still split
    const b64 = decodeDataUrl('data:image/png;charset=x;base64,QUJD')
    expect(b64?.mime).toBe('image/png')
    expect(b64?.bytes.toString()).toBe('ABC')
    const charset = decodeDataUrl('data:image/svg+xml;charset=utf-8,%3Csvg%2F%3E')
    expect(charset?.mime).toBe('image/svg+xml')
    expect(charset?.bytes.toString()).toBe('<svg/>')
    expect(decodeDataUrl('data:,hello')?.bytes.toString()).toBe('hello')
  })

  it('is not exponential on a comma-less data URL', () => {
    // A `;[^,]*` parameter run can cover the remaining text in one iteration
    // or in many, so with no comma the engine tried every split: ~4x per 4
    // extra characters, i.e. minutes from ~60 characters and unbounded past
    // ~100. This ran on any data: image URL in a document, in the main process.
    const url = 'data:image/png' + ';a'.repeat(40_000)
    const started = performance.now()
    expect(decodeDataUrl(url)).toBeNull()
    const elapsed = performance.now() - started
    expect(elapsed).toBeLessThan(1_000)
  })

  it('requires an exact ;base64 parameter, not a substring match', () => {
    const lookalike = decodeDataUrl('data:image/png;base64x=1,%3Csvg%3E')
    expect(lookalike?.bytes.toString()).toBe('<svg>')
    const real = decodeDataUrl('data:image/png;charset=utf-8;base64,aGVsbG8=')
    expect(real?.bytes.toString()).toBe('hello')
  })
})
