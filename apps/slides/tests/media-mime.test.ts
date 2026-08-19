import { describe, it, expect } from 'vitest'
import { sniffImageMime, displayMime } from '../src/main/media-mime'

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0])
const EMF = new Uint8Array(48)
EMF.set([0x01, 0x00, 0x00, 0x00], 0)
EMF.set([0x20, 0x45, 0x4d, 0x46], 40)
const WMF = new Uint8Array([0xd7, 0xcd, 0xc6, 0x9a])
const TIFF = new Uint8Array([0x49, 0x49, 0x2a, 0x00])

describe('sniffImageMime', () => {
  it('recognizes common containers', () => {
    expect(sniffImageMime(PNG)).toBe('image/png')
    expect(sniffImageMime(JPEG)).toBe('image/jpeg')
    expect(sniffImageMime(EMF)).toBe('image/x-emf')
    expect(sniffImageMime(WMF)).toBe('image/x-wmf')
    expect(sniffImageMime(TIFF)).toBe('image/tiff')
    expect(sniffImageMime(new Uint8Array([0, 1, 2, 3]))).toBeNull()
  })
})

describe('displayMime', () => {
  it('magic bytes win over a lying extension (PNG stored as .emf)', () => {
    expect(displayMime('ppt/media/image1.emf', PNG)).toBe('image/png')
    expect(displayMime('ppt/media/image1.png', EMF)).toBe('image/x-emf')
  })

  it('falls back to the extension, then PNG', () => {
    const unknown = new Uint8Array([0x00, 0x11, 0x22, 0x33])
    expect(displayMime('ppt/media/a.wmf', unknown)).toBe('image/x-wmf')
    expect(displayMime('ppt/media/a.xyz', unknown)).toBe('image/png')
  })
})
