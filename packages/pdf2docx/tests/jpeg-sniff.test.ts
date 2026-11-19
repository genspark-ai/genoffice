/** P29 A: only 1-/3-component JPEGs may embed raw — Word paints CMYK/YCCK black. */
import { describe, expect, it } from 'vitest'
import { jpegSofComponents } from '../src/extract'

function jpeg(components: number, sofMarker = 0xc0): Uint8Array {
  const app0 = [0xff, 0xe0, 0x00, 0x04, 0x00, 0x00]
  const sof = [0xff, sofMarker, 0x00, 0x08, 8, 0x00, 0x10, 0x00, 0x10, components]
  return Uint8Array.from([0xff, 0xd8, ...app0, ...sof])
}

describe('jpegSofComponents', () => {
  it('reads the component count from SOF0', () => {
    expect(jpegSofComponents(jpeg(3))).toBe(3)
  })
  it('reads progressive SOF2 too', () => {
    expect(jpegSofComponents(jpeg(4, 0xc2))).toBe(4)
  })
  it('returns 0 for non-JPEG bytes', () => {
    expect(jpegSofComponents(Uint8Array.from([0x89, 0x50, 0x4e, 0x47]))).toBe(0)
  })
})
