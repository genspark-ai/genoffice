import { describe, expect, it } from 'vitest'
import { isMetafileMime, metafileToDataUrl } from '../src/metafile'

/** EMR_HEADER (88 bytes) + EMR_EOF (20 bytes), one logical 100x100 frame */
function minimalEmf(): Uint8Array {
  const bytes = new Uint8Array(108)
  const dv = new DataView(bytes.buffer)
  dv.setUint32(0, 1, true) // EMR_HEADER
  dv.setUint32(4, 88, true)
  dv.setInt32(16, 100, true) // rclBounds right
  dv.setInt32(20, 100, true) // rclBounds bottom
  dv.setInt32(32, 2646, true) // rclFrame right (.01mm)
  dv.setInt32(36, 2646, true) // rclFrame bottom
  dv.setUint32(40, 0x464d4520, true) // ' EMF' signature
  dv.setUint32(44, 0x00010000, true) // version
  dv.setUint32(48, 108, true) // nBytes
  dv.setUint32(52, 2, true) // nRecords
  dv.setUint16(56, 1, true) // nHandles
  dv.setUint32(88, 14, true) // EMR_EOF
  dv.setUint32(92, 20, true)
  dv.setUint32(96, 0, true) // nPalEntries
  dv.setUint32(100, 16, true) // offPalEntries
  dv.setUint32(104, 20, true) // nSizeLast
  return bytes
}

describe('isMetafileMime', () => {
  it('accepts emf/wmf including x- variants', () => {
    for (const m of ['image/emf', 'image/x-emf', 'image/wmf', 'image/x-wmf']) {
      expect(isMetafileMime(m)).toBe(true)
    }
    expect(isMetafileMime('image/png')).toBe(false)
    expect(isMetafileMime(undefined)).toBe(false)
  })
})

describe('metafileToDataUrl', () => {
  it('returns null for garbage bytes instead of throwing', async () => {
    const garbage = new Uint8Array([0xde, 0xad, 0xbe, 0xef, 1, 2, 3, 4])
    expect(await metafileToDataUrl(garbage, 'image/emf')).toBeNull()
    expect(await metafileToDataUrl(garbage, 'image/wmf')).toBeNull()
    expect(await metafileToDataUrl(new Uint8Array(0), 'image/x-emf')).toBeNull()
  })

  it('returns null for non-metafile mimes', async () => {
    expect(await metafileToDataUrl(minimalEmf(), 'image/png')).toBeNull()
  })

  // Rendering needs OffscreenCanvas/HTMLCanvasElement (renderer process only);
  // the actual pixel output is covered by app-level screenshot verification.
  it.runIf(
    typeof (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas === 'undefined' &&
      typeof (globalThis as { document?: unknown }).document === 'undefined',
  )('degrades to null for a valid EMF when no canvas API exists', async () => {
    expect(await metafileToDataUrl(minimalEmf(), 'image/emf')).toBeNull()
  })
})
