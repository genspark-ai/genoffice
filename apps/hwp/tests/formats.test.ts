import { describe, expect, it } from 'vitest'
import {
  HWP_EXTENSIONS,
  HWP_RE,
  bytesForSaveFormat,
  ensureHwpSavePath,
  isHwpPath,
  saveFormatForPath,
} from '../src/shared/formats'

describe('HWP_RE', () => {
  it('matches rhwp formats case-insensitively', () => {
    expect(isHwpPath('/tmp/a.hwp')).toBe(true)
    expect(isHwpPath('/tmp/a.HWPX')).toBe(true)
    expect(isHwpPath('C:\\docs\\form.Hml')).toBe(true)
  })

  it('rejects other office formats', () => {
    expect(isHwpPath('/tmp/a.docx')).toBe(false)
    expect(isHwpPath('/tmp/a.hwt')).toBe(false)
    expect(isHwpPath('/tmp/a.hwp.bak')).toBe(false)
  })

  it('lists the same extensions the regex accepts', () => {
    for (const ext of HWP_EXTENSIONS) {
      expect(HWP_RE.test(`file.${ext}`)).toBe(true)
    }
  })
})

describe('save path helpers', () => {
  it('defaults untitled dialog paths to .hwp', () => {
    expect(ensureHwpSavePath('/tmp/note')).toBe('/tmp/note.hwp')
    expect(ensureHwpSavePath('/tmp/note.hwpx')).toBe('/tmp/note.hwpx')
  })

  it('picks export bytes from the save extension', () => {
    const payload = {
      hwp: new Uint8Array([1]),
      hwpx: new Uint8Array([2]),
      hml: new Uint8Array([3]),
    }
    expect(saveFormatForPath('a.HWPX')).toBe('hwpx')
    expect(bytesForSaveFormat('hwp', payload)).toBe(payload.hwp)
    expect(bytesForSaveFormat('hwpx', payload)).toBe(payload.hwpx)
    expect(bytesForSaveFormat('hml', payload)).toBe(payload.hml)
  })
})
