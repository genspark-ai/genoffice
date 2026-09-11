import { describe, expect, it } from 'vitest'
import { isHwpPrintSurfaceUrl } from '../src/shared/print-surface'

describe('isHwpPrintSurfaceUrl', () => {
  const opener = 'http://127.0.0.1:4173/'

  it('allows the same-origin studio print surface', () => {
    expect(isHwpPrintSurfaceUrl(opener, 'http://127.0.0.1:4173/rhwp/print.html')).toBe(true)
    expect(isHwpPrintSurfaceUrl(`${opener}rhwp/?chrome=embed`, 'print.html')).toBe(true)
  })

  it('rejects other paths and origins', () => {
    expect(isHwpPrintSurfaceUrl(opener, 'http://127.0.0.1:4173/rhwp/index.html')).toBe(false)
    expect(isHwpPrintSurfaceUrl(opener, 'http://evil.example/rhwp/print.html')).toBe(false)
    expect(isHwpPrintSurfaceUrl(opener, 'https://127.0.0.1:4173/rhwp/print.html')).toBe(false)
  })
})
