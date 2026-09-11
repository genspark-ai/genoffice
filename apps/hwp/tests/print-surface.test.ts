import { describe, expect, it } from 'vitest'
import { isHwpPrintSurfaceUrl, printSurfaceLoadUrl } from '../src/shared/print-surface'

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

describe('printSurfaceLoadUrl', () => {
  const opener = 'http://127.0.0.1:4173/'
  const print = 'http://127.0.0.1:4173/rhwp/print.html'

  it('loads print.html when Electron left the popup blank', () => {
    expect(printSurfaceLoadUrl(opener, print, 'about:blank')).toBe(print)
    expect(printSurfaceLoadUrl(opener, print, 'about:blank#blocked')).toBe(print)
    expect(printSurfaceLoadUrl(`${opener}rhwp/?chrome=embed`, 'print.html', '')).toBe(print)
  })

  it('does not reload a popup that already has the print surface', () => {
    expect(printSurfaceLoadUrl(opener, print, print)).toBe(null)
  })

  it('ignores other window.open targets', () => {
    expect(printSurfaceLoadUrl(opener, `${opener}rhwp/index.html`, 'about:blank')).toBe(null)
  })
})
