import { describe, expect, it } from 'vitest'

import { winMaterialSupported } from '../src/main/window-material'

describe('winMaterialSupported', () => {
  it('needs Windows 11 22H2 or later', () => {
    expect(winMaterialSupported('win32', '10.0.22621')).toBe(true)
    expect(winMaterialSupported('win32', '10.0.26100')).toBe(true)
    expect(winMaterialSupported('win32', '10.0.22000')).toBe(false)
    expect(winMaterialSupported('win32', '10.0.19045')).toBe(false)
  })

  it('is off on other platforms and on an unreadable release', () => {
    expect(winMaterialSupported('darwin', '24.0.0')).toBe(false)
    expect(winMaterialSupported('linux', '6.8.0')).toBe(false)
    expect(winMaterialSupported('win32', 'garbage')).toBe(false)
  })
})
