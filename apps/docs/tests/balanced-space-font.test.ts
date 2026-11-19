/**
 * Contract of the bundled balanced-space face (tools/build-balanced-space-font.py):
 * one 0.5em U+0020 glyph inside a 0.8/0.2 line box, so a .doc-balanced-space span
 * advances half the run size without lifting the line.
 */
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { advanceEm, readWoff2 } from './helpers/woff2-metrics'

const font = readWoff2(join(__dirname, '../src/renderer/fonts/GenOfficeBalancedSpace.woff2'))

describe('GenOffice Balanced Space', () => {
  it('maps only U+0020, at a 0.5em advance', () => {
    expect(advanceEm(font, 0x20)).toBe(0.5)
    expect(font.cmap.size).toBe(1)
  })
})
