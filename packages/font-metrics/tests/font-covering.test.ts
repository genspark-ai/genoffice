import { describe, expect, it } from 'vitest'

import { findFontCovering } from '../src/font-locate'

/** Table tags of a standalone sfnt buffer */
function tableTags(font: Buffer): string[] {
  const n = font.readUInt16BE(4)
  const tags: string[] = []
  for (let i = 0; i < n; i++) tags.push(font.toString('latin1', 12 + i * 16, 16 + i * 16))
  return tags
}

describe('findFontCovering', () => {
  it('returns null for empty text', () => {
    expect(findFontCovering('')).toBeNull()
    expect(findFontCovering('\n')).toBeNull()
  })

  it('finds an installed face covering plain Latin', () => {
    const bytes = findFontCovering('Hello world')
    if (bytes === null && process.platform === 'linux') return // fontless CI container
    expect(bytes).not.toBeNull()
    expect(tableTags(bytes!).some((t) => t === 'glyf' || t === 'CFF ')).toBe(true)
  })

  it('returns null for unassigned codepoints no real font maps', () => {
    expect(findFontCovering('\u0378')).toBeNull()
  })

  it('never picks a color-emoji face', () => {
    const bytes = findFontCovering('🦄')
    if (bytes === null) return // expected on most machines: only color fonts map emoji
    const tags = tableTags(bytes)
    for (const color of ['sbix', 'COLR', 'CBDT', 'CBLC']) expect(tags).not.toContain(color)
  })
})
