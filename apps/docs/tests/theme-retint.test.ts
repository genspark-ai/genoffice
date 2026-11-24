import { describe, expect, it } from 'vitest'
import { themeRunColor } from '../src/renderer/editor/theme-retint'

const FACET = { accent1: '90C226', dk2: '3E3D2D' }

describe('themeRunColor', () => {
  it('resolves the run colour reference on the new palette', () => {
    expect(themeRunColor('<w:color w:val="4472C4" w:themeColor="accent1"/>', FACET)).toBe('90C226')
    expect(themeRunColor('<w:b/><w:color w:val="44546A" w:themeColor="text2"/>', FACET)).toBe(
      '3E3D2D',
    )
  })

  it('applies the shade of the reference', () => {
    const shaded = themeRunColor(
      '<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>',
      FACET,
    )
    expect(shaded).toMatch(/^[0-9A-F]{6}$/)
    expect(shaded).not.toBe('90C226')
  })

  it('has nothing to say about a literal colour', () => {
    expect(themeRunColor('<w:color w:val="FF0000"/>', FACET)).toBeNull()
    expect(themeRunColor('<w:b/>', FACET)).toBeNull()
  })
})
