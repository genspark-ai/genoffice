/**
 * Weight contract of the JP Mincho fallback alias (fonts.css): Word's
 * Japanese serif faces ship no bold and Word synthesizes uniform thick
 * strokes, so the alias must register weight-normal only. A real W6 bold
 * face keeps hairline horizontals — bold U+4E00 reads as an arrow at body sizes.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/fonts/fonts.css'), 'utf8')

function facesOf(family: string): string[] {
  return [...css.matchAll(/@font-face\s*\{[^}]*\}/g)]
    .map((m) => m[0])
    .filter((f) => f.includes(`font-family: '${family}'`))
}

describe('JP Mincho fallback aliases synthesize bold', () => {
  for (const family of ['GenOffice Hiragino Mincho', 'GenOffice MS Mincho']) {
    it(`${family} registers weight-normal faces only`, () => {
      const faces = facesOf(family)
      expect(faces.length).toBeGreaterThan(0)
      for (const face of faces) expect(face).toContain('font-weight: normal')
    })
  }
})
