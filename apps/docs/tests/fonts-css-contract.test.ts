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

/**
 * The writing view's bundled faces (--story-face names them first). The
 * variable Monaspace face must declare the full weight range so bold and
 * the heavier writing weights come from the one file; the Plex statics
 * pair a normal and a bold face.
 */
describe('writing-view faces are bundled', () => {
  it('Monaspace Neon registers the variable face across its weight range', () => {
    const faces = facesOf('Monaspace Neon')
    expect(faces).toHaveLength(1)
    expect(faces[0]).toContain('MonaspaceNeon-VF.ttf')
    expect(faces[0]).toContain('font-weight: 200 800')
    expect(faces[0]).toContain('truetype-variations')
  })

  it('IBM Plex Mono registers a normal and a bold face', () => {
    const faces = facesOf('IBM Plex Mono')
    expect(faces).toHaveLength(2)
    expect(faces.filter((f) => f.includes('IBMPlexMono-Regular.ttf')).join('\n')).toContain(
      'font-weight: normal',
    )
    expect(faces.filter((f) => f.includes('IBMPlexMono-Bold.ttf')).join('\n')).toContain(
      'font-weight: bold',
    )
  })

  it('the bundled files exist next to fonts.css', () => {
    const dir = join(__dirname, '../src/renderer/fonts')
    for (const file of [
      'MonaspaceNeon-VF.ttf',
      'IBMPlexMono-Regular.ttf',
      'IBMPlexMono-Bold.ttf',
    ]) {
      expect(() => readFileSync(join(dir, file))).not.toThrow()
    }
  })
})
