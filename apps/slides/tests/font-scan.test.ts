import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'
import { scanFontDirs } from '../src/main/font-scan'

const tmp = mkdtempSync(join(tmpdir(), 'font-scan-'))
afterAll(() => rmSync(tmp, { recursive: true, force: true }))

const flat = join(tmp, 'flat')
mkdirSync(flat)
writeFileSync(join(flat, 'Arial.ttf'), '')
writeFileSync(join(flat, 'Notes.txt'), '')
mkdirSync(join(flat, 'Folder.ttf'))

const base = join(tmp, 'FontCache')
mkdirSync(join(base, '4', 'CloudFonts', 'Montserrat'), { recursive: true })
mkdirSync(join(base, '4', 'CloudFonts', 'Empty'))
mkdirSync(join(base, 'stale'))
writeFileSync(join(base, '4', 'CloudFonts', 'Montserrat', '1001.ttf'), '')
writeFileSync(join(base, '4', 'CloudFonts', 'Montserrat', 'readme.txt'), '')

describe('scanFontDirs', () => {
  it('lists font files of flat dirs and cloud families off the worker', () => {
    const [a, b, c] = scanFontDirs(
      [
        { kind: 'flat', dir: flat },
        { kind: 'flat', dir: join(tmp, 'missing') },
        { kind: 'cloud', base, sub: 'CloudFonts' },
      ],
      5000,
    )
    expect(a).toEqual({ kind: 'flat', files: ['Arial.ttf'] })
    expect(b).toEqual({ kind: 'flat', files: [] })
    const root = join(base, '4', 'CloudFonts')
    expect(c).toEqual({
      kind: 'cloud',
      roots: [{ root, families: [['Montserrat', [join(root, 'Montserrat', '1001.ttf')]]] }],
    })
  })

  it('an unreadable cloud base yields no roots', () => {
    expect(
      scanFontDirs([{ kind: 'cloud', base: join(tmp, 'nope'), sub: 'CloudFonts' }], 5000),
    ).toEqual([{ kind: 'cloud', roots: [] }])
  })

  it('returns at the deadline instead of blocking on unfinished dirs', () => {
    const t0 = Date.now()
    const out = scanFontDirs([{ kind: 'flat', dir: flat }], 0)
    expect(Date.now() - t0).toBeLessThan(1000)
    expect(out).toHaveLength(1)
  })
})
