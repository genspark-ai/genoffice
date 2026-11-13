import { describe, expect, it } from 'vitest'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fontDirs, xdgUserFontDir } from '../src/sfnt'

const linux = process.platform === 'linux'

describe('xdgUserFontDir', () => {
  it('prefers XDG_DATA_HOME and appends fonts', () => {
    expect(xdgUserFontDir('/tmp/xdg-data')).toBe('/tmp/xdg-data/fonts')
  })

  it('falls back to ~/.local/share when XDG_DATA_HOME is unset or empty', () => {
    const fallback = join(homedir(), '.local', 'share', 'fonts')
    expect(xdgUserFontDir('')).toBe(fallback)
    expect(xdgUserFontDir(undefined)).toBe(
      process.env.XDG_DATA_HOME ? join(process.env.XDG_DATA_HOME, 'fonts') : fallback,
    )
  })
})

describe('fontDirs', () => {
  it('scans the per-user XDG font dir on Linux', () => {
    if (!linux) return
    // Regression: dropping this dir left families installed "for the current user" (KDE/
    // GNOME installers, manual copies) unknown to metrics while fontconfig still drew them,
    // so layout fell back to heuristic advances and CJK glyphs overlapped their neighbours.
    const dirs = fontDirs()
    expect(dirs).toContain(xdgUserFontDir())
    // legacy per-user dir stays as well, system dirs come first so bundled same-name faces win
    expect(dirs).toContain(join(homedir(), '.fonts'))
    expect(dirs.indexOf('/usr/share/fonts')).toBeLessThan(dirs.indexOf(xdgUserFontDir()))
  })
})
