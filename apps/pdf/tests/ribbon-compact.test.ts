import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

const appSrc = read('src/renderer/App.tsx')
const css = read('src/renderer/styles.css')

/**
 * Compact ribbon (docs parity, #362): ⌥⌘K / Ctrl+Alt+K cycles the pdf ribbon
 * between full and an icon-only middle density. The density lives in the
 * shared useRibbonCollapse hook (unit-tested in apps/docs); what these tests
 * pin is the pdf side of the contract — the opt-in labels, the stylesheet
 * rules the density needs, and the markup shape the CSS relies on (every
 * rb-big label is a span, so the visually-hidden selector can reach it),
 * since jsdom applies no stylesheet and a rendering test would be vacuous.
 */
describe('pdf ribbon compact density', () => {
  it('opts in: the collapse hook gets the compact labels', () => {
    expect(appSrc).toMatch(
      /useRibbonCollapse\('genoffice-pdf-ribbon-collapsed', \{[^}]*compact: t\('ribbonCompact'\),[^}]*expandFull: t\('ribbonExpand'\),[^}]*\}\)/s,
    )
  })

  it('the compact band is a shorter, still fixed-height band', () => {
    expect(css).toMatch(/\.ribbon-compact \.ribbon-body \{[^}]*height: 45px;/s)
  })

  it('big buttons drop to icon-only with the label kept for accessibility', () => {
    // the label is visually hidden, not display: none — the accessible name
    // and the e2e suite's getByRole('button', { name }) selectors read it
    expect(css).toMatch(
      /\.ribbon-compact \.rb-big > span:not\(\.rb-big-icon\) \{[^}]*clip-path: inset\(50%\);/s,
    )
    expect(css).not.toMatch(/\.ribbon-compact[^{]*\{[^}]*display: none/s)
  })

  it('every rb-big label is span-wrapped, the shape the compact selector needs', () => {
    // a bare {t('…')} text child would escape the visually-hidden rule and
    // paint over the icon row; the shape to match is icon span → label span
    const bare = /<span className="rb-big-icon[^"]*">[\s\S]*?<\/span>\s*\{t\(/
    expect(bare.test(appSrc)).toBe(false)
  })

  it('stacked columns go side by side so nothing taller than the band remains', () => {
    expect(css).toMatch(/\.ribbon-compact \.rb-col \{[^}]*flex-direction: row;/s)
  })
})
