import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

const ribbonSrc = read('src/renderer/components/Ribbon.tsx')
const css = read('src/renderer/styles.css')

/**
 * Compact ribbon (docs parity, #362): ⌥⌘K / Ctrl+Alt+K cycles the slides
 * ribbon between full and an icon-only middle density. The density lives in
 * the shared useRibbonCollapse hook (unit-tested in apps/docs); what these
 * tests pin is the slides side of the contract — the opt-in labels that turn
 * the hook's compact axis on, and the stylesheet rules the density needs,
 * since jsdom applies no stylesheet and a rendering test would be vacuous.
 */
describe('slides ribbon compact density', () => {
  it('opts in: the collapse hook gets the compact labels', () => {
    expect(ribbonSrc).toMatch(
      /useRibbonCollapse\('ai-slides-ribbon-collapsed', \{[^}]*compact: t\('ribbonCompact'\),[^}]*expandFull: t\('ribbonExpand'\),[^}]*\}\)/s,
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

  it('stacked columns go side by side so nothing taller than the band remains', () => {
    expect(css).toMatch(/\.ribbon-compact \.rb-col \{[^}]*flex-direction: row;/s)
  })

  it('the Design tab theme cards shrink instead of clipping', () => {
    expect(css).toMatch(/\.ribbon-compact \.theme-card \{[^}]*height: 30px;/s)
  })
})
