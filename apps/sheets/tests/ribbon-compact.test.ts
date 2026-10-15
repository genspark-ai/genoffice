import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

const shellSrc = read('src/renderer/ExcelShell.tsx')
const css = read('src/renderer/styles.css')

/**
 * Compact ribbon (docs parity, #362): ⌥⌘K / Ctrl+Alt+K cycles the sheets
 * ribbon between full and an icon-only middle density. The density lives in
 * the shared useRibbonCollapse hook (unit-tested in apps/docs); what these
 * tests pin is the sheets side of the contract — the opt-in labels and the
 * stylesheet rules the density needs, since jsdom applies no stylesheet and
 * a rendering test would be vacuous. The no-clipping guarantee itself is the
 * e2e spec (e2e/sheets-ribbon-compact-clipping.spec.ts), the same split the
 * docs compact band uses.
 */
describe('sheets ribbon compact density', () => {
  it('opts in: the collapse hook gets the compact labels', () => {
    expect(shellSrc).toMatch(
      /useRibbonCollapse\('ai-sheets-ribbon-collapsed', \{[^}]*compact: t\('appRibbonCompact'\),[^}]*expandFull: t\('appRibbonExpand'\),[^}]*\}\)/s,
    )
  })

  it('The compact band is a shorter, still fixed-height band', () => {
    expect(css).toMatch(/\.ribbon-compact \.ribbon \{[^}]*height: 48px;/s)
  })

  it('tall tools drop to icon-only with the label kept for accessibility', () => {
    // the label is visually hidden, not display: none — the accessible name
    // and the e2e suite's getByRole('button', { name }) selectors read it
    expect(css).toMatch(
      /\.ribbon-compact \.ribbon-tool\.large > span:not\(\.tool-icon-row\) \{[^}]*clip-path: inset\(50%\);/s,
    )
    expect(css).not.toMatch(/\.ribbon-compact[^{]*\{[^}]*display: none/s)
  })

  it('stacked row/grid groups go side by side so nothing taller than the band remains', () => {
    expect(css).toMatch(/\.ribbon-compact \.ribbon-rows,[^{]*\{[^}]*flex-direction: row;/s)
    expect(css).toMatch(/\.ribbon-compact \.tool-stack \{[^}]*grid-auto-flow: column;/s)
    expect(css).toMatch(/\.ribbon-compact \.chart-grid \{[^}]*grid-auto-flow: column;/s)
  })
})
