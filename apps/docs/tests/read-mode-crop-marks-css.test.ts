/**
 * Word's print layout marks each text-boundary corner with an L of crop marks.
 * Read mode keeps the paginated paper but hides the editing chrome, and the
 * marks end up floating away from the text, against a scroll offset that is no
 * longer the page's own margin — a reading view should not show them at all.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

const readModeRule = (): string => {
  const m = css.match(
    /\.read-mode \.page-gap-paper::after,\s*\n\.read-mode \.page-wrap:not\(\.paper-mixed\)::before,\s*\n\.read-mode \.page-wrap:not\(\.paper-mixed\)::after \{([^}]*)\}/,
  )
  expect(m, 'read-mode crop-mark rule').toBeTruthy()
  return m![1]
}

describe('read mode hides the paper crop marks', () => {
  it('suppresses the pseudo-elements that draw them', () => {
    expect(readModeRule()).toMatch(/\bcontent: none;/)
  })

  it('keeps the marks in the editing views', () => {
    // the base rules still draw them: read mode is the only suppression
    expect(css).toMatch(/\.page-gap-paper::after,\n\.page-wrap:not\(\.paper-mixed\)::before/)
    expect(css).toMatch(/\.page-wrap:not\(\.paper-mixed\)::before,[\s\S]*?content: '';/)
  })
})
