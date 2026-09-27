/**
 * Print: each preview sheet is size-contained, so an absolutely positioned
 * header/footer or body picture whose offset lies beyond the paper edge stays
 * inside the sheet's fragment. Without it Chromium's shrink-to-fit scaled the
 * whole export down (a header logo anchored 8.8 in from the margin shrank every
 * page of an 8.5 x 13 in document to 0.87x).
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

describe('print sheet containment', () => {
  it('declares contain: size on .pv-page inside the print media block', () => {
    const print = css.slice(css.indexOf('@media print {'))
    const m = print.match(/\n {2}\.pv-page \{([^}]*)\}/)
    expect(m, 'print .pv-page rule').toBeTruthy()
    expect(m![1]).toMatch(/\n\s*break-after: page;/)
    expect(m![1]).toMatch(/\n\s*contain: size;/)
  })
})
