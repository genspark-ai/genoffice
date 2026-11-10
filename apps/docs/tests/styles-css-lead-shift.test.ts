/** leading shift (styles.css --doc-lead-shift) and the anchored-picture counter-shift cascade */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const css = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)
const ruleAt = (selector: string) => {
  const esc = selector.replace(/[.>*+?^$(){}|[\]\\]/g, '\\$&')
  const m = new RegExp(`(^|\\n)${esc}\\s*\\{([^}]*)\\}`).exec(css)
  return m ? { index: m.index, body: m[2] } : null
}
const specificity = (selector: string) => (selector.match(/\.[\w-]+|\[[^\]]+\]/g) ?? []).length

const anchored = '.doc-page .doc-anchor-origin .doc-inline-img-anchor > img'
const pageRel = ".doc-anchor-origin .doc-inline-img-anchor > img[data-page-rel-v='1']"
const pageRelPreview = ".pv-clip .doc-inline-img-anchor > img[data-page-rel-v='1']"

describe('styles.css leading shift', () => {
  it('shifts the paragraph box by --doc-lead-shift', () => {
    // field blocks (TOC lines, PAGE fields) take the document line rule too
    const block = ruleAt(
      '.doc-page :is(p, h1, h2, h3, h4, h5, h6, .doc-li, .doc-textbox-para, .doc-protected-field)',
    )
    expect(block?.body).toMatch(/position:\s*relative/)
    expect(block?.body).toMatch(/top:\s*var\(--doc-lead-shift\)/)
  })

  it('anchored pictures translate back by the shift', () => {
    expect(ruleAt(anchored)?.body).toMatch(
      /translate:\s*0 calc\(-1 \* var\(--doc-lead-shift, 0px\)\)/,
    )
  })

  it('a page-relative picture keeps --page-float-dy and the counter-shift', () => {
    const base = ruleAt(anchored)
    const rel = ruleAt(pageRel)
    expect(rel?.body).toMatch(
      /translate:\s*0 calc\(var\(--page-float-dy, 0px\) - var\(--doc-lead-shift, 0px\)\)/,
    )
    // the cascade must reach this rule: no lower specificity, and later in source
    expect(specificity(pageRel)).toBeGreaterThanOrEqual(specificity(anchored))
    expect(rel!.index).toBeGreaterThan(base!.index)
  })

  it('the preview clone drops the float channel but keeps the counter-shift', () => {
    const rel = ruleAt(pageRel)
    const pv = ruleAt(pageRelPreview)
    expect(pv?.body).toMatch(/translate:\s*0 calc\(-1 \* var\(--doc-lead-shift, 0px\)\)/)
    expect(pv?.body).not.toMatch(/--page-float-dy/)
    expect(specificity(pageRelPreview)).toBeGreaterThanOrEqual(specificity(pageRel))
    expect(pv!.index).toBeGreaterThan(rel!.index)
  })
})
