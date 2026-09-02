/**
 * Section-width decorations vs floating-anchor wrappers (prod100r4/049): the
 * wrapper CSS forces width:100%, which outranks `.doc-page .doc-col-block`.
 * The canvas then wraps the wrapper's stray line at the full canvas width
 * while preview clones wrap it at the section width, drifting every page
 * boundary below the block (a body line ends up split across two pages).
 * The decorated wrapper must take --col-w on both sides.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

describe('doc-col-block width on floating wrappers', () => {
  it.each(['.doc-col-block.doc-protected-floating', '.doc-col-block.doc-img-float'])(
    '%s honors the section-width decoration',
    (sel) => {
      const m = css.match(
        new RegExp(`\\.doc-page ${sel.replace(/[.]/g, '\\$&')}[^{]*\\{([^}]*)\\}`),
      )
      expect(m, `rule for ${sel}`).toBeTruthy()
      expect(m![1]).toContain('width: var(--col-w, 100%)')
      expect(m![1]).toContain('max-width: var(--col-w, 100%)')
    },
  )
})
