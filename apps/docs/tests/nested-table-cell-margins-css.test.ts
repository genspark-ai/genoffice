/**
 * A nested table without w:tblCellMar takes Word's table defaults (0 top/bottom,
 * 108 sides), not the host table's margins: the host's --doc-cell-pad-* custom
 * properties are inherited through the host cell unless the table resets them
 * (corpus: host 43/43 twips made every nested row 4.3pt taller than Word).
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

describe('nested table cell margins', () => {
  it('resets the inherited --doc-cell-pad-* on every table', () => {
    const m = css.match(/\n\.doc-table \{([^}]*)\}/)
    expect(m).toBeTruthy()
    for (const side of ['t', 'r', 'b', 'l']) {
      expect(m![1]).toMatch(new RegExp(`\\n\\s*--doc-cell-pad-${side}: initial;`))
    }
  })
})
