/**
 * Comment anchor tint (prod_041, sas run 20260831): Word shades commented text
 * with a pale author tint — #ECFDD7 over white paper with #82AB51 decorations
 * (measured from Word for Mac PDF export). The anchor used to be branded
 * orange, which read as a highlight-color mismatch in fidelity diffs.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

function varValue(name: string): string {
  const m = css.match(new RegExp(`${name}:\\s*([^;]+);`))
  expect(m, `definition of ${name}`).toBeTruthy()
  return m![1].trim()
}

describe('comment anchor colors', () => {
  it('anchor tint composites over white paper to Word’s #ECFDD7', () => {
    const m = /rgb\((\d+) (\d+) (\d+) \/ (\d+)%\)/.exec(varValue('--docs-paper-comment-bg'))
    expect(m).toBeTruthy()
    const [r, g, b, pct] = m!.slice(1).map(Number)
    const a = pct / 100
    const over = (c: number) => Math.round(255 * (1 - a) + c * a)
    expect([over(r), over(g), over(b)]).toEqual([0xec, 0xfd, 0xd7])
  })

  it('anchor line matches Word’s author-color decoration', () => {
    expect(varValue('--docs-paper-comment-line')).toBe('#82ab51')
  })
})
