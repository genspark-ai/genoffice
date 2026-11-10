/**
 * A picture-only paragraph is exactly the picture extent in Word. The block
 * picture wrapper must not add editing chrome (border/padding) to the flow:
 * every such picture was 3px taller above and below, spilling captions and
 * paired drawings onto extra pages.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

const ruleBody = (selector: string): string => {
  const esc = selector.replace(/[.[\]'()]/g, '\\$&')
  const m = css.match(new RegExp(`(?:^|\\n)${esc} \\{([^}]*)\\}`))
  expect(m, `rule for ${selector}`).toBeTruthy()
  return m![1]
}

describe('block picture wrapper takes no layout space beyond the picture', () => {
  it('zeroes border and padding on the image block', () => {
    const body = ruleBody(".doc-protected[data-doc-protected='image']")
    expect(body).toMatch(/\n\s*border: 0;/)
    expect(body).toMatch(/\n\s*padding: 0;/)
    expect(body).toMatch(/\n\s*line-height: 0;/)
  })

  it('draws the hover frame as an outline, not a border', () => {
    const body = ruleBody(".doc-protected[data-doc-protected='image']:not(.doc-img-float):hover")
    expect(body).toMatch(/outline: 1px solid var\(--border\)/)
    expect(body).not.toMatch(/\bborder(?:-color)?:/)
  })
})
