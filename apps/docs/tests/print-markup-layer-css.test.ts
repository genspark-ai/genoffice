/**
 * Print: the markup strip (comment and revision balloons) paints as its own
 * PDF group. pdfium decides a page's text direction from the page's top-level
 * text objects; balloons beside short body lines tipped it to vertical, and the
 * body text then read back one glyph per line (genoffice#2839). Inside a group
 * the balloons stay out of that vote, so the body reads as it does without
 * markup. The rule is print-only and matches nothing on a page without markup.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

/** top-level `@media print { … }` blocks, by brace matching */
function printBlocks(): string[] {
  const blocks: string[] = []
  let at = css.indexOf('@media print {')
  while (at !== -1) {
    let depth = 0
    let end = css.indexOf('{', at)
    for (; end < css.length; end++) {
      if (css[end] === '{') depth++
      else if (css[end] === '}' && --depth === 0) break
    }
    blocks.push(css.slice(at, end + 1))
    at = css.indexOf('@media print {', end)
  }
  return blocks
}

const markupRules = (text: string) => [...text.matchAll(/(?:^|\n)\s*\.pv-markup \{([^}]*)\}/g)]

describe('print markup layer', () => {
  it('groups the markup strip in print with a sub-1 opacity', () => {
    const rules = printBlocks().flatMap(markupRules)
    expect(rules).toHaveLength(1)
    const opacity = Number(rules[0][1].match(/\bopacity:\s*([\d.]+);/)?.[1])
    expect(opacity).toBeLessThan(1)
    // invisible: under half of one 8-bit level on any channel
    expect((1 - opacity) * 255).toBeLessThan(0.5)
  })

  it('keeps the screen preview strip and the printed sheet ungrouped', () => {
    const screen = printBlocks().reduce((text, block) => text.replace(block, ''), css)
    for (const [, body] of markupRules(screen)) expect(body).not.toMatch(/\bopacity:/)
    for (const block of printBlocks()) {
      for (const [, body] of block.matchAll(/\n\s*\.pv-(?:sheet|page)[^{]*\{([^}]*)\}/g))
        expect(body).not.toMatch(/\bopacity:/)
    }
  })
})
