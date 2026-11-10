/**
 * Flow gaps the paginated preview inherits from the canvas stylesheet: Word
 * gives a table no vertical margin of its own, and an anchored shape keeps
 * its declared extent past the text column.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const css = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)
// Selectors are matched against whole selector lists (a rule may group the
// direct box with the chrome-rotator-nested one), so scan every declaration
// block and test the prelude split on commas. Comments are stripped first —
// they would otherwise ride along on the prelude of the next rule.
const rule = (selector: string): string => {
  const wanted = selector.trim()
  const sheet = css.replace(/\/\*[\s\S]*?\*\//g, '')
  for (const match of sheet.matchAll(/([^{}]+)\{([^}]*)\}/g)) {
    const selectors = match[1].split(',').map((s) => s.trim().replace(/\s+/g, ' '))
    if (selectors.includes(wanted)) return match[2]
  }
  return ''
}

describe('styles.css flow gaps', () => {
  it('body tables carry no vertical margin', () => {
    expect(rule('.doc-page > .doc-table')).toMatch(/margin:\s*0;/)
  })

  it('anchored textboxes are not capped at the column width', () => {
    // the floating box keeps its extent; a single shape's slot is carried by
    // the chrome rotator, whose rect mirrors the box
    expect(rule('.doc-protected-floating > .doc-textbox')).toMatch(/max-width:\s*none/)
    expect(rule('.doc-protected-floating > .box-rotator')).toMatch(/max-width:\s*none/)
  })

  it('the chrome rotator shrink-wraps the shape (inline pivot)', () => {
    // a plain block would stretch to the wrapper's width (an anchor line
    // sharing the paragraph widens it), so the turn would pivot off centre
    expect(rule('.box-rotator')).toMatch(/width:\s*fit-content/)
    expect(rule('.box-rotator')).toMatch(/max-width:\s*100%/)
    // floating boxes keep their declared extent instead of the wrapper cap
    expect(rule('.doc-protected-floating > .box-rotator')).toMatch(/max-width:\s*none/)
  })

  it('a floating rotator carries the canvas page-gap display shift', () => {
    expect(rule('.doc-protected-floating > .box-rotator')).toMatch(
      /translate:\s*0 var\(--page-float-dy/,
    )
  })

  it('charts reveal their direct-child resize handle (they have no rotator)', () => {
    expect(rule('.doc-protected-textboxes:hover > .box-rotator > .box-resize-handle')).toMatch(
      /opacity:\s*1/,
    )
    expect(rule('.doc-protected-chart:hover > .box-resize-handle')).toMatch(/opacity:\s*1/)
    // the rotator-nested chart selector would never match: charts emit the
    // handle straight onto the wrapper
    expect(rule('.doc-protected-chart:hover > .box-rotator > .box-resize-handle')).toBe('')
    expect(rule('.doc-content-editing > .box-resize-handle')).toMatch(/opacity:\s*0/)
  })
})
