/**
 * Footnote/endnote separator geometry (github run 20260825 samples 18/20/21/22):
 * Word draws a 2in rule on its own row above the note area; the row is the
 * notes part's separator paragraph (--note-sep-h, FOOTNOTE_SEPARATOR_H when the
 * document has none), reserved by the pagination engine, so the CSS must consume
 * exactly that height with the rule --note-sep-rule down the row — a top border
 * used to paint straight through the last body line like an underline.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { FOOTNOTE_SEPARATOR_H } from '../src/renderer/line-metrics'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

function ruleOf(selector: string): string {
  const m = css.match(new RegExp(`${selector.replace(/[.:]/g, '\\$&')}\\s*\\{([^}]*)\\}`))
  expect(m, `rule for ${selector}`).toBeTruthy()
  return m![1]
}

const SEPARATORS = ['.page-gap-notes::before', '.pv-footnotes::before', '.pv-endnote-separator']

describe('note separator CSS', () => {
  it.each(SEPARATORS)('%s draws Word’s 2in rule at the bottom of the reserved row', (sel) => {
    const rule = ruleOf(sel)
    expect(rule).toContain('width: min(2in, 100%)')
    expect(rule).toContain('border-bottom: 1px solid')
    expect(rule).not.toContain('border-top')
    // rule offset + 1px border + spacing consume the row; the fallbacks are the engine's default
    const height = /height:\s*var\(--note-sep-rule,\s*(\d+)px\)/.exec(rule)?.[1]
    expect(height).toBeDefined()
    expect(rule).toContain(
      `margin-bottom: calc(var(--note-sep-h, ${FOOTNOTE_SEPARATOR_H}px) - var(--note-sep-rule, ${height}px) - 1px)`,
    )
  })

  it('canvas end-of-document notes use the same 2in bottom-border rule', () => {
    const rule = ruleOf('.page-notes::before')
    expect(rule).toContain('width: min(2in, 100%)')
    expect(rule).toContain('border-bottom: 1px solid')
  })

  it('note markers render at Word’s superscript scale', () => {
    for (const sel of ['.doc-note-ref', '.pv-footnote sup', '.page-gap-note sup']) {
      expect(ruleOf(sel)).toContain('font-size: 0.65em')
    }
  })
})
