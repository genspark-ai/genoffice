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
const rule = (selector: string) =>
  new RegExp(`${selector.replace(/[.>*+?^$(){}|[\\]\\]/g, '\\$&')}\\s*\\{([^}]*)\\}`).exec(
    css,
  )?.[1] ?? ''

describe('styles.css flow gaps', () => {
  it('body tables carry no vertical margin', () => {
    expect(rule('.doc-page > .doc-table')).toMatch(/margin:\s*0;/)
  })

  it('anchored textboxes are not capped at the column width', () => {
    expect(rule('.doc-protected-floating > .doc-textbox')).toMatch(/max-width:\s*none/)
  })
})
