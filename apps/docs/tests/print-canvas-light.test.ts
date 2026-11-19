/**
 * Print: the canvas behind the page boxes is the paper color under the light
 * scheme — an A4 sheet is 841.92pt tall, so its last device row is only
 * partly covered by the page box, and a transparent canvas let Chromium's
 * dark-scheme color show through as a 1px dark strip along the bottom of
 * every page when the app theme was dark (#2837).
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

describe('print canvas', () => {
  it('paints the paper behind the page boxes and pins the light scheme', () => {
    const print = css.slice(css.indexOf('@media print {'))
    const scheme = print.match(/\n {2}html \{\n {4}color-scheme: light;\n {2}\}/)
    expect(scheme, 'print light scheme on html').toBeTruthy()
    const canvas = print.match(/\n {2}html,\n {2}body,\n {2}#root \{([^}]*)\}/)
    expect(canvas, 'print canvas rule').toBeTruthy()
    expect(canvas![1]).toMatch(/background: var\(--docs-paper\)/)
  })
})
