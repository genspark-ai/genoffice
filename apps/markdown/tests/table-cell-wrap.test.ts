/**
 * A long inline-code token in a table cell (`split_part(src_schema_id,'.',1)`)
 * used to pin its column at the token's width and push the table past the
 * page. `break-word` wraps it only after the column is sized; `anywhere` lets
 * the cell's min-content shrink too. jsdom computes no layout, so this checks
 * that the stylesheets that ship still give table cells that property.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { buildPrintHtml } from '../src/renderer/export/printHtml'

vi.mock('katex/dist/katex.min.css?inline', () => ({ default: '' }))

const styles = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

/** does some rule in `css` that selects `el` set `overflow-wrap: anywhere`? */
function wrapsAnywhere(css: string, el: Element): boolean {
  return [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)]
    .filter(([, , body]) => /\boverflow-wrap\s*:\s*anywhere\b/.test(body))
    .some(([, sel]) => {
      try {
        return el.matches(sel.replace(/\/\*[\s\S]*?\*\//g, '').trim())
      } catch {
        return false
      }
    })
}

function table(): { th: Element; td: Element } {
  const host = document.createElement('div')
  host.innerHTML =
    '<div class="tableWrapper"><table><tbody>' +
    '<tr><th><p>위치</p></th></tr>' +
    "<tr><td><p><code>split_part(src_schema_id,'.',1)</code></p></td></tr>" +
    '</tbody></table></div>'
  return { th: host.querySelector('th')!, td: host.querySelector('td')! }
}

describe('table cells wrap long tokens', () => {
  it('in the editor', () => {
    const editor = document.createElement('div')
    editor.className = 'doc-editor'
    const { th, td } = table()
    editor.appendChild(th.closest('.tableWrapper')!)
    expect(wrapsAnywhere(styles, th)).toBe(true)
    expect(wrapsAnywhere(styles, td)).toBe(true)
  })

  it('in the print / PDF export', () => {
    const root = document.createElement('div')
    root.appendChild(table().th.closest('.tableWrapper')!)
    const doc = new DOMParser().parseFromString(buildPrintHtml(root, 'Notes'), 'text/html')
    const css = [...doc.querySelectorAll('style')].map((s) => s.textContent ?? '').join('\n')
    expect(wrapsAnywhere(css, doc.querySelector('th')!)).toBe(true)
    expect(wrapsAnywhere(css, doc.querySelector('td')!)).toBe(true)
  })
})
