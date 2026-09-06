/**
 * Word paints automatic-colour text white on dark shading (cell, paragraph, run
 * or table-style fill). The renderer marks such elements with data-ink="light";
 * styles.css flips the inherited ink there while explicit colours stay inline.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import type { ParsedDocFull, StyleInfo, TableStyleDisplay } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { DARK_PAPER_HEX } from '../src/renderer/editor/dark-page'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { fillInk, isDarkFill } from '../src/renderer/editor/shading-ink'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const STYLES_CSS = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)

const cell = (fill: string, runs: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="${fill}"/></w:tcPr><w:p>${runs}</w:p></w:tc>`
const run = (text: string, rPr = '') => `<w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r>`

const BODY_XML =
  '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>' +
  cell('000000', run('auto on black')) +
  cell('000000', run('red on black', '<w:color w:val="FF0000"/>')) +
  cell('F2F2F2', run('auto on light')) +
  '</w:tr></w:tbl>' +
  '<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="1F3864"/></w:pPr>' +
  run('para on navy') +
  '</w:p>' +
  '<w:p>' +
  run('run on black', '<w:shd w:val="clear" w:color="auto" w:fill="000000"/>') +
  run('run on yellow', '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>') +
  '</w:p>'

async function openDoc() {
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY_XML }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

describe('isDarkFill', () => {
  it('matches the Word probe cut (77R + 151G + 28B) / 256 < 60', () => {
    for (const dark of ['000000', '3B3B3B', 'C60000', '006500', '0000FF', '800080', '1F3864'])
      expect(isDarkFill(dark), dark).toBe(true)
    for (const light of ['3C3C3C', 'C80000', '006800', 'FF0000', '008000', '2A6EBB', 'FFFFFF'])
      expect(isDarkFill(light), light).toBe(false)
    expect(isDarkFill('#000000')).toBe(true)
    expect(isDarkFill(null)).toBe(false)
    expect(isDarkFill('auto')).toBe(false)
  })
})

describe('auto ink on dark shading', () => {
  it('marks dark cells, paragraphs and runs; explicit colours and light fills stay put', async () => {
    const editor = await openDoc()
    const dom = editor.view.dom
    const tds = [...dom.querySelectorAll('td')]
    expect(tds).toHaveLength(3)
    expect(tds[0].getAttribute('data-ink')).toBe('light')
    expect(tds[0].querySelector('span[style*="color"]')).toBeNull()
    expect(tds[1].getAttribute('data-ink')).toBe('light')
    expect(tds[1].querySelector('span')!.style.color).toBe('rgb(255, 0, 0)')
    expect(tds[2].getAttribute('data-ink')).toBe('dark')

    const paras = [...dom.querySelectorAll('p')].filter((p) => p.closest('td') === null)
    expect(paras[0].getAttribute('data-ink')).toBe('light')
    expect(paras[0].style.backgroundColor).toBe('rgb(31, 56, 100)')
    const spans = [...paras[1].querySelectorAll('span[data-doc-style]')]
    expect(spans.map((s) => s.getAttribute('data-ink'))).toEqual(['light', 'dark'])
    editor.destroy()
  })

  it('the nearest fill wins: a light cell inside a dark table and a light run on a dark paragraph reset the ink', () => {
    expect(fillInk('000000')).toBe('light')
    expect(fillInk('F2F2F2')).toBe('dark')
    expect(fillInk('auto')).toBeUndefined()
    expect(fillInk(null)).toBeUndefined()
    // only nested light fills reset: a light fill on its own must not override an inherited colour
    expect(STYLES_CSS).not.toMatch(/\n\[data-ink='dark'\]\s*\{/)
    const dark = /\[data-ink='light'\] \[data-ink='dark'\]\s*\{([^}]*)\}/.exec(STYLES_CSS)!
    expect(dark[1]).toContain('--docs-paper-ink: #000')
    expect(dark[1]).toContain('color: var(--docs-paper-ink)')
    expect(STYLES_CSS).toMatch(
      /\.page-dark \[data-ink='light'\] \[data-ink='dark'\]:where\(:not\(\.doc-textbox-filled \*\)\)\s*\{\s*--docs-paper-ink: #ffffff;/,
    )
  })

  it('styles.css flips the ink on data-ink="light" and re-darkens it on the dark page', () => {
    const light = /\[data-ink='light'\]\s*\{([^}]*)\}/.exec(STYLES_CSS)!
    expect(light[1]).toContain('--docs-paper-ink: #fff')
    expect(light[1]).toContain('color: var(--docs-paper-ink)')
    expect(STYLES_CSS).toMatch(
      /\.page-dark \[data-ink='light'\]:where\(:not\(\.doc-textbox-filled \*\)\)\s*\{\s*--docs-paper-ink: var\(--docs-paper\);/,
    )
  })

  it('table-style dark fills get the same ink through docStyleCss', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('DarkHead', {
      styleId: 'DarkHead',
      name: 'DarkHead',
      type: 'table',
      tableDisplay: {
        band1Fill: 'd9d9d9',
        firstRow: { fill: '1f1f1f', bold: true },
      } as TableStyleDisplay,
    } as StyleInfo)
    styles.set('RedHead', {
      styleId: 'RedHead',
      name: 'RedHead',
      type: 'table',
      tableDisplay: { firstRow: { fill: '1f1f1f', color: 'ff0000' } } as TableStyleDisplay,
    } as StyleInfo)
    styles.set('DarkTable', {
      styleId: 'DarkTable',
      name: 'DarkTable',
      type: 'table',
      tableDisplay: { fill: '000000', band1Fill: 'eeeeee' } as TableStyleDisplay,
    } as StyleInfo)
    const css = docStyleCss({ styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull)
    expect(css).toContain(
      `.doc-page table[data-tbl-style="DarkTable"] td:not([data-ink]), .doc-page table[data-tbl-style="DarkTable"] th:not([data-ink]) { --docs-paper-ink:#fff;color:var(--docs-paper-ink) }`,
    )
    expect(css).toContain(
      `.doc-page table[data-tbl-style="DarkTable"] tr:nth-child(even) td:not([data-ink]) { --docs-paper-ink:#000;color:var(--docs-paper-ink) }`,
    )
    const light = '--docs-paper-ink:#fff;color:var(--docs-paper-ink)'
    const reset = '--docs-paper-ink:#000;color:var(--docs-paper-ink)'
    const head = 'table[data-tbl-style="DarkHead"] tr:first-child'
    expect(css).toContain(`${head} td, .doc-page ${head} th { background:#1f1f1f;font-weight:600 }`)
    // cells with their own fill (data-ink) are left to that attribute: the nearest fill wins
    expect(css).toContain(
      `${head} td:not([data-ink]), .doc-page ${head} th:not([data-ink]) { ${light} }`,
    )
    expect(css).toContain(`.page-dark .doc-page ${head} td:not([data-ink])`)
    expect(css).toContain(`--docs-paper-ink:${DARK_PAPER_HEX}`)
    expect(css).toContain(`background:#d9d9d9 }`)
    // the light band has no dark whole-table fill above it: reset only under a dark data-ink ancestor
    expect(css).not.toContain(
      `.doc-page table[data-tbl-style="DarkHead"] tr:nth-child(even) td:not([data-ink]) {`,
    )
    expect(css).toContain(
      `.doc-page :is([data-ink='light'] table, table[data-ink='light'])[data-tbl-style="DarkHead"] tr:nth-child(even) td:not([data-ink]) { ${reset} }`,
    )
    // light data-ink fills inside the style's dark cells (only those) take the paper ink back
    expect(css).toContain(
      `${head} td:not([data-ink]) [data-ink='dark'], .doc-page ${head} th:not([data-ink]) [data-ink='dark'] { ${reset} }`,
    )
    expect(css).not.toContain(`table[data-tbl-style="DarkHead"] [data-ink='dark'] {`)
    expect(css).toContain(`--docs-paper-ink:#ffffff`)
    expect(css).toContain(`background:#1f1f1f;color:#ff0000 }`)
    expect(css).not.toContain(`table[data-tbl-style="RedHead"] tr:first-child td:not([data-ink])`)
    expect(css).not.toContain(
      `table[data-tbl-style="RedHead"] tr:first-child td:not([data-ink]) [data-ink`,
    )
  })
})
