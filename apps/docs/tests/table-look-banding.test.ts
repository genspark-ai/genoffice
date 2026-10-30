// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import type { ParsedDocFull, StyleInfo, TableStyleDisplay } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const tableXml = (look: string) =>
  `<w:tbl><w:tblPr><w:tblStyle w:val="Banded"/><w:tblW w:w="6000" w:type="dxa"/>${look}</w:tblPr>` +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  [0, 1, 2]
    .map(
      (r) =>
        '<w:tr>' +
        `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>r${r}</w:t></w:r></w:p></w:tc>`.repeat(
          2,
        ) +
        '</w:tr>',
    )
    .join('') +
  '</w:tbl>'

async function tableEl(look: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml: tableXml(look) }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return editor.view.dom.querySelector('table')!
}

describe('w:tblLook opt-outs reach the table element', () => {
  it('noHBand / firstRow off become data attributes', async () => {
    const off = await tableEl(
      '<w:tblLook w:val="0600" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>',
    )
    expect(off.getAttribute('data-tbl-nohband')).toBe('1')
    expect(off.getAttribute('data-tbl-nofirstrow')).toBe('1')
    const on = await tableEl('<w:tblLook w:val="04A0"/>')
    expect(on.hasAttribute('data-tbl-nohband')).toBe(false)
    expect(on.hasAttribute('data-tbl-nofirstrow')).toBe(false)
  })

  it('style band and header rules skip opted-out tables', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('Banded', {
      styleId: 'Banded',
      name: 'Banded',
      type: 'table',
      tableDisplay: {
        band1Fill: 'cccccc',
        firstRow: { fill: '1f1f1f' },
      } as TableStyleDisplay,
    } as StyleInfo)
    const css = docStyleCss({ styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull)
    const sel = '.doc-page table[data-tbl-style="Banded"]'
    expect(css).toContain(
      `${sel}:not([data-tbl-nohband]):not([data-tbl-nofirstrow]) tr:nth-child(even) td, ${sel}:not([data-tbl-nohband])[data-tbl-nofirstrow] tr:nth-child(odd) td { background:#cccccc }`,
    )
    expect(css).toContain(`${sel}:not([data-tbl-nofirstrow]) tr:first-child td`)
    expect(css).not.toMatch(/table\[data-tbl-style="Banded"\] tr:nth-child\(even\)/)
  })
})
