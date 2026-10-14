import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const css = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)

// a degenerate one-cell table (HTML import): the only content is an empty
// paragraph and the cell carries w:hideMark, so Word sizes the row by its margins
const TABLE_XML =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="76"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>' +
  '<w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/></w:tcMar>' +
  '<w:hideMark/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'

describe('w:hideMark cells', () => {
  it('marks the editor cell so the empty end-of-cell paragraph collapses', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE_XML }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const td = editor.view.dom.querySelector('td')!
    expect(td.getAttribute('data-hide-mark')).toBe('1')
    expect(td.querySelector(':scope > .doc-p-empty:only-child')).not.toBeNull()
  })

  it('marks read-only table cells the same way', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE_XML }))
    const spec = JSON.stringify(renderTableSpec(parsed.blocks[0].table!))
    expect(spec).toContain('"data-hide-mark":"1"')
  })

  it('zeroes the lone empty paragraph of a hideMark cell', () => {
    const rule = /\.doc-table td\[data-hide-mark\] > \.doc-p-empty:only-child[^{]*\{([^}]*)\}/.exec(
      css,
    )
    expect(rule).not.toBeNull()
    expect(rule![1]).toContain('height: 0')
    expect(rule![1]).toContain('line-height: 0')
    expect(rule![1]).toContain('margin: 0 !important')
  })
})

describe('VML rule alone in its paragraph', () => {
  it('hides the separator image ProseMirror puts before the rule', () => {
    expect(css).toMatch(
      /\.ProseMirror-separator:has\(\+ \.doc-inline-rule\)\s*\{\s*display: none !important;/,
    )
  })
})
