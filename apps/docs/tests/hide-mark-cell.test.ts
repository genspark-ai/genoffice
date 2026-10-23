import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { renderTableSpec } from '../src/renderer/editor/protected-render'
import { indexedRowBoxes } from '../src/renderer/pagination-lines'

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
    const rule =
      /\.doc-table td\[data-hide-mark\] > \.doc-p-empty:not\(\.doc-p-pic\):only-child[^{]*\{([^}]*)\}/.exec(
        css,
      )
    expect(rule).not.toBeNull()
    expect(rule![1]).toContain('height: 0')
    expect(rule![1]).toContain('line-height: 0')
    expect(rule![1]).toContain('margin: 0 !important')
  })

  // spacer-gif rows (genoffice#2117): the picture-only paragraph is textless too,
  // but its line is the picture extent and must not collapse with the cell mark
  it('leaves a picture-only paragraph in a hideMark cell at its picture height', async () => {
    const xml = TABLE_XML.replace('<w:p/>', IMAGE_PARAGRAPH_XML)
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml, withImage: true }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const p = editor.view.dom.querySelector('td[data-hide-mark] > .doc-p-empty')!
    expect(p.classList.contains('doc-p-pic')).toBe(true)
    expect(css).toContain('td[data-hide-mark] > .doc-p-empty:not(.doc-p-pic):only-child')
  })
})

describe('indexedRowBoxes', () => {
  it('keeps a collapsed 0px first row paired with its own tr', () => {
    // rows 0 and 1 measure 0px (collapsed spacer rows); the long row 2 starts at 0
    // and must stay the third box, not be dropped onto trs[0]
    expect(indexedRowBoxes([0, 0, 0], 1200)).toEqual([
      { offsetInBlock: 0, height: 0 },
      { offsetInBlock: 0, height: 0 },
      { offsetInBlock: 0, height: 1200 },
    ])
  })

  it('tiles the block height from the measured row tops', () => {
    expect(indexedRowBoxes([1, 40, 80], 100)).toEqual([
      { offsetInBlock: 0, height: 40 },
      { offsetInBlock: 40, height: 40 },
      { offsetInBlock: 80, height: 20 },
    ])
  })

  it('clamps offsets monotonic and within the block', () => {
    expect(indexedRowBoxes([0, 50, 30, 500], 100)).toEqual([
      { offsetInBlock: 0, height: 50 },
      { offsetInBlock: 50, height: 0 },
      { offsetInBlock: 50, height: 50 },
      { offsetInBlock: 100, height: 0 },
    ])
  })
})

describe('VML rule alone in its paragraph', () => {
  it('hides the separator image ProseMirror puts before the rule', () => {
    expect(css).toMatch(
      /\.ProseMirror-separator:has\(\+ \.doc-inline-rule\)\s*\{\s*display: none !important;/,
    )
  })
})
