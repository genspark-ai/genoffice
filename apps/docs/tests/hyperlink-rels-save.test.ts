import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

// Hyperlinks (and their relationship entries) used to vanish on save:
// regenerated runs merged adjacent same-target links under one rId, and a
// nested table re-read from the DOM lost its trailing empty paragraph and went
// through the plain-text cell patch (first pPr/rPr on every paragraph).
const link = (rId: string, text: string, extraRPr = '') =>
  `<w:hyperlink r:id="${rId}"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/>${extraRPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink>`

const BODY =
  `<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>${link('rId20', 'the site')}<w:r><w:t>.</w:t></w:r></w:p>` +
  // two adjacent hyperlinks to the same URL, each with its own relationship
  `<w:p><w:r><w:t xml:space="preserve">Read </w:t></w:r>${link('rId21', ' ')}${link('rId22', 'the agreement', '<w:u w:val="single"/>')}<w:r><w:t>.</w:t></w:r></w:p>` +
  // tracked insertion wrapping a hyperlink
  `<w:p><w:r><w:t xml:space="preserve">Mail </w:t></w:r><w:ins w:id="7" w:author="A" w:date="2026-01-01T00:00:00Z">${link('rId23', 'us')}</w:ins></w:p>` +
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr>' +
  // hyperlink in a plain cell
  `<w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Contact </w:t></w:r>${link('rId24', 'team')}</w:p></w:tc>` +
  // nested table whose cell ends with an empty paragraph and holds a mailto link
  '<w:tc><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr><w:tc><w:p><w:r><w:t>Jane Doe</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="both"/></w:pPr>${link('rId25', 'jane@example.com')}</w:p><w:p/></w:tc></w:tr></w:tbl><w:p/></w:tc>` +
  '</w:tr></w:tbl>'

const REL = (id: string, target: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${target}" TargetMode="External"/>`
const RELS =
  REL('rId20', 'https://example.com/') +
  REL('rId21', 'https://example.com/agreement.pdf') +
  REL('rId22', 'https://example.com/agreement.pdf') +
  REL('rId23', 'mailto:hello@example.com') +
  REL('rId24', 'mailto:team@example.com') +
  REL('rId25', 'mailto:jane@example.com')

async function openDoc(bodyXml = BODY) {
  const source = await buildDocx({ bodyXml, extraRels: RELS })
  const parsed = await parseDocx(source)
  const host = document.createElement('div')
  document.body.appendChild(host)
  const editor = new Editor({
    element: host,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { source, parsed, editor, host }
}

async function partsOf(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes)
  return {
    doc: await zip.file('word/document.xml')!.async('string'),
    rels: await zip.file('word/_rels/document.xml.rels')!.async('string'),
  }
}

const ALL_IDS = ['rId20', 'rId21', 'rId22', 'rId23', 'rId24', 'rId25']

describe('hyperlinks and their relationships survive saving', () => {
  it('leaves every block byte-exact when nothing was edited', async () => {
    const { source, parsed, editor, host } = await openDoc()
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
    host.remove()
  })

  it('keeps distinct relationships for adjacent same-target links in a regenerated paragraph', async () => {
    const { parsed, editor, host } = await openDoc()
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'Read ') at = pos + 'Read'.length
    })
    expect(at).toBeGreaterThan(-1)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)).insertText('X', at),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const { doc, rels } = await partsOf(await saveDocx(parsed, plan.saveBlocks))
    expect(doc).toContain('<w:hyperlink r:id="rId21">')
    expect(doc).toContain('<w:hyperlink r:id="rId22">')
    for (const id of ALL_IDS) expect(rels).toContain(`Id="${id}"`)
    editor.destroy()
    host.remove()
  })

  it('patches only the edited nested-cell paragraph and keeps its neighbours', async () => {
    const { parsed, editor, host } = await openDoc()
    const nestedTd = host.querySelector<HTMLElement>('.doc-nested-table td')!
    const first = nestedTd.querySelector<HTMLElement>('div')!
    expect(first.textContent).toBe('Jane Doe')
    first.textContent = 'Jane Roe'
    first.dispatchEvent(new Event('input', { bubbles: true }))
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const { doc, rels } = await partsOf(await saveDocx(parsed, plan.saveBlocks))
    expect(doc).toContain('Jane Roe')
    // untouched sibling paragraphs keep their bytes: hyperlink, pPr and the trailing empty paragraph
    expect(doc).toContain(
      `<w:p><w:pPr><w:jc w:val="both"/></w:pPr>${link('rId25', 'jane@example.com')}</w:p><w:p/></w:tc>`,
    )
    for (const id of ALL_IDS) expect(rels).toContain(`Id="${id}"`)
    editor.destroy()
    host.remove()
  })

  it('reads paragraphs below the exact-height clip wrapper', async () => {
    const exactRow = BODY.replace(
      '<w:tr><w:tc><w:p><w:r><w:t>Jane Doe',
      '<w:tr><w:trPr><w:trHeight w:val="500" w:hRule="exact"/></w:trPr><w:tc><w:p><w:r><w:t>Jane Doe',
    )
    expect(exactRow).not.toBe(BODY)
    const { parsed, editor, host } = await openDoc(exactRow)
    const nestedTd = host.querySelector<HTMLElement>('.doc-nested-table td')!
    const first = nestedTd.querySelector<HTMLElement>('.cell-clip > div')!
    expect(first.textContent).toBe('Jane Doe')
    first.textContent = 'Jane Roe'
    first.dispatchEvent(new Event('input', { bubbles: true }))
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const { doc } = await partsOf(await saveDocx(parsed, plan.saveBlocks))
    expect(doc).toContain('Jane Roe')
    expect(doc).toContain(
      `<w:p><w:pPr><w:jc w:val="both"/></w:pPr>${link('rId25', 'jane@example.com')}</w:p><w:p/></w:tc>`,
    )
    editor.destroy()
    host.remove()
  })
})
