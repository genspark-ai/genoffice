import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyRevisionSelection, listRevisionEntries } from '../src/renderer/ai/revision-ops'

const AT = 'w:author="A" w:date="2026-01-01T00:00:00Z"'

// issue #2808 repro: a deleted hyperlink (w:del around fldChar + delInstrText)
// plus an unrelated w:ins in the same paragraph
const DELETED_FIELD_P =
  '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>' +
  `<w:del w:id="1" ${AT}>` +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:delInstrText xml:space="preserve"> HYPERLINK &quot;https://example.com&quot; </w:delInstrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:delText>link text</w:delText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:del>' +
  `<w:ins w:id="90" ${AT}><w:r><w:t> other</w:t></w:r></w:ins>` +
  '</w:p>'

async function openWith(bodyXml: string) {
  const source = await buildDocx({ bodyXml })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

async function saveAndRead(
  editor: Editor,
  parsed: Awaited<ReturnType<typeof parseDocx>>,
): Promise<string> {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const zip = await JSZip.loadAsync(bytes)
  return (await zip.file('word/document.xml')!.async('string')).replace(/<w:sectPr[\s\S]*$/, '')
}

describe('revisions inside protected passthrough blocks (issue #2808)', () => {
  it('the deleted field and the plain w:ins both list as pending changes', async () => {
    const { editor } = await openWith(DELETED_FIELD_P)
    const entries = listRevisionEntries(editor.state.doc)
    expect(entries.map((e) => [e.type, e.author, e.recordId])).toEqual([
      ['deletion', 'A', '1'],
      ['insertion', 'A', '90'],
    ])
    expect(entries[0].blockIndex).toBe(0)
    expect(entries[0].text.length).toBeGreaterThan(0)
  })

  it('accept_changes all resolves the whole paragraph: field gone, insertion kept', async () => {
    const { editor, parsed } = await openWith(DELETED_FIELD_P)
    const picked = applyRevisionSelection(editor, { all: true }, 'accept')
    expect('error' in picked).toBe(false)
    const xml = await saveAndRead(editor, parsed)
    expect(xml).not.toContain('delInstrText')
    expect(xml).not.toContain('<w:del ')
    expect(xml).not.toContain('link text')
    expect(xml).toContain(' other')
    // nothing left pending
    expect(listRevisionEntries(editor.state.doc)).toEqual([])
  })

  it('a selective accept (the deleted field only) leaves the insertion pending', async () => {
    const { editor, parsed } = await openWith(DELETED_FIELD_P)
    const picked = applyRevisionSelection(editor, { ids: ['r1'] }, 'accept')
    expect('error' in picked).toBe(false)
    const xml = await saveAndRead(editor, parsed)
    expect(xml).not.toContain('<w:del ')
    expect(xml).not.toContain('link text')
    // the w:ins was not picked: its record survives, still resolvable
    const left = listRevisionEntries(editor.state.doc)
    expect(left.map((e) => e.recordId)).toEqual(['90'])
    const second = applyRevisionSelection(editor, { all: true }, 'accept')
    expect('error' in second).toBe(false)
    expect(listRevisionEntries(editor.state.doc)).toEqual([])
    const xml2 = await saveAndRead(editor, parsed)
    expect(xml2).toContain(' other')
  })

  it('reject brings the deleted text back and drops the insertion', async () => {
    const { editor, parsed } = await openWith(DELETED_FIELD_P)
    const picked = applyRevisionSelection(editor, { all: true }, 'reject')
    expect('error' in picked).toBe(false)
    const xml = await saveAndRead(editor, parsed)
    expect(xml).toContain('link text')
    expect(xml).not.toContain(' other')
    expect(xml).not.toContain('<w:del ')
    expect(listRevisionEntries(editor.state.doc)).toEqual([])
  })

  it('a paragraph without revisions keeps its passthrough untouched', async () => {
    // no-op accept on a clean selection errors instead of touching the block
    const clean = await openWith('<w:p><w:r><w:t>plain</w:t></w:r></w:p>')
    const result = applyRevisionSelection(clean.editor, { all: true }, 'accept')
    expect('error' in result).toBe(true)
    const plan = pmDocToSavePlan(clean.editor.getJSON() as PmNode, clean.parsed.blocks)
    expect(plan.saveBlocks[0]?.kind).toBe('original')
  })
})
