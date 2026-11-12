/**
 * Table-of-contents update: References → Update Directory and the TOC
 * context menu share findTocRange/updateTocField. The update must
 * regenerate the entries from the live headings for both TOC shapes:
 * the GenOffice TOC field and Word's automatic TOC (a w:sdt gallery
 * control, which has PAGEREF entries and no TOC instruction).
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { generateTocFieldXml, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectTocEntries, findTocRange, updateTocField } from '../src/renderer/editor/toc-update'

const editors = new Set<Editor>()

afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const realAlert = window.alert
beforeEach(() => {
  window.alert = () => {}
})
afterEach(() => {
  window.alert = realAlert
})

const heading = (text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

/** Word automatic TOC: an sdt gallery control (heading + PAGEREF entries + tail) */
const SDT_TOC =
  '<w:sdt>' +
  '<w:sdtPr><w:id w:val="-1"/>' +
  '<w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj>' +
  '</w:sdtPr>' +
  '<w:sdtEndPr/>' +
  '<w:sdtContent>' +
  '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Contents</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Stale One</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Stale Two</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc2 \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>2</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
  '<w:p><w:r><w:t>tail</w:t></w:r></w:p>' +
  '</w:sdtContent>' +
  '</w:sdt>'

async function openDoc(bodyXml: string) {
  const bytes = await buildDocx({ bodyXml })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return { editor, parsed }
}

function tocLineTitles(editor: Editor): string[] {
  const titles: string[] = []
  editor.state.doc.forEach((node) => {
    const field = node.attrs.fieldDisplay as { kind?: string; left?: string } | null
    if (field?.kind === 'tocLine') titles.push(field.left ?? '')
  })
  return titles
}

async function documentXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

describe('findTocRange', () => {
  it('finds a Word sdt TOC and keeps its heading paragraph out of the range', async () => {
    const { editor, parsed } = await openDoc(SDT_TOC + heading('Chapter 1') + heading('Chapter 2'))
    const range = findTocRange(editor.state.doc, parsed.blocks)
    expect(range?.kind).toBe('sdt')
    expect(range?.to).toBeGreaterThan(0)
    // the kept heading is before the replaced range
    expect(editor.state.doc.textBetween(0, range!.from)).toContain('Contents')
  })

  it('finds a TOC field and reports null without any TOC', async () => {
    const [fieldXml] = generateTocFieldXml([{ level: 1, text: 'Chapter 1' }])
    const { editor, parsed } = await openDoc(fieldXml + heading('Chapter 1'))
    expect(findTocRange(editor.state.doc, parsed.blocks)?.kind).toBe('field')

    const plain = await openDoc(heading('Chapter 1'))
    expect(findTocRange(plain.editor.state.doc, plain.parsed.blocks)).toBeNull()
  })
})

describe('updateTocField', () => {
  it('regenerates a Word sdt TOC from the live headings, keeping the heading', async () => {
    const { editor, parsed } = await openDoc(SDT_TOC + heading('Chapter 1') + heading('Chapter 2'))
    expect(updateTocField(editor, parsed.blocks)).toBe(true)
    expect(tocLineTitles(editor)).toEqual(['Chapter 1', 'Chapter 2'])
    expect(editor.state.doc.textContent).toContain('Contents')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    const opens = (xml.match(/<w:sdt>/g) ?? []).length
    const closes = (xml.match(/<\/w:sdt>/g) ?? []).length
    expect(opens).toBe(1)
    expect(closes).toBe(1)
    expect(xml).toContain('<w:t>Contents</w:t>')
    expect(xml).not.toContain('Stale One')
    expect(xml).not.toContain('Stale Two')
    expect(xml).toContain('TOC \\o')
  })

  it('refreshes stale entries of a TOC field', async () => {
    const [fieldXml] = generateTocFieldXml([{ level: 1, text: 'Old Title' }])
    const { editor, parsed } = await openDoc(fieldXml + heading('New Title'))
    expect(findTocRange(editor.state.doc, parsed.blocks)?.kind).toBe('field')
    expect(updateTocField(editor, parsed.blocks)).toBe(true)
    expect(tocLineTitles(editor)).toEqual(['New Title'])
  })

  it('reports false when the document has no table of contents', async () => {
    const { editor, parsed } = await openDoc(heading('Chapter 1'))
    expect(collectTocEntries(editor)).toEqual([{ level: 1, text: 'Chapter 1' }])
    expect(updateTocField(editor, parsed.blocks)).toBe(false)
  })
})
