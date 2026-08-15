import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'

/**
 * A Word automatic TOC is one w:sdt whose sdtContent holds a
 * heading paragraph + N PAGEREF field paragraphs. The engine splits it into
 * one block per paragraph; the save plan must keep the single sdt balanced
 * whichever member is edited, split, or deleted.
 */

const TOC_ENTRY = (n: number) =>
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
  `<w:r><w:t>Chapter ${n}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc${n} \\h </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${n}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:p>'

const TOC_SDT =
  '<w:sdt>' +
  '<w:sdtPr><w:id w:val="-1"/>' +
  '<w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj>' +
  '</w:sdtPr>' +
  '<w:sdtEndPr/>' +
  '<w:sdtContent>' +
  '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Contents</w:t></w:r></w:p>' +
  Array.from({ length: 13 }, (_, i) => TOC_ENTRY(i + 1)).join('') +
  '<w:p><w:r><w:t>tail</w:t></w:r></w:p>' +
  '</w:sdtContent>' +
  '</w:sdt>'

async function documentXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

const sdtOpenCount = (xml: string) => (xml.match(/<w:sdt>/g) ?? []).length
const sdtCloseCount = (xml: string) => (xml.match(/<\/w:sdt>/g) ?? []).length

async function openToc() {
  const bytes = await buildDocx({ bodyXml: TOC_SDT })
  const parsed = await parseDocx(bytes)
  const doc = blocksToPmDoc(parsed.blocks)
  return { bytes, parsed, doc }
}

describe('multi-paragraph TOC sdt in the editor', () => {
  it('renders all 15 sdtContent paragraphs (1 heading + 13 entries + 1 tail)', async () => {
    const { doc } = await openToc()
    expect(doc.content?.length).toBe(15)
    const protectedNodes = doc.content!.filter((n) => n.type === 'docProtected')
    expect(protectedNodes.length).toBe(13)
    expect(protectedNodes.every((n) => n.attrs?.fieldDisplay || n.attrs?.previewText)).toBe(true)
  })

  it('untouched TOC round-trips byte-identical', async () => {
    const { bytes, parsed, doc } = await openToc()
    const plan = pmDocToSavePlan(doc, parsed.blocks)
    expect(plan.saveBlocks.every((b) => b.kind === 'original')).toBe(true)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
  })

  it('deleting the last member re-attaches the sdt close to the new last member', async () => {
    const { parsed, doc } = await openToc()
    doc.content!.pop()
    const plan = pmDocToSavePlan(doc, parsed.blocks)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(sdtOpenCount(xml)).toBe(1)
    expect(sdtCloseCount(xml)).toBe(1)
    expect(xml).not.toContain('tail')
    expect(xml.indexOf('Chapter 13')).toBeLessThan(xml.indexOf('</w:sdtContent>'))
  })

  it('deleting the first member re-attaches the sdt open to the new first member', async () => {
    const { parsed, doc } = await openToc()
    doc.content!.shift()
    const plan = pmDocToSavePlan(doc, parsed.blocks)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(sdtOpenCount(xml)).toBe(1)
    expect(sdtCloseCount(xml)).toBe(1)
    expect(xml).not.toContain('<w:t>Contents</w:t>')
    expect(xml.indexOf('<w:sdtContent>')).toBeLessThan(xml.indexOf('Chapter 1<'))
  })

  it('a split twin of a boundary paragraph does not duplicate shell bytes', async () => {
    const { parsed, doc } = await openToc()
    const heading = doc.content![0]
    const twin: PmNode = {
      ...heading,
      content: [{ type: 'text', text: 'twin half' }],
    }
    doc.content!.splice(1, 0, twin)
    const plan = pmDocToSavePlan(doc, parsed.blocks)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(sdtOpenCount(xml)).toBe(1)
    expect(sdtCloseCount(xml)).toBe(1)
    const twinAt = xml.indexOf('twin half')
    expect(twinAt).toBeGreaterThan(xml.indexOf('<w:sdtContent>'))
    expect(twinAt).toBeLessThan(xml.indexOf('</w:sdtContent>'))
  })

  it('editing only the heading keeps one balanced sdt and all entries', async () => {
    const { parsed, doc } = await openToc()
    doc.content![0] = {
      ...doc.content![0],
      content: [{ type: 'text', text: 'Table of Contents (edited)' }],
    }
    const plan = pmDocToSavePlan(doc, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(sdtOpenCount(xml)).toBe(1)
    expect(sdtCloseCount(xml)).toBe(1)
    expect(xml.indexOf('Table of Contents (edited)')).toBeGreaterThan(xml.indexOf('<w:sdtContent>'))
    for (let n = 1; n <= 13; n++) expect(xml).toContain(`Chapter ${n}`)
  })
})

describe('nested cover-page sdt in the editor', () => {
  const NESTED_SDT =
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/></w:docPartObj></w:sdtPr><w:sdtContent>' +
    '<w:sdt><w:sdtPr><w:alias w:val="Title"/><w:text/></w:sdtPr><w:sdtContent>' +
    '<w:p><w:r><w:t>Resume Title</w:t></w:r></w:p>' +
    '</w:sdtContent></w:sdt>' +
    '<w:p><w:r><w:t>plain between</w:t></w:r></w:p>' +
    '<w:sdt><w:sdtPr><w:alias w:val="Company"/><w:text/></w:sdtPr><w:sdtContent>' +
    '<w:p><w:r><w:t>ACME Corp</w:t></w:r></w:p>' +
    '</w:sdtContent></w:sdt>' +
    '</w:sdtContent></w:sdt>'

  async function openNested() {
    const bytes = await buildDocx({ bodyXml: NESTED_SDT })
    const parsed = await parseDocx(bytes)
    const doc = blocksToPmDoc(parsed.blocks)
    return { bytes, parsed, doc }
  }

  it('untouched nested sdt round-trips byte-identical', async () => {
    const { bytes, parsed, doc } = await openNested()
    const plan = pmDocToSavePlan(doc, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
  })

  it('deleting the first inner field keeps the sdt tree balanced', async () => {
    const { parsed, doc } = await openNested()
    // drop "Resume Title": its member carries the inner </sdt><sdt> wrapper fragment
    const pruned: PmNode = { ...doc, content: doc.content!.slice(1) }
    const plan = pmDocToSavePlan(pruned, parsed.blocks)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(sdtOpenCount(xml)).toBe(sdtCloseCount(xml))
    expect(xml).toContain('plain between')
    expect(xml).toContain('ACME Corp')
    expect(xml).not.toContain('Resume Title')
  })

  it('deleting the middle plain paragraph keeps the sdt tree balanced', async () => {
    const { parsed, doc } = await openNested()
    const pruned: PmNode = { ...doc, content: [doc.content![0], doc.content![2]] }
    const plan = pmDocToSavePlan(pruned, parsed.blocks)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(sdtOpenCount(xml)).toBe(sdtCloseCount(xml))
    expect(xml).toContain('Resume Title')
    expect(xml).toContain('ACME Corp')
  })

  it('deleting the last inner field still emits the outer close', async () => {
    // regression: middle survivors carry fragment closeXml, which used to satisfy
    // the group-close check and drop the real outer close entirely
    const { parsed, doc } = await openNested()
    const pruned: PmNode = { ...doc, content: doc.content!.slice(0, 2) }
    const plan = pmDocToSavePlan(pruned, parsed.blocks)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(sdtOpenCount(xml)).toBe(sdtCloseCount(xml))
    expect(xml).toContain('Resume Title')
    expect(xml).not.toContain('ACME Corp')
  })

  it('deleting several leading members keeps the wrapper fragments in document order', async () => {
    const { parsed, doc } = await openNested()
    const pruned: PmNode = { ...doc, content: doc.content!.slice(2) }
    const plan = pmDocToSavePlan(pruned, parsed.blocks)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(sdtOpenCount(xml)).toBe(sdtCloseCount(xml))
    expect(xml).toContain('ACME Corp')
    // the Title sdt's close must come after its open, before the Company sdt opens
    const titleOpen = xml.indexOf('w:val="Title"')
    const companyOpen = xml.indexOf('w:val="Company"')
    expect(titleOpen).toBeGreaterThan(-1)
    expect(companyOpen).toBeGreaterThan(titleOpen)
    const between = xml.slice(titleOpen, companyOpen)
    expect((between.match(/<\/w:sdt>/g) ?? []).length).toBeGreaterThanOrEqual(1)
  })

  it('deleting the whole group leaves no dangling sdt bytes', async () => {
    const { parsed, doc } = await openNested()
    const pruned: PmNode = { ...doc, content: [] }
    const plan = pmDocToSavePlan(pruned, parsed.blocks)
    const xml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(sdtOpenCount(xml)).toBe(0)
    expect(sdtCloseCount(xml)).toBe(0)
  })
})
