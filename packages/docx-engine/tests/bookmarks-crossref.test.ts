import { describe, expect, it } from 'vitest'
import { generateParagraphXml, parseDocx, saveDocx, type GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

const BOOKMARKED_PARA =
  '<w:p><w:bookmarkStart w:id="3" w:name="市场规模"/><w:bookmarkEnd w:id="3"/>' +
  '<w:bookmarkStart w:id="4" w:name="_GoBack"/><w:bookmarkEnd w:id="4"/>' +
  '<w:r><w:t>2025 年市场规模达到 1200 亿。</w:t></w:r></w:p>'

const REF_PARA =
  '<w:p><w:r><w:t xml:space="preserve">详见</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> REF 市场规模 \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>2025 年市场规模达到 1200 亿。</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t xml:space="preserve">一节。</w:t></w:r></w:p>'

describe('bookmarks', () => {
  it('parses user bookmarks and splits Word internals into hiddenBookmarks', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BOOKMARKED_PARA }))
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].bookmarks).toEqual(['市场规模'])
    expect(doc.blocks[0].hiddenBookmarks).toEqual(['_GoBack'])
  })

  it('re-emits hidden (_Ref/_Toc) bookmarks on regeneration so REF anchors survive editing', async () => {
    const body =
      '<w:p><w:bookmarkStart w:id="7" w:name="_Ref12345"/><w:bookmarkEnd w:id="7"/>' +
      '<w:r><w:t>被引用的段落。</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const block = doc.blocks[0]
    expect(block.hiddenBookmarks).toEqual(['_Ref12345'])
    // Rebuild after editing the paragraph text → the _Ref bookmark survives (lost here
    // before the broken-link fix)
    const saved = await saveDocx(doc, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          hiddenBookmarks: block.hiddenBookmarks,
          runs: [{ text: '改过的段落。' }],
        },
      },
    ])
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].hiddenBookmarks).toEqual(['_Ref12345'])
  })

  it('re-emits bookmarkStart/End pairs on regeneration', () => {
    const xml = generateParagraphXml(
      { type: 'paragraph', bookmarks: ['目标位置'], runs: [{ text: '正文' }] },
      GEN_CTX,
    )
    const id = /<w:bookmarkStart w:id="(\d+)" w:name="目标位置"\/>/.exec(xml)?.[1]
    expect(id).toBeTruthy()
    expect(xml).toContain(`<w:bookmarkEnd w:id="${id}"/>`)
    expect(xml.indexOf('<w:bookmarkStart')).toBeLessThan(xml.indexOf('<w:r>'))
  })

  it('round-trips bookmarks through generate -> parse', async () => {
    const para = generateParagraphXml(
      { type: 'paragraph', bookmarks: ['章节A'], runs: [{ text: 'x' }] },
      GEN_CTX,
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    expect(doc.blocks[0].bookmarks).toEqual(['章节A'])
  })
})

describe('cross-references (REF fields)', () => {
  it('keeps REF paragraphs editable and folds the field into Run.refField', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: REF_PARA }))
    const block = doc.blocks[0]
    expect(block.type).toBe('paragraph')
    const runs = block.runs!
    expect(runs.map((r) => r.text)).toEqual(['详见', '2025 年市场规模达到 1200 亿。', '一节。'])
    expect(runs[1].refField).toBe('市场规模')
    expect(runs[0].refField).toBeUndefined()
  })

  it('regenerates the full REF field with cached display text', () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [{ text: '见' }, { text: '第一章', refField: 'chap1' }, { text: '。' }],
      },
      GEN_CTX,
    )
    expect(xml).toContain('<w:instrText xml:space="preserve"> REF chap1 \\h </w:instrText>')
    expect(xml).toContain('<w:fldChar w:fldCharType="separate"/>')
    expect(xml).toContain('<w:t xml:space="preserve">第一章</w:t>')
    expect(xml.indexOf('fldCharType="begin"')).toBeLessThan(xml.indexOf('第一章'))
    expect(xml.indexOf('第一章')).toBeLessThan(xml.indexOf('fldCharType="end"'))
  })

  it('round-trips a REF run through generate -> parse', async () => {
    const para = generateParagraphXml(
      { type: 'paragraph', runs: [{ text: '目标内容', refField: 'target' }] },
      GEN_CTX,
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    expect(doc.blocks[0].runs![0]).toMatchObject({ text: '目标内容', refField: 'target' })
  })

  it('keeps untouched bookmark/REF paragraphs byte-identical on save', async () => {
    const source = await buildDocx({ bodyXml: BOOKMARKED_PARA + REF_PARA })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'original', docxIndex: 1 },
    ])
    expect(saved).toEqual(source)
  })
})

describe('REF field switch preservation', () => {
  const SWITCHED_REF_PARA =
    '<w:p><w:r><w:t xml:space="preserve">见第</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> REF _Ref12345 \\r \\h </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>3</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r><w:t xml:space="preserve">章。</w:t></w:r></w:p>'

  it('keeps the original instruction (with \\r) after editing the paragraph', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: SWITCHED_REF_PARA }))
    const runs = doc.blocks[0].runs!
    expect(runs[1]).toMatchObject({ refField: '_Ref12345', refInstr: ' REF _Ref12345 \\r \\h ' })
    const saved = await saveDocx(doc, [
      { kind: 'generated', block: { type: 'paragraph', runs: [{ text: '改' }, ...runs.slice(1)] } },
    ])
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).toContain(
      '<w:instrText xml:space="preserve"> REF _Ref12345 \\r \\h </w:instrText>',
    )
  })

  it('falls back to the default REF instruction for newly created references', () => {
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ text: '第一章', refField: 'chap1' }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:instrText xml:space="preserve"> REF chap1 \\h </w:instrText>')
  })
})
