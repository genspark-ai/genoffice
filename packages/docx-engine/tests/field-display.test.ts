import { describe, expect, it } from 'vitest'
import { generateParagraphXml, generateTocFieldXml, parseDocx } from '../src/index'
import type { GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

// TOC entry paragraph as Word writes it: TOC field begin + hyperlink entry
// with a dot-leader tab and a nested PAGEREF field for the page number.
const TOC_ENTRY_PARAGRAPH =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/>' +
  '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:hyperlink w:anchor="_Toc1" w:history="1">' +
  '<w:r><w:t>第一章 概述</w:t></w:r>' +
  '<w:r><w:tab/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>2</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:hyperlink></w:p>'

const FIELD_END_PAGEBREAK_PARAGRAPH =
  '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:br w:type="page"/></w:r></w:p>'

const PAGE_FIELD_PARAGRAPH =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>- 8 -</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

describe('field paragraph display model', () => {
  it('TOC entry becomes a tocLine with title, page number and level', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TOC_ENTRY_PARAGRAPH }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.fieldDisplay).toEqual({
      kind: 'tocLine',
      left: '第一章 概述',
      right: '2',
      level: 1,
      anchor: '_Toc1',
    })
  })

  it('field-end + page break paragraph shows as a pageBreak marker', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: FIELD_END_PAGEBREAK_PARAGRAPH }))
    expect(doc.blocks[0].fieldDisplay).toEqual({ kind: 'pageBreak' })
  })

  it('PAGE field paragraphs collapse into an editable inline field run (cached result as display text)', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PAGE_FIELD_PARAGRAPH }))
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].runs?.[0]).toMatchObject({ text: '- 8 -', instrField: 'PAGE' })
  })

  it('TOC entry carries its hyperlink anchor for click-to-jump', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TOC_ENTRY_PARAGRAPH }))
    expect(doc.blocks[0].fieldDisplay?.anchor).toBe('_Toc1')
  })
})

describe('generateTocFieldXml', () => {
  const entries = [
    { level: 1, text: '第一章 概述' },
    { level: 2, text: '1.1 背景 <特殊&字符>' },
    { level: 1, text: '第二章 分析' },
  ]

  it('emits a dirty TOC field spanning one paragraph per entry', () => {
    const frags = generateTocFieldXml(entries)
    expect(frags).toHaveLength(3)
    // field structure: dirty begin + instruction in the first, single end in the last
    expect(frags[0]).toContain('w:fldCharType="begin" w:dirty="true"')
    expect(frags[0]).toContain(' TOC \\o "1-2" \\h \\z \\u ')
    expect(frags[0]).toContain('<w:pStyle w:val="TOC1"/>')
    expect(frags[1]).toContain('<w:pStyle w:val="TOC2"/>')
    expect(frags[1]).toContain('&lt;特殊&amp;字符&gt;')
    expect(frags[2]).toContain('w:fldCharType="end"')
    const joined = frags.join('')
    expect(joined.match(/w:fldCharType="end"/g)).toHaveLength(1)
    expect(joined.match(/w:fldCharType="begin"/g)).toHaveLength(1)
  })

  it('round-trips through the parser as tocLine display blocks', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: generateTocFieldXml(entries).join('') }))
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible).toHaveLength(3)
    for (const [i, block] of visible.entries()) {
      expect(block.type).toBe('passthrough')
      expect(block.fieldDisplay?.kind).toBe('tocLine')
      expect(block.fieldDisplay?.left).toBe(entries[i].text)
      expect(block.fieldDisplay?.level).toBe(entries[i].level)
    }
  })
})

describe('HYPERLINK field folding', () => {
  const hyperlinkField = (instr: string) =>
    '<w:r><w:t xml:space="preserve">see </w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText>${instr}</w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>creativets.org</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

  const LIST_HYPERLINK_PARAGRAPH =
    '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>' +
    '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
    hyperlinkField('HYPERLINK "http://creativets.org" \\o "tip"') +
    '</w:p>'

  it('a pure HYPERLINK field folds into an editable link run, keeping list geometry', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: LIST_HYPERLINK_PARAGRAPH, withNumbering: true }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('listItem')
    expect(block.list).toMatchObject({ numId: '1', ilvl: 0 })
    expect(block.runs?.map((r) => r.text).join('')).toBe('see creativets.org')
    const linkRun = block.runs?.find((r) => r.link)
    expect(linkRun).toMatchObject({
      text: 'creativets.org',
      link: { href: 'http://creativets.org', tooltip: 'tip' },
    })
    // cached-result formatting survives the fold
    expect(linkRun?.rawRPr).toContain('<w:u w:val="single"/>')
  })

  it('the folded link regenerates as a w:hyperlink with a fresh rel', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: LIST_HYPERLINK_PARAGRAPH, withNumbering: true }),
    )
    const block = doc.blocks[0]
    const xml = generateParagraphXml(
      { type: 'listItem', list: block.list, runs: block.runs ?? [] },
      {
        headingStyleIds: new Map(),
        allocateHyperlinkRel: (href) => (href === 'http://creativets.org' ? 'rId77' : 'rId0'),
      },
    )
    expect(xml).toContain('<w:hyperlink r:id="rId77" w:tooltip="tip">')
    expect(xml).toContain('creativets.org')
  })

  it('a HYPERLINK with a bookmark switch stays a protected field paragraph', async () => {
    const para = `<w:p>${hyperlinkField('HYPERLINK \\l "bm1"')}</w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    expect(doc.blocks[0].type).toBe('passthrough')
  })
})

// Legacy FORMCHECKBOX form field as Word writes it (POI checkboxes.docx):
// ffData on the begin fldChar defines the box; there is no cached result.
const checkboxParagraph = (state: string) =>
  '<w:p><w:r><w:t xml:space="preserve">item: </w:t></w:r>' +
  `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check1"/><w:enabled/><w:checkBox><w:sizeAuto/>${state}</w:checkBox></w:ffData></w:fldChar></w:r>` +
  '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

describe('FORMCHECKBOX form fields', () => {
  it('unchecked box folds into an editable run with the empty-box glyph', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: checkboxParagraph('<w:default w:val="0"/>') }),
    )
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].runs?.map((r) => r.text)).toEqual(['item: ', '☐'])
  })

  it('checked state comes from w:checked (wins over w:default)', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: checkboxParagraph('<w:default w:val="0"/><w:checked/>') }),
    )
    expect(doc.blocks[0].runs?.[1]).toMatchObject({ text: '☒', instrField: 'FORMCHECKBOX' })
  })

  it('regeneration writes the ffData begin run back verbatim with no cached glyph', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: checkboxParagraph('<w:checked w:val="1"/>') }),
    )
    const run = doc.blocks[0].runs![1]
    const ctx: GenerateContext = {
      headingStyleIds: new Map(),
      allocateHyperlinkRel: () => 'rId1',
    }
    const xml = generateParagraphXml({ type: 'paragraph', runs: doc.blocks[0].runs! }, ctx)
    expect(xml).toContain('<w:ffData>')
    expect(xml).toContain('FORMCHECKBOX')
    expect(xml).not.toContain('☒')
    expect(run.fldBeginXml).toContain('<w:checkBox>')
  })

  it('text typed beside the glyph survives as a plain run after the field', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: checkboxParagraph('<w:checked/>') }))
    const run = { ...doc.blocks[0].runs![1], text: '\u2612yes' }
    const ctx: GenerateContext = {
      headingStyleIds: new Map(),
      allocateHyperlinkRel: () => 'rId1',
    }
    const xml = generateParagraphXml({ type: 'paragraph', runs: [run] }, ctx)
    expect(xml).toContain('<w:ffData>')
    expect(xml).not.toContain('\u2612')
    expect(/<w:fldChar w:fldCharType="end"\/><\/w:r>.*<w:t[^>]*>yes<\/w:t>/s.test(xml)).toBe(true)
  })

  it('editor-merged identical checkboxes regenerate one field per glyph', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: checkboxParagraph('<w:checked/>') }))
    const run = { ...doc.blocks[0].runs![1], text: '\u2612\u2612' }
    const ctx: GenerateContext = {
      headingStyleIds: new Map(),
      allocateHyperlinkRel: () => 'rId1',
    }
    const xml = generateParagraphXml({ type: 'paragraph', runs: [run] }, ctx)
    expect(xml.match(/<w:ffData>/g)).toHaveLength(2)
    expect(xml.match(/FORMCHECKBOX/g)).toHaveLength(2)
    expect(xml).not.toContain('\u2612')
  })

  it('an in-editor glyph flip lands in w:checked on save', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: checkboxParagraph('<w:checked w:val="1"/>') }),
    )
    const run = { ...doc.blocks[0].runs![1], text: '\u2610' }
    const ctx: GenerateContext = {
      headingStyleIds: new Map(),
      allocateHyperlinkRel: () => 'rId1',
    }
    const xml = generateParagraphXml({ type: 'paragraph', runs: [run] }, ctx)
    expect(xml).toContain('<w:checked w:val="0"/>')
    expect(xml).not.toContain('<w:checked w:val="1"/>')
  })

  it('replacing the glyph with text deletes the form field, like Word', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: checkboxParagraph('<w:checked/>') }))
    const run = { ...doc.blocks[0].runs![1], text: 'replaced' }
    const ctx: GenerateContext = {
      headingStyleIds: new Map(),
      allocateHyperlinkRel: () => 'rId1',
    }
    const xml = generateParagraphXml({ type: 'paragraph', runs: [run] }, ctx)
    expect(xml).not.toContain('<w:ffData>')
    expect(xml).not.toContain('FORMCHECKBOX')
    expect(xml).toContain('>replaced<')
  })

  it('FORMCHECKBOX without a w:checkBox definition stays on the passthrough path', async () => {
    const bodyXml =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="X"/></w:ffData></w:fldChar></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].type).toBe('passthrough')
  })
})
