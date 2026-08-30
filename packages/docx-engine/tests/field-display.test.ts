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

// paste artifact from Word web copy: an INCLUDEPICTURE field with a dead local
// path sandwiched between styled text runs (public issue #118 demo.docx shape)
const RPR =
  '<w:rPr><w:rFonts w:ascii="\u5b8b\u4f53" w:eastAsia="\u5b8b\u4f53"/><w:sz w:val="24"/></w:rPr>'
const INCLUDEPICTURE_PARAGRAPH =
  `<w:p><w:pPr><w:jc w:val="left"/>${RPR}</w:pPr>` +
  `<w:r>${RPR}<w:t>\u9636\u8d70\u5230\u6cb3</w:t></w:r>` +
  `<w:r>${RPR}<w:t xml:space="preserve">     </w:t></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r>${RPR}<w:instrText xml:space="preserve"> INCLUDEPICTURE "/tmp/x.jpeg" \\* MERGEFORMATINET </w:instrText></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>` +
  `<w:r>${RPR}<w:t>\u5cb8\u8fb9</w:t></w:r></w:p>`

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

  it('a resultless INCLUDEPICTURE text field keeps spaces and run metrics (public issue #118)', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: INCLUDEPICTURE_PARAGRAPH }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.fieldDisplay).toMatchObject({
      kind: 'text',
      left: '\u9636\u8d70\u5230\u6cb3     \u5cb8\u8fb9',
      szHalfPoints: 24,
      fontFamily: '\u5b8b\u4f53',
      align: 'left',
    })
  })

  it('a text field with an explicit line multiple carries the spacing', async () => {
    const xml = INCLUDEPICTURE_PARAGRAPH.replace(
      '<w:pPr><w:jc w:val="left"/>',
      '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="left"/>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({
      kind: 'text',
      lineRule: 'auto',
      lineRawTwips: 360,
      lineSpacing: 1.5,
    })
  })

  it('TOC entry carries its hyperlink anchor for click-to-jump', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TOC_ENTRY_PARAGRAPH }))
    expect(doc.blocks[0].fieldDisplay?.anchor).toBe('_Toc1')
  })

  it('page number follows the LAST tab; a leading outline number becomes the num cell', async () => {
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>1.1.</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>Latar Belakang Masalah</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>7</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({
      kind: 'tocLine',
      num: '1.1.',
      left: 'Latar Belakang Masalah',
      right: '7',
      level: 2,
    })
  })

  it('a long first segment stays part of the title, not the num cell', async () => {
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> TOC </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>BAB I</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>PENDAHULUAN</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>1</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'BAB I PENDAHULUAN',
      right: '1',
    })
    expect(doc.blocks[0].fieldDisplay?.num).toBeUndefined()
  })

  it('entry font size comes from visible result runs, not field-machinery runs', async () => {
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
      '<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:instrText xml:space="preserve"> TOC \\o "1-2" </w:instrText></w:r>' +
      '<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>Chapter One</w:t></w:r>' +
      '<w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].fieldDisplay?.szHalfPoints).toBe(21)
  })

  it('TableofFigures entries render as level-1 toc lines (dot leader + protection)', async () => {
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="TableofFigures"/></w:pPr>' +
      '<w:r><w:t>Tabel 2. 1 Sintaks model pembelajaran</w:t></w:r>' +
      '<w:r><w:tab/></w:r><w:r><w:t>11</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].type).toBe('passthrough')
    expect(doc.blocks[0].fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'Tabel 2. 1 Sintaks model pembelajaran',
      right: '11',
      level: 1,
    })
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

  it('a non-convertible HYPERLINK inside a textbox keeps its cached text visible', async () => {
    // production resumes carry file:///C:\... HYPERLINK fields (backslashes)
    // inside header textboxes; the cached email text must not vanish
    const field = hyperlinkField(
      'HYPERLINK "file:///C:\\Users\\u\\INetCache\\ph.hussam@gmail.com"',
    ).replace('creativets.org', 'ph.hussam@gmail.com')
    const para =
      '<w:p><w:r><w:drawing><wp:anchor behindDoc="0"><wp:extent cx="914400" cy="914400"/>' +
      '<a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>' +
      `<w:p>${field}</w:p>` +
      '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic>' +
      '</wp:anchor></w:drawing></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const box = doc.blocks[0].textboxes?.[0]
    const text = box?.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n')
    expect(text).toContain('ph.hussam@gmail.com')
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

// Mail-merge label/business-card layout: the visible text lives entirely in
// field results inside table cells (complex MERGEFIELD runs and fldSimple).
const MERGE_CELL_PARAGRAPH =
  '<w:p>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> MERGEFIELD Vorname </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:rPr><w:noProof/></w:rPr><w:t>Erika</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t xml:space="preserve"> </w:t></w:r>' +
  '<w:fldSimple w:instr=" MERGEFIELD Nachname ">' +
  '<w:r><w:rPr><w:noProof/></w:rPr><w:t>Mustermann</w:t></w:r>' +
  '</w:fldSimple>' +
  '</w:p>'

const NEXT_FIELD_PARAGRAPH =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> NEXT </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

function cellTable(content: string): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
    content +
    '</w:tc></w:tr></w:tbl>'
  )
}

describe('field cached results in table cells', () => {
  it('complex MERGEFIELD and fldSimple results stay visible as cell runs', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: cellTable(MERGE_CELL_PARAGRAPH) }))
    const cell = doc.blocks[0].table!.rows[0][0]
    const text = (cell.richParas?.[0]?.runs ?? []).map((r) => r.text).join('')
    expect(text).toBe('Erika Mustermann')
    // instruction text must not leak into the visible runs
    expect(text).not.toContain('MERGEFIELD')
  })

  it('a resultless field (NEXT) contributes no text', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: cellTable(NEXT_FIELD_PARAGRAPH) }))
    const cell = doc.blocks[0].table!.rows[0][0]
    const text = (cell.richParas?.[0]?.runs ?? []).map((r) => r.text).join('')
    expect(text).toBe('')
  })
})

describe('mixed-size text fields (manual drop cap)', () => {
  const DROPCAP_FIELD_PARAGRAPH =
    '<w:p>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "/tmp/x.jpeg" \\* MERGEFORMATINET </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r><w:rPr><w:sz w:val="96"/></w:rPr><w:t>L</w:t></w:r>' +
    '<w:r><w:t>ight. Living give. Rule grass light.</w:t></w:r></w:p>'

  it('one oversized letter does not set the whole field size; runs carry per-run sizes', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: DROPCAP_FIELD_PARAGRAPH }))
    const field = doc.blocks[0].fieldDisplay!
    expect(field.kind).toBe('text')
    // dominant size = the inherited default (body text outweighs the cap letter)
    expect(field.szHalfPoints).toBeUndefined()
    expect(field.runs).toEqual([
      { text: 'L', szHalfPoints: 96 },
      { text: 'ight. Living give. Rule grass light.' },
    ])
  })

  it('a uniform explicit size keeps the single-size shape (no runs)', async () => {
    const xml =
      '<w:p>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "/tmp/x.jpeg" \\* MERGEFORMATINET </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>uniform text</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const field = doc.blocks[0].fieldDisplay!
    expect(field.szHalfPoints).toBe(24)
    expect(field.runs).toBeUndefined()
  })
})
