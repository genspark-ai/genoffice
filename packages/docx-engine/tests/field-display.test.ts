import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { generateParagraphXml, generateTocFieldXml, parseDocx, saveDocx } from '../src/index'
import { sdtCheckboxGlyphs } from '../src/checkbox-control'
import type { Block, GenerateContext, GeneratedBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const documentXmlOf = async (bytes: Uint8Array): Promise<string> =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

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

const ZOTERO_CITATION_INSTR =
  'ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"citation-1","citationItems":[{"id":1}]}'
const ZOTERO_CITATION_PARAGRAPH =
  '<w:p><w:r><w:t>Evidence </w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${ZOTERO_CITATION_INSTR} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>(Smith, 2024)</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t>.</w:t></w:r></w:p>'

const ZOTERO_BIBLIOGRAPHY_INSTR =
  'ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY'
const ZOTERO_BIBLIOGRAPHY_PARAGRAPHS =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${ZOTERO_BIBLIOGRAPHY_INSTR} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:rPr><w:i/></w:rPr><w:t>Alpha, A. (2024). First study.</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Beta, B. (2023). Second study.</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Gamma, G. (2022). Third study.</w:t></w:r>' +
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
      leader: 'dot',
    })
  })

  it('TOC entry splits on spaced and paired tab variants (LO/Google converters)', async () => {
    for (const tab of ['<w:tab />', '<w:tab></w:tab>']) {
      const body =
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
        `<w:r><w:t>Chapter One</w:t></w:r><w:r>${tab}</w:r><w:r><w:t>12</w:t></w:r></w:p>`
      const doc = await parseDocx(await buildDocx({ bodyXml: body }))
      expect(doc.blocks[0].fieldDisplay).toMatchObject({
        kind: 'tocLine',
        left: 'Chapter One',
        right: '12',
      })
    }
  })

  it('a TOC entry carries the leader of its page-number tab (direct stop, else the style)', async () => {
    const entry = (pPrTail: string, style = 'TOC1') =>
      `<w:p><w:pPr><w:pStyle w:val="${style}"/>${pPrTail}</w:pPr>` +
      '<w:r><w:t>Title</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:p>'
    const styles =
      '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/>' +
      '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/></w:style>'
    const doc = await parseDocx(
      await buildDocx({
        extraStylesXml: styles,
        bodyXml:
          entry('<w:tabs><w:tab w:val="right" w:leader="hyphen" w:pos="9350"/></w:tabs>') +
          entry('<w:tabs><w:tab w:val="right" w:pos="9350"/></w:tabs>') +
          entry('<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>') +
          entry('', 'TOC2') +
          entry(
            '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:leader="underscore" w:pos="9350"/></w:tabs>',
          ),
      }),
    )
    const leaders = doc.blocks.slice(0, 5).map((b) => b.fieldDisplay?.leader)
    // direct stop wins; a direct right stop without w:leader is a bare tab;
    // no direct right stop falls back to the style; nothing known stays undefined
    expect(leaders).toEqual(['hyphen', 'none', 'dot', undefined, 'underscore'])
  })

  it('a TOC entry reads single-quoted tab stop values and leaders', async () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/>' +
      "<w:tabs><w:tab w:val='right' w:leader='dot' w:pos='9350'/></w:tabs></w:pPr>" +
      '<w:r><w:t>Title</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({ kind: 'tocLine', leader: 'dot' })
  })

  it('a TOC entry carries the leading result run face and weight (Word draws the entry with its runs)', async () => {
    const entry = (rPr: string) =>
      '<w:p><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="right" w:pos="8786"/></w:tabs>' +
      '<w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>' +
      '<w:hyperlink w:anchor="_Toc1">' +
      `<w:r>${rPr}<w:t>Annexe 1-1 : Classification</w:t></w:r>` +
      '<w:r><w:tab/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>6</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '</w:hyperlink></w:p>'
    const styled = entry(
      '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="MS Gothic" w:hAnsi="Times New Roman"/><w:b/></w:rPr>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: styled + entry('') }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'Annexe 1-1 : Classification',
      right: '6',
      fontFamily: 'Times New Roman',
      bold: true,
    })
    // the paragraph mark's sz is not the entry's text size
    expect(doc.blocks[0].fieldDisplay?.szHalfPoints).toBeUndefined()
    expect(doc.blocks[1].fieldDisplay?.fontFamily).toBeUndefined()
    expect(doc.blocks[1].fieldDisplay?.bold).toBeUndefined()
    // a wholly deleted CJK entry still resolves the East Asian face
    const deleted =
      '<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:del w:id="1" w:author="a" w:date="2024-01-01T00:00:00Z">' +
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="SimSun"/></w:rPr><w:delText>\u7b2c\u4e00\u7ae0</w:delText></w:r>' +
      '</w:del></w:p>'
    const del = await parseDocx(await buildDocx({ bodyXml: deleted }))
    expect(del.blocks[0].fieldDisplay).toMatchObject({ deleted: true, fontFamily: 'SimSun' })
  })

  it('TOC bold respects off/none/case variants (onOff parity)', async () => {
    const offEntry = (b: string) =>
      '<w:p><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="right" w:pos="8786"/></w:tabs>' +
      '<w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>' +
      '<w:hyperlink w:anchor="_Toc1">' +
      `<w:r><w:rPr>${b}</w:rPr><w:t>Title</w:t></w:r>` +
      '<w:r><w:tab/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>6</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '</w:hyperlink></w:p>'
    for (const v of ['off', 'none', 'False', 'OFF', '0']) {
      const doc = await parseDocx(await buildDocx({ bodyXml: offEntry(`<w:b w:val="${v}"/>`) }))
      expect(doc.blocks[0].fieldDisplay?.bold).toBeUndefined()
    }
    const on = await parseDocx(await buildDocx({ bodyXml: offEntry('<w:b w:val="true"/>') }))
    expect(on.blocks[0].fieldDisplay?.bold).toBe(true)
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

  it('Zotero citation fields remain editable and round-trip as Word ADDIN fields', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: ZOTERO_CITATION_PARAGRAPH }))
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].runs).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ text: '(Smith, 2024)', instrField: ZOTERO_CITATION_INSTR }),
      ]),
    )

    const block = doc.blocks[0]
    const saved = await saveDocx(doc, [{ kind: 'generated', block: block as GeneratedBlock }])
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].runs).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ text: '(Smith, 2024)', instrField: ZOTERO_CITATION_INSTR }),
      ]),
    )
  })

  it('keeps styled runs inside one single-paragraph Zotero field', async () => {
    const xml = ZOTERO_CITATION_PARAGRAPH.replace(
      '<w:r><w:t>(Smith, 2024)</w:t></w:r>',
      '<w:r><w:t xml:space="preserve">(Smith, </w:t></w:r>' +
        '<w:r><w:rPr><w:i/></w:rPr><w:t>2024)</w:t></w:r>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const fieldRuns = doc.blocks[0].runs?.filter((run) => run.instrField) ?? []
    expect(fieldRuns.map((run) => run.zoteroFieldPart)).toEqual(['begin', 'end'])
    expect(fieldRuns[1].italic).toBe(true)

    const saved = await saveDocx(doc, [
      { kind: 'generated', block: doc.blocks[0] as GeneratedBlock },
    ])
    const reparsed = await parseDocx(saved)
    const reparsedRuns = reparsed.blocks[0].runs?.filter((run) => run.instrField) ?? []
    expect(reparsedRuns.map((run) => run.text).join('')).toBe('(Smith, 2024)')
    expect(new Set(reparsedRuns.map((run) => run.zoteroFieldId)).size).toBe(1)
  })

  it('keeps a Zotero bibliography spanning multiple paragraphs as one editable field', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: ZOTERO_BIBLIOGRAPHY_PARAGRAPHS }))
    const blocks = doc.blocks.filter((block) => !block.hidden)

    expect(blocks).toHaveLength(3)
    expect(blocks.map((block) => block.type)).toEqual(['paragraph', 'paragraph', 'paragraph'])
    expect(blocks.map((block) => block.runs?.map((run) => run.text).join(''))).toEqual([
      'Alpha, A. (2024). First study.',
      'Beta, B. (2023). Second study.',
      'Gamma, G. (2022). Third study.',
    ])
    const fieldRuns = blocks.flatMap((block) => block.runs ?? [])
    expect(new Set(fieldRuns.map((run) => run.zoteroFieldId)).size).toBe(1)
    expect(fieldRuns.map((run) => run.zoteroFieldPart)).toEqual(['begin', 'inside', 'end'])
    expect(fieldRuns.every((run) => run.instrField === ZOTERO_BIBLIOGRAPHY_INSTR)).toBe(true)
    expect(fieldRuns[0].italic).toBe(true)

    const generated = blocks.map((block) => ({
      kind: 'generated' as const,
      block: {
        type: 'paragraph' as const,
        styleId: block.styleId,
        format: block.format,
        rawPPr: block.rawPPr,
        runs: block.runs ?? [],
      },
    }))
    const saved = await saveDocx(doc, generated)
    const reparsed = await parseDocx(saved)
    const reparsedBlocks = reparsed.blocks.filter((block) => !block.hidden)
    expect(reparsedBlocks.map((block) => block.runs?.map((run) => run.text).join(''))).toEqual([
      'Alpha, A. (2024). First study.',
      'Beta, B. (2023). Second study.',
      'Gamma, G. (2022). Third study.',
    ])
    const reparsedRuns = reparsedBlocks.flatMap((block) => block.runs ?? [])
    expect(new Set(reparsedRuns.map((run) => run.zoteroFieldId)).size).toBe(1)
    expect(reparsedRuns.map((run) => run.zoteroFieldPart)).toEqual(['begin', 'inside', 'end'])
  })

  it('closes a bibliography field whose end paragraph was deleted instead of writing an open field', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: ZOTERO_BIBLIOGRAPHY_PARAGRAPHS }))
    const blocks = doc.blocks.filter((block) => !block.hidden)
    // the first two paragraphs are untouched originals; the paragraph carrying the field end is gone
    const saved = await saveDocx(
      doc,
      blocks
        .slice(0, 2)
        .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! })),
    )
    const xml = await documentXmlOf(saved)
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(1)
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(1)
    const reparsed = await parseDocx(saved)
    const reparsedBlocks = reparsed.blocks.filter((block) => !block.hidden)
    expect(reparsedBlocks.map((block) => block.runs?.map((run) => run.text).join(''))).toEqual([
      'Alpha, A. (2024). First study.',
      'Beta, B. (2023). Second study.',
    ])
    expect(reparsedBlocks[0].runs?.map((run) => run.zoteroFieldPart)).toEqual(['single'])
    expect(reparsedBlocks[1].runs?.every((run) => run.instrField === undefined)).toBe(true)
  })

  it('does not merge an inline citation with a cross-paragraph bibliography', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: ZOTERO_CITATION_PARAGRAPH + ZOTERO_BIBLIOGRAPHY_PARAGRAPHS }),
    )
    const blocks = doc.blocks.filter((block) => !block.hidden)
    const citationId = blocks[0].runs?.find((run) => run.instrField)?.zoteroFieldId
    const bibliographyIds = blocks
      .slice(1)
      .flatMap((block) => block.runs ?? [])
      .map((run) => run.zoteroFieldId)

    expect(citationId).toBeDefined()
    expect(new Set(bibliographyIds).size).toBe(1)
    expect(bibliographyIds[0]).not.toBe(citationId)
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

// IF field whose instruction text runs across three paragraphs: Word hides
// the first two paragraph marks and shows one line "Left {It’s not "
const SPLIT_CODE_PARAGRAPHS =
  '<w:p><w:r><w:t xml:space="preserve">Left </w:t></w:r><w:r><w:t>{</w:t></w:r>' +
  '<w:r><w:t xml:space="preserve">It’s </w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve">IF DATE \\@ "M-d" &lt;&gt; "1-4" "not " </w:instrText></w:r></w:p>' +
  '<w:p><w:r><w:instrText xml:space="preserve">second code paragraph </w:instrText></w:r></w:p>' +
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
  '<w:r><w:instrText xml:space="preserve">\\* MERGEFORMAT </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t>not</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t xml:space="preserve"> </w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>After</w:t></w:r></w:p>'

describe('field code spanning paragraphs', () => {
  const shownMarkers = (blocks: Block[]) =>
    blocks.filter((b) => !b.hidden).map((b) => b.invisibleMarker ?? false)

  it('paragraph marks inside field code are hidden; the last mark shows the joined result', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: SPLIT_CODE_PARAGRAPHS }))
    expect(shownMarkers(doc.blocks)).toEqual([true, true, false, false])
    const tail = doc.blocks[2]
    expect(tail.type).toBe('passthrough')
    expect(tail.fieldDisplay?.kind).toBe('text')
    expect(tail.fieldDisplay?.left).toBe('Left {It’s not')
    expect(tail.fieldDisplay?.align).toBe('center')
    expect(tail.fieldDisplay?.runs?.map((r) => [r.text, r.bold ?? false])).toEqual([
      ['Left {It’s ', false],
      ['not', true],
    ])
    expect(doc.blocks[3].type).toBe('paragraph')
    // the blocks still save byte-identically
    for (const b of doc.blocks.slice(0, 3)) expect(b.originalXml).toMatch(/^<w:p>/)
  })

  it('a field code left open at the end of the body folds nothing', async () => {
    const xml =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r></w:p>' +
      '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(shownMarkers(doc.blocks)).toEqual([false, false])
    expect(doc.blocks[1].type).toBe('paragraph')
  })

  it('single-quoted fldChar runs fold like double-quoted ones', async () => {
    const xml = SPLIT_CODE_PARAGRAPHS.replaceAll(
      /w:fldCharType="(begin|separate|end)"/g,
      "w:fldCharType='$1'",
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(shownMarkers(doc.blocks)).toEqual([true, true, false, false])
    const tail = doc.blocks[2]
    expect(tail.fieldDisplay?.kind).toBe('text')
    expect(tail.fieldDisplay?.left).toBe('Left {It’s not')
  })

  it('paragraph marks inside a field result (TOC entries) stay visible', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: TOC_ENTRY_PARAGRAPH + FIELD_END_PAGEBREAK_PARAGRAPH }),
    )
    expect(shownMarkers(doc.blocks)).toEqual([false, false])
    expect(doc.blocks[0].fieldDisplay?.kind).toBe('tocLine')
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

  it('handles more entries than the argument limit and clamps the level range', () => {
    const many = Array.from({ length: 200_000 }, (_, i) => ({
      level: (i % 9) + 1,
      text: `H${i}`,
    }))
    const frags = generateTocFieldXml(many)
    expect(frags).toHaveLength(many.length)
    expect(frags[0]).toContain(' TOC \\o "1-9" \\h \\z \\u ')
    // a level above 9 still clamps the instruction range
    expect(generateTocFieldXml([{ level: 42, text: 'deep' }])[0]).toContain(
      ' TOC \\o "1-9" \\h \\z \\u ',
    )
    expect(
      generateTocFieldXml([
        { level: 1, text: 'a' },
        { level: 4, text: 'b' },
      ])[0],
    ).toContain(' TOC \\o "1-4" \\h \\z \\u ')
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

  it('sizes the glyph by the begin run rPr like a sizeAuto box', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: checkboxParagraph('<w:default w:val="0"/>').replace(
          '<w:r><w:fldChar w:fldCharType="begin">',
          '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="16"/></w:rPr><w:fldChar w:fldCharType="begin">',
        ),
      }),
    )
    expect(doc.blocks[0].runs?.[1]).toMatchObject({
      text: '☐',
      instrField: 'FORMCHECKBOX',
      sizeHalfPoints: 16,
      font: 'Calibri',
    })
    expect(doc.blocks[0].runs?.[1].fldBeginXml).toContain('<w:sz w:val="16"/>')
  })

  it('checked state comes from w:checked (wins over w:default)', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: checkboxParagraph('<w:default w:val="0"/><w:checked/>') }),
    )
    expect(doc.blocks[0].runs?.[1]).toMatchObject({ text: '☒', instrField: 'FORMCHECKBOX' })
  })

  it('reads uppercase checked values (TRUE/ON) like Word does', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: checkboxParagraph('<w:checked w:val="ON"/>') }),
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
      { text: 'L', sizeHalfPoints: 96 },
      { text: 'ight. Living give. Rule grass light.' },
    ])
  })

  it('a uniform explicit size keeps the dominant size on the wrapper', async () => {
    const xml =
      '<w:p>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "/tmp/x.jpeg" \\* MERGEFORMATINET </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>uniform text</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const field = doc.blocks[0].fieldDisplay!
    expect(field.szHalfPoints).toBe(24)
    expect(field.runs).toEqual([{ text: 'uniform text', sizeHalfPoints: 24 }])
  })
})

describe('editable citation fields (ADDIN ZOTERO_ITEM in a justified body paragraph)', () => {
  const CITATION_PARAGRAPH =
    '<w:p><w:pPr><w:jc w:val="both"/></w:pPr>' +
    '<w:r><w:t xml:space="preserve">Published in </w:t></w:r>' +
    '<w:r><w:rPr><w:i/></w:rPr><w:t>European Radiology</w:t></w:r>' +
    '<w:r><w:t xml:space="preserve">, this study </w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"x"} </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:rPr><w:i/></w:rPr><w:t>(21)</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r><w:t xml:space="preserve">. </w:t></w:r></w:p>'

  it('keeps the result runs formatted and the paragraph justified', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: CITATION_PARAGRAPH }))
    const block = doc.blocks[0]
    expect(block.type).toBe('paragraph')
    expect(block.format?.align).toBe('justify')
    expect(block.runs).toEqual([
      expect.objectContaining({ text: 'Published in ' }),
      expect.objectContaining({ text: 'European Radiology', italic: true }),
      expect.objectContaining({ text: ', this study ' }),
      expect.objectContaining({
        text: '(21)',
        italic: true,
        instrField: 'ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"x"}',
      }),
      expect.objectContaining({ text: '. ' }),
    ])
  })

  it('keeps tabs inside the editable field result', async () => {
    const xml = CITATION_PARAGRAPH.replace('<w:t>(21)</w:t>', '<w:tab/><w:t>(21)</w:t>')
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const fieldRun = doc.blocks[0].runs?.find((run) => run.instrField)
    expect(fieldRun).toMatchObject({
      text: '\t(21)',
      italic: true,
      instrField: 'ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"x"}',
      zoteroFieldPart: 'single',
    })
  })
})

// Content-control checkbox as Word writes it: the state lives in w14:checkbox,
// the run in sdtContent repeats the glyph the state selects.
const sdtCheckboxParagraph = (checked: '0' | '1', glyph: string) =>
  '<w:p><w:r><w:t xml:space="preserve">agree </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="42"/>' +
  `<w14:checkbox><w14:checked w14:val="${checked}"/>` +
  '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
  '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>' +
  '<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr>' +
  `<w:t>${glyph}</w:t></w:r></w:sdtContent></w:sdt></w:p>`

describe('w14:checkbox content controls', () => {
  const ctx: GenerateContext = { headingStyleIds: new Map(), allocateHyperlinkRel: () => 'rId1' }

  it('folds into one glyph run that keeps the control properties', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: sdtCheckboxParagraph('1', '\u2612') }))
    expect(doc.blocks[0].runs?.map((r) => r.text)).toEqual(['agree ', '\u2612'])
    expect(doc.blocks[0].runs?.[1]).toMatchObject({ font: 'MS Gothic' })
    expect(doc.blocks[0].runs?.[1].sdtCheckboxXml).toContain('<w14:checkbox>')
  })

  it('draws the glyph from the state, not from the text the file carried', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: sdtCheckboxParagraph('0', '\u2612') }))
    expect(doc.blocks[0].runs?.[1].text).toBe('\u2610')
  })

  it('writes the control back around the glyph with w14:checked following an in-editor toggle', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: sdtCheckboxParagraph('0', '\u2610') }))
    const runs = doc.blocks[0].runs!.map((r) => (r.sdtCheckboxXml ? { ...r, text: '\u2612' } : r))
    const xml = generateParagraphXml({ type: 'paragraph', runs }, ctx)
    expect(xml).toMatch(
      /<w:sdt><w:sdtPr>.*<w14:checked w14:val="1"\/>.*<\/w:sdtPr><w:sdtContent><w:r>.*\u2612.*<\/w:sdtContent><\/w:sdt>/,
    )
    expect(xml).not.toContain('w14:val="0"')
  })

  it('drops the control when the glyph was typed over', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: sdtCheckboxParagraph('1', '☒') }))
    const runs = doc.blocks[0].runs!.map((r) => (r.sdtCheckboxXml ? { ...r, text: 'yes' } : r))
    const xml = generateParagraphXml({ type: 'paragraph', runs }, ctx)
    expect(xml).not.toContain('<w:sdt>')
    expect(xml).toContain('yes')
  })

  it('reads uppercase checked values (TRUE) as checked', async () => {
    const xml = sdtCheckboxParagraph('1', '☐').replace('w14:val="1"', 'w14:val="TRUE"')
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].runs?.[1].text).toBe('☒')
  })

  it('reads single-quoted checkbox glyph values', () => {
    const glyphs = sdtCheckboxGlyphs(
      "<w:sdtPr><w14:checkbox><w14:checkedState w14:val='2611'/><w14:uncheckedState w14:val='2610'/></w14:checkbox></w:sdtPr>",
    )
    expect(glyphs).toEqual({ checked: '☑', unchecked: '☐' })
  })
})

describe('dirty inline fields', () => {
  const DIRTY_DATE_P =
    '<w:p><w:r><w:t xml:space="preserve">Printed </w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> DATE \\@ "yyyy-MM-dd" </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>2026-01-01</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

  it('parses w:dirty on the begin fldChar into Run.fldDirty and writes it back', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: DIRTY_DATE_P }))
    expect(doc.blocks[0].runs?.[1]).toMatchObject({
      text: '2026-01-01',
      instrField: 'DATE \\@ "yyyy-MM-dd"',
      fldDirty: true,
    })
    const out = await saveDocx(doc, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [
            { text: 'Printed ' },
            { text: '1', instrField: 'NUMPAGES', fldDirty: true },
            { text: ' of ', instrField: undefined },
            { text: '3', instrField: 'PAGE' },
          ],
        },
      },
    ])
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:fldChar w:fldCharType="begin" w:dirty="true"/>')
    expect(xml).toMatch(/NUMPAGES[\s\S]*<w:fldChar w:fldCharType="begin"\/>[\s\S]*PAGE/)
    const reparsed = await parseDocx(out)
    const fields = reparsed.blocks[0].runs?.filter((r) => r.instrField) ?? []
    expect(fields.map((r) => [r.instrField, r.fldDirty ?? false])).toEqual([
      ['NUMPAGES', true],
      ['PAGE', false],
    ])
  })
})

// Word's TOC without TOC styles: plain paragraphs holding a bookmark hyperlink,
// a right tab stop with a dot leader and a nested PAGEREF field; the first
// entry opens the TOC field, the last one closes it
const UNSTYLED_TOC_PPR = (extra: string) =>
  '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9746"/></w:tabs>' +
  `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>${extra}</w:pPr>`
const PAGEREF_ENTRY = (n: number, page: string) =>
  `<w:hyperlink w:anchor="_Toc${n}" w:history="1">` +
  `<w:r><w:rPr><w:b/></w:rPr><w:t>Chapter ${n}</w:t></w:r><w:r><w:tab/></w:r>` +
  '<w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc${n} \\h </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:rPr><w:noProof/></w:rPr><w:t>${page}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>'
const UNSTYLED_TOC =
  `<w:p>${UNSTYLED_TOC_PPR('')}<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>' +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${PAGEREF_ENTRY(1, '3')}</w:p>` +
  `<w:p>${UNSTYLED_TOC_PPR('<w:ind w:left="240"/>')}${PAGEREF_ENTRY(2, '5')}</w:p>` +
  `<w:p>${UNSTYLED_TOC_PPR('<w:ind w:left="240"/>')}${PAGEREF_ENTRY(3, '9')}` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

describe('TOC entries without TOC styles (PAGEREF behind a right tab)', () => {
  it('a PAGEREF-only entry is an editable paragraph: bookmark link, tab, inline field', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: UNSTYLED_TOC }))
    const entry = doc.blocks[1]
    expect(entry.type).toBe('paragraph')
    expect(entry.format?.tabStops).toEqual([{ pos: 9746, val: 'right', leader: 'dot' }])
    expect(entry.format?.spaceAfter).toBe(0)
    expect(entry.runs?.map((r) => r.text).join('')).toBe('Chapter 2\t5')
    const page = entry.runs?.find((r) => r.instrField)
    expect(page).toMatchObject({
      text: '5',
      instrField: 'PAGEREF _Toc2 \\h',
      link: { href: '#_Toc2', plain: true },
    })
    expect(entry.runs?.[0]).toMatchObject({ text: 'Chapter 2', bold: true, link: { plain: true } })
  })

  it('the entries opening and closing the TOC field stay protected toc lines with their own geometry', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: UNSTYLED_TOC }))
    const [head, , tail] = doc.blocks
    expect(head.type).toBe('passthrough')
    expect(head.fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'Chapter 1',
      right: '3',
      level: 1,
      leader: 'dot',
      anchor: '_Toc1',
      spaceAfterTwips: 0,
    })
    expect(head.fieldDisplay?.indentLeftTwips).toBeUndefined()
    expect(tail.type).toBe('passthrough')
    expect(tail.fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'Chapter 3',
      right: '9',
      indentLeftTwips: 240,
      spaceAfterTwips: 0,
    })
  })

  it('an edited entry regenerates the PAGEREF field inside its unstyled bookmark link', async () => {
    const source = await buildDocx({ bodyXml: UNSTYLED_TOC })
    const doc = await parseDocx(source)
    const entry = doc.blocks[1]
    const runs = entry.runs!.map((r) =>
      r.text === 'Chapter 2' ? { ...r, text: 'Chapter Two' } : r,
    )
    const saved = await saveDocx(doc, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'generated', block: { type: 'paragraph', runs, format: entry.format } },
      { kind: 'original', docxIndex: 2 },
    ] as never)
    const xml = await documentXmlOf(saved)
    expect(xml).toMatch(
      /<w:hyperlink w:anchor="_Toc2">[\s\S]*Chapter Two[\s\S]*<w:tab\/>[\s\S]*<w:fldChar w:fldCharType="begin"\/>[\s\S]*<w:instrText xml:space="preserve"> PAGEREF _Toc2 \\h <\/w:instrText>[\s\S]*>5<\/w:t>[\s\S]*<w:fldChar w:fldCharType="end"\/><\/w:r><\/w:hyperlink>/,
    )
    expect(xml).not.toContain('w:rStyle w:val="Hyperlink"')
    // the TOC field itself is untouched: one begin, and its end still in the last entry
    expect(xml.match(/ TOC \\o/g)?.length).toBe(1)
    expect(xml.match(/fldCharType="begin"/g)?.length).toBe(4)
    expect(xml.match(/fldCharType="end"/g)?.length).toBe(4)
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[1].runs?.find((r) => r.instrField)?.text).toBe('5')
    expect(reparsed.blocks[2].type).toBe('passthrough')
  })

  it('a fldSimple PAGEREF keeps the enclosing bookmark link and saves inside it', async () => {
    const simple =
      '<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9746"/></w:tabs></w:pPr>' +
      '<w:hyperlink w:anchor="_Toc7"><w:r><w:t>Chapter 7</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:fldSimple w:instr=" PAGEREF _Toc7 \\h "><w:r><w:t>12</w:t></w:r></w:fldSimple></w:hyperlink></w:p>'
    const source = await buildDocx({ bodyXml: simple })
    const doc = await parseDocx(source)
    const entry = doc.blocks[0]
    expect(entry.type).toBe('paragraph')
    expect(entry.runs?.find((r) => r.instrField)).toMatchObject({
      text: '12',
      instrField: 'PAGEREF _Toc7 \\h',
      link: { href: '#_Toc7', plain: true },
    })
    const saved = await saveDocx(doc, [
      { kind: 'generated', block: { type: 'paragraph', runs: entry.runs, format: entry.format } },
    ] as never)
    const xml = await documentXmlOf(saved)
    expect(xml).toMatch(
      /<w:hyperlink w:anchor="_Toc7">[\s\S]*Chapter 7[\s\S]*PAGEREF _Toc7 \\h [\s\S]*>12<\/w:t>[\s\S]*<w:fldChar w:fldCharType="end"\/><\/w:r><\/w:hyperlink>/,
    )
    expect(xml).not.toContain('w:rStyle w:val="Hyperlink"')
  })

  it('a paragraph with a stray field end stays protected even when its own fields fold', async () => {
    const stray =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: stray }))
    expect(doc.blocks[0].type).toBe('passthrough')
  })
})
