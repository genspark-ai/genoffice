/**
 * Body paragraphs, runs and table cells read the property resolver. The saved
 * model stays direct-only; `effective` is the chain-resolved display view.
 * Word probe 2026-10-22: toggle properties XOR across the style layers
 * (paragraph style bold + character style bold = regular; two vanish layers =
 * visible), ECMA-376 17.7.3.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const styles = (body: string, normalPPr = '') =>
  `${XML_DECL}<w:styles ${W}><w:docDefaults>` +
  '<w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault>' +
  '</w:docDefaults>' +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>${normalPPr}</w:style>` +
  '<w:style w:type="paragraph" w:styleId="BoldPara"><w:name w:val="BoldPara"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="BoldChar"><w:name w:val="BoldChar"/><w:rPr><w:b/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="HidePara"><w:name w:val="HidePara"/><w:basedOn w:val="Normal"/><w:rPr><w:vanish/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="HideChar"><w:name w:val="HideChar"/><w:rPr><w:vanish/></w:rPr></w:style>' +
  '<w:style w:type="table" w:styleId="HeadBold"><w:name w:val="HeadBold"/><w:rPr><w:b/><w:color w:val="1F3864"/></w:rPr>' +
  '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPr></w:style>' +
  `${body}</w:styles>`

const p = (pPr: string, runs: string) => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`
const r = (rPr: string, text: string) =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`

describe('run resolution through the resolver', () => {
  it('XORs toggle properties across the style layers; the saved fields stay direct', async () => {
    const body =
      p('<w:pStyle w:val="BoldPara"/>', r('<w:rStyle w:val="BoldChar"/>', 'both') + r('', 'para')) +
      p('', r('<w:rStyle w:val="BoldChar"/>', 'char') + r('<w:b/>', 'direct'))
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: styles('') }))
    const [both, para] = doc.blocks[0].runs!
    const [char, direct] = doc.blocks[1].runs!
    expect(both.effective!.bold).toBe(false)
    expect(both.effective!.source.bold).toBe('charStyle')
    expect(both.bold).toBeUndefined()
    expect(para.effective!.bold).toBe(true)
    expect(char.effective!.bold).toBe(true)
    expect(direct.effective!.bold).toBe(true)
    expect(direct.bold).toBe(true)
  })

  it('two vanish layers leave the run visible; one hides it', async () => {
    const body = p(
      '<w:pStyle w:val="HidePara"/>',
      r('<w:rStyle w:val="HideChar"/>', 'visible') +
        r('', 'hidden') +
        r('<w:vanish w:val="0"/>', 'anchor'),
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: styles('') }))
    const [visible, hidden, anchor] = doc.blocks[0].runs!
    expect(visible.vanish).toBeUndefined()
    expect(hidden.vanish).toBe(true)
    expect(anchor.vanish).toBeUndefined()
    expect(anchor.vanishOwn).toBe(false)
  })

  it('a style-chain Cs set reaches an rtl run through effective, not csStyle', async () => {
    const extra =
      '<w:style w:type="paragraph" w:styleId="Rtl"><w:name w:val="Rtl"/><w:rPr><w:rtl/><w:bCs/><w:szCs w:val="30"/></w:rPr></w:style>'
    const body = p('<w:pStyle w:val="Rtl"/>', r('', 'שלום') + r('<w:szCs w:val="40"/>', 'עולם'))
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: styles(extra) }))
    const [plain, sized] = doc.blocks[0].runs!
    expect(plain.cs).toBe(true)
    expect(plain.effective).toMatchObject({ cs: true, bold: true, sizeHalfPoints: 30 })
    expect(plain.sizeHalfPoints).toBeUndefined()
    expect(sized.sizeHalfPoints).toBe(40)
    expect(sized.effective!.sizeHalfPoints).toBe(40)
  })
})

describe('hidden paragraphs', () => {
  it('style vanish hides the whole paragraph, a direct mark vanish only an empty one', async () => {
    const body =
      p('<w:pStyle w:val="HidePara"/>', r('', 'gone')) +
      p('<w:rPr><w:vanish/></w:rPr>', '') +
      p('<w:rPr><w:vanish/></w:rPr>', r('', 'kept text')) +
      p('<w:rPr><w:vanish/><w:specVanish/></w:rPr>', '')
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: styles('') }))
    expect(doc.blocks.slice(0, 4).map((b) => b.type)).toEqual([
      'passthrough',
      'passthrough',
      'paragraph',
      'paragraph',
    ])
    expect(doc.blocks[0].label).toBe('Hidden paragraph')
    expect(doc.blocks[2].format?.effective?.hidden).toBe(true)
  })

  it('a vanish-carrying default style hides style-less paragraphs', async () => {
    const hidden = styles('').replace(
      '<w:name w:val="Normal"/>',
      '<w:name w:val="Normal"/><w:rPr><w:vanish/></w:rPr>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: p('', r('', 'x')), stylesXml: hidden }))
    expect(doc.blocks[0].type).toBe('passthrough')
  })
})

describe('paragraph format from the resolver', () => {
  it('attaches the resolved view and keeps the direct-only keys as before', async () => {
    const extra =
      '<w:style w:type="paragraph" w:styleId="Tabs"><w:name w:val="Tabs"/>' +
      '<w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="9000"/></w:tabs><w:autoSpaceDN w:val="0"/><w:spacing w:before="120"/></w:pPr></w:style>'
    const body = p(
      '<w:pStyle w:val="Tabs"/><w:tabs><w:tab w:val="clear" w:pos="720"/><w:tab w:val="center" w:pos="4000"/></w:tabs><w:ind w:left="300"/>',
      r('', 'x'),
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: styles(extra) }))
    const f = doc.blocks[0].format!
    expect(f.indentLeft).toBe(300)
    expect(f.spaceBefore).toBeUndefined()
    expect(f.autoSpace).toBe('latin')
    expect(f.tabStops).toEqual([
      { val: 'clear', pos: 720 },
      { val: 'center', pos: 4000 },
      { val: 'right', pos: 9000, inherited: true },
    ])
    expect(f.effective!.spacing.before).toBe(120)
    expect(f.effective!.source.spaceBefore).toBe('paraStyle')
    expect(f.effective!.source.indentLeft).toBe('direct')
  })

  it('a run-less line without a sized mark still takes the dropped empty run size', async () => {
    const body =
      p('', '<w:r><w:rPr><w:sz w:val="2"/></w:rPr></w:r>') +
      p('<w:rPr><w:sz w:val="16"/></w:rPr>', '<w:r><w:rPr><w:sz w:val="2"/></w:rPr></w:r>') +
      p('', '<w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>')
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: styles('') }))
    expect(doc.blocks[0].format?.emptyRunSizeHalfPoints).toBe(2)
    expect(doc.blocks[1].format?.emptyRunSizeHalfPoints).toBe(16)
    expect(doc.blocks[2].format?.emptyRunSizeHalfPoints).toBeUndefined()
  })

  it('leftChars on a list item sits the body lines the level hanging further in (Word cached w:left)', async () => {
    const numbering =
      `${XML_DECL}<w:numbering ${W}>` +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
      '<w:pPr><w:ind w:left="480" w:hanging="480"/></w:pPr></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    const body = p(
      '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:leftChars="292" w:left="1093"/>',
      r('', 'item'),
    )
    const doc = await parseDocx(
      await buildDocx({ bodyXml: body, stylesXml: styles(''), numberingXml: numbering }),
    )
    const f = doc.blocks[0].format!
    // 2.92 chars of the 10.5pt Normal (613 twips) + the level's 480 hanging
    expect(f.indentLeft).toBe(1093)
    expect(f.indentFirstLine).toBe(-480)
    expect(f.charIndents).toEqual({ left: 292 })
  })

  it('an explicit firstLine="0" next to leftChars survives as a cancel of the style indent', async () => {
    const indented =
      '<w:style w:type="paragraph" w:styleId="Indented"><w:name w:val="Indented"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:ind w:firstLine="420"/></w:pPr></w:style>'
    const body =
      p('<w:pStyle w:val="Indented"/><w:ind w:leftChars="200" w:firstLine="0"/>', r('', 'cancel')) +
      p('<w:pStyle w:val="Indented"/><w:ind w:leftChars="200"/>', r('', 'inherit'))
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: styles(indented) }))
    expect(doc.blocks[0].format!.indentFirstLine).toBe(0)
    expect(doc.blocks[0].format!.effective!.source.indentFirstLine).toBe('direct')
    expect(doc.blocks[1].format!.indentFirstLine).toBe(420)
  })
})

describe('table cells', () => {
  it('cell paragraphs and runs resolve under the table style', async () => {
    const body =
      '<w:tbl><w:tblPr><w:tblStyle w:val="HeadBold"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr><w:tc>${p('', r('', 'cell'))}</w:tc></w:tr></w:tbl>` +
      p('', r('', 'after'))
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: styles('') }))
    const cell = doc.blocks[0].table!.rows[0][0]
    const run = cell.richParas![0].runs[0]
    expect(run.effective).toMatchObject({ bold: true, color: '1F3864' })
    expect(run.effective!.source.bold).toBe('tableStyle')
    expect(run.bold).toBeUndefined()
    expect(doc.blocks[1].runs![0].effective!.bold).toBe(false)
  })
})

describe('conditional table-style layers', () => {
  const style =
    '<w:style w:type="table" w:styleId="Cond"><w:name w:val="Cond"/>' +
    '<w:pPr><w:spacing w:after="120"/></w:pPr>' +
    '<w:tblStylePr w:type="firstRow"><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:b/></w:rPr></w:tblStylePr>' +
    '<w:tblStylePr w:type="lastCol"><w:pPr><w:jc w:val="right"/></w:pPr><w:rPr><w:i/></w:rPr></w:tblStylePr>' +
    '</w:style>'
  const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
  const table = (look: string) =>
    `<w:tbl><w:tblPr><w:tblStyle w:val="Cond"/>${look}</w:tblPr>` +
    `<w:tr>${cell('a')}${cell('b')}</w:tr><w:tr>${cell('c')}${cell('d')}</w:tr></w:tbl>`

  it('a cell paragraph resolves the pPr/rPr of its row and column conditions, rows winning', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: table('<w:tblLook w:firstRow="1" w:lastColumn="1"/>'),
        extraStylesXml: style,
      }),
    )
    const rows = doc.blocks[0].table!.rows
    const eff = (r: number, c: number) => rows[r][c].richParas![0].effective!
    const run = (r: number, c: number) => rows[r][c].richParas![0].runs[0].effective!
    expect(eff(0, 0).align).toBe('center')
    expect(eff(0, 0).source.align).toBe('tableStyle')
    expect(eff(0, 0).spacing.after).toBe(0)
    // firstRow beats lastCol on the shared cell
    expect(eff(0, 1).align).toBe('center')
    expect(eff(1, 1).align).toBe('right')
    expect(eff(1, 0).align).toBeUndefined()
    expect(eff(1, 0).spacing.after).toBe(120)
    expect(run(0, 0).bold).toBe(true)
    expect(run(1, 1).italic).toBe(true)
    expect(run(1, 0).bold).toBe(false)
    // saved fields stay direct-only
    expect(rows[0][0].richParas![0].runs[0].bold).toBeUndefined()
    expect(rows[0][0].richParas![0].align).toBeUndefined()
  })

  it('a legacy hMerge restart whose fold reaches the last column takes lastCol', async () => {
    const merged = (text: string, hMerge: string) =>
      `<w:tc><w:tcPr><w:hMerge${hMerge}/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:tbl><w:tblPr><w:tblStyle w:val="Cond"/><w:tblLook w:firstRow="0" w:lastColumn="1"/></w:tblPr>` +
          `<w:tr>${cell('a')}${cell('b')}${cell('c')}</w:tr>` +
          `<w:tr>${cell('d')}${merged('e', ' w:val="restart"')}${merged('', '')}</w:tr></w:tbl>`,
        extraStylesXml: style,
      }),
    )
    const rows = doc.blocks[0].table!.rows
    expect(rows[1]).toHaveLength(2)
    expect(rows[1][1].colSpan).toBe(2)
    expect(rows[1][1].richParas![0].effective!.align).toBe('right')
    expect(rows[1][1].richParas![0].runs[0].effective!.italic).toBe(true)
    expect(rows[1][0].richParas![0].effective!.align).toBeUndefined()
  })

  it('tblLook switches the conditions off', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: table('<w:tblLook w:firstRow="0" w:lastColumn="0"/>'),
        extraStylesXml: style,
      }),
    )
    const eff = doc.blocks[0].table!.rows[0][1].richParas![0].effective!
    expect(eff.align).toBeUndefined()
    expect(eff.spacing.after).toBe(120)
  })
})
