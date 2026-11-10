import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

/**
 * Legacy VML textboxes and content-control (w:sdt) wrappers, as produced by
 * broker research-report templates: every field lives in an sdt,
 * sidebars are v:shape textboxes full of tables, and whole tables are wrapped
 * in content controls. Display must look through all of it.
 */

const V_NS = 'xmlns:v="urn:schemas-microsoft-com:vml"'

function vmlTextboxParagraph(txbxContent: string, shapeAttrs = ''): string {
  return (
    '<w:p><w:r><w:pict>' +
    `<v:shape ${V_NS} id="s1" type="#_x0000_t202" ` +
    `style="position:absolute;margin-left:380.75pt;margin-top:103.2pt;width:207pt;height:54pt" ${shapeAttrs}>` +
    `<v:textbox><w:txbxContent>${txbxContent}</w:txbxContent></v:textbox>` +
    '</v:shape></w:pict></w:r></w:p>'
  )
}

describe('VML textbox display extraction', () => {
  it('extracts a v:shape textbox as a Text box block with pt→px geometry', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: vmlTextboxParagraph('<w:p><w:r><w:t>Sidebar title</w:t></w:r></w:p>'),
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Text box')
    const box = block.textboxes?.[0]
    expect(box?.widthPx).toBe(276) // 207pt
    expect(box?.heightPx).toBe(72) // 54pt
    expect(box?.minHeightPx).toBe(72)
    expect(box?.paras[0].runs[0].text).toBe('Sidebar title')
  })

  it('reads fillcolor / strokecolor, honoring filled="f" / stroked="f"', async () => {
    const parse = async (attrs: string) => {
      const doc = await parseDocx(
        await buildDocx({
          bodyXml: vmlTextboxParagraph('<w:p><w:r><w:t>x</w:t></w:r></w:p>', attrs),
        }),
      )
      return doc.blocks[0].textboxes?.[0]
    }
    const painted = await parse('fillcolor="#dbe5f1" strokecolor="#4472c4"')
    expect(painted?.fill).toBe('dbe5f1')
    expect(painted?.borderColor).toBe('4472c4')
    const bare = await parse('fillcolor="#dbe5f1" strokecolor="#4472c4" filled="f" stroked="f"')
    expect(bare?.fill).toBeUndefined()
    expect(bare?.borderColor).toBeUndefined()
  })

  it('renders tables inside the textbox as cell rows, nested sdt-wrapped tables flattened into their host cell', async () => {
    const nested =
      '<w:sdt><w:sdtPr><w:alias w:val="analyst"/></w:sdtPr><w:sdtContent>' +
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Alice</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>alice@example.com</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
      '</w:sdtContent></w:sdt>'
    const tbl =
      '<w:tbl>' +
      '<w:tr><w:tc><w:p><w:r><w:t>Rating</w:t></w:r></w:p></w:tc>' +
      `<w:tc><w:p><w:r><w:t>Buy</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc>${nested}<w:p/></w:tc></w:tr>` +
      '</w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: vmlTextboxParagraph(tbl) }))
    const box = doc.blocks[0].textboxes?.[0]
    const rows = box?.paras.map((p) =>
      p.cells!.map((c) => c.paras.map((rs) => rs.map((r) => r.text).join(''))),
    )
    expect(rows).toEqual([[['Rating'], ['Buy']], [['Alice', 'alice@example.com']]])
    // rows don't map 1:1 to w:p children — the box must not be editable,
    // or a sub-editor commit through patchTxbxContent would drop the table
    expect(box?.readOnly).toBe(true)
  })

  it('keeps a numbered anchor paragraph in its list: strayRuns + strayList, no stray box', async () => {
    const numberingXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="1"><w:start w:val="1"/>' +
      '<w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/>' +
      '<w:pPr><w:ind w:left="564" w:hanging="328"/></w:pPr></w:lvl></w:abstractNum>' +
      '<w:num w:numId="20"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    const host = vmlTextboxParagraph('<w:p><w:r><w:t>$918,600</w:t></w:r></w:p>').replace(
      '<w:p><w:r>',
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="20"/></w:numPr>' +
        '<w:ind w:left="537" w:hanging="301"/></w:pPr><w:r>',
    )
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: host.replace(/<\/w:p>$/, '<w:r><w:t>Total Amount Requested:</w:t></w:r></w:p>'),
        numberingXml,
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Text box')
    expect(block.strayRuns?.map((r) => r.text).join('')).toBe('Total Amount Requested:')
    expect(block.strayList).toEqual({ numId: '20', ilvl: 1 })
    expect(block.strayIndent).toEqual({ leftTwips: 537, firstLineTwips: -301 })
    expect(block.textboxes).toHaveLength(1)
    expect(block.textboxes?.[0].paras[0].runs[0].text).toBe('$918,600')
    // the anchor line stays in the preview text like the DrawingML path
    expect(block.previewText).toBe('Total Amount Requested:\n$918,600')
  })

  it('keeps plain-paragraph textboxes editable (no readOnly flag)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: vmlTextboxParagraph('<w:p><w:r><w:t>Editable</w:t></w:r></w:p>'),
      }),
    )
    expect(doc.blocks[0].textboxes?.[0].readOnly).toBeUndefined()
  })

  it('marks boxes whose content is sdt-wrapped as read-only (shells would be lost on rewrite)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: vmlTextboxParagraph(
          '<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>Wrapped</w:t></w:r></w:p></w:sdtContent></w:sdt>',
        ),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.paras[0].runs[0].text).toBe('Wrapped')
    expect(box?.readOnly).toBe(true)
  })

  it('keeps textboxes whose content holds field codes on the Text box path (not Field passthrough)', async () => {
    const fieldPara =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>2022-04-29</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: vmlTextboxParagraph(fieldPara) }))
    const block = doc.blocks[0]
    expect(block.label).toBe('Text box')
    expect(block.textboxes?.[0].paras[0].runs.map((r) => r.text).join('')).toContain('2022-04-29')
  })

  it('a body paragraph whose own runs carry a field is not a Text box', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>',
      }),
    )
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].runs).toEqual([{ text: '1', instrField: 'PAGE' }])
  })
})

describe('sdt-wrapped tables and rows', () => {
  it('a body-level sdt whose content is a table displays as a table block', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:sdt><w:sdtPr><w:alias w:val="report"/></w:sdtPr><w:sdtContent>' +
          '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc></w:tr>' +
          '<w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
          '</w:sdtContent></w:sdt>',
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('table')
    expect(block.table?.rows).toHaveLength(2)
    expect(block.table?.rows[0][0].paras[0]).toBe('H1')
    expect(block.table?.rows[1][0].paras[0]).toBe('Body')
  })

  it('finds rows, cells and cell paragraphs through sdt wrappers inside a table', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl>' +
          '<w:sdt><w:sdtPr/><w:sdtContent>' +
          '<w:tr><w:tc><w:sdt><w:sdtPr/><w:sdtContent>' +
          '<w:p><w:r><w:t>Wrapped</w:t></w:r></w:p>' +
          '</w:sdtContent></w:sdt></w:tc>' +
          '<w:sdt><w:sdtPr/><w:sdtContent><w:tc><w:p><w:r><w:t>Cell2</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>' +
          '</w:tr>' +
          '</w:sdtContent></w:sdt>' +
          '<w:tr><w:tc><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:tc>' +
          '<w:tc><w:p><w:r><w:t>Row</w:t></w:r></w:p></w:tc></w:tr>' +
          '</w:tbl>',
      }),
    )
    const table = doc.blocks[0].table
    expect(table?.rows).toHaveLength(2)
    expect(table?.rows[0].map((c) => c.paras[0])).toEqual(['Wrapped', 'Cell2'])
    expect(table?.rows[1].map((c) => c.paras[0])).toEqual(['Plain', 'Row'])
  })
})

describe('VML textbox frame and Word placement keywords', () => {
  const shapeType = (attrs = '') =>
    `<v:shapetype ${V_NS} id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe" ${attrs}>` +
    '<v:stroke joinstyle="miter"/></v:shapetype>'
  const para = (style: string, shapeAttrs = '', type = shapeType()) =>
    `<w:p><w:r><w:pict>${type}<v:shape ${V_NS} id="s1" type="#_x0000_t202" style="${style}" ${shapeAttrs}>` +
    '<v:textbox><w:txbxContent><w:p><w:r><w:t>Boxed</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
    '</v:shape></w:pict></w:r></w:p>'
  const boxOf = async (bodyXml: string) =>
    (await parseDocx(await buildDocx({ bodyXml }))).blocks[0].textboxes?.[0]

  it('draws the default black frame of a top-level textbox shape', async () => {
    const box = await boxOf(
      para('position:absolute;margin-left:1.5pt;margin-top:-3pt;width:161.25pt;height:89.25pt'),
    )
    expect(box?.borderColor).toBe('000000')
    expect(box?.borderWidthPx).toBeUndefined()
    expect(box?.floating).toBe(true)
    expect(box?.offsetXEmu).toBe(19050)
    expect(box?.offsetYEmu).toBe(-38100)
  })

  it('inherits stroked="f" and stroke colors from the shapetype, honors strokeweight', async () => {
    const off = await boxOf(para('width:100pt;height:50pt', '', shapeType('stroked="f"')))
    expect(off?.borderColor).toBeUndefined()
    const typed = await boxOf(
      para('width:100pt;height:50pt', 'strokeweight="1.5pt"', shapeType('strokecolor="#ff0000"')),
    )
    expect(typed?.borderColor).toBe('ff0000')
    expect(typed?.borderWidthPx).toBe(2)
    const child =
      `<w:p><w:r><w:pict><v:shape ${V_NS} id="s1" style="width:100pt;height:50pt">` +
      '<v:stroke on="f"/><v:textbox><w:txbxContent><w:p><w:r><w:t>x</w:t></w:r></w:p>' +
      '</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>'
    expect((await boxOf(child))?.borderColor).toBeUndefined()
  })

  it('resolves mso-position-horizontal:center against the column', async () => {
    // A4 with 1in margins: 9026 twips of column, box 186.35pt = 3727 twips
    const box = await boxOf(
      para(
        'position:absolute;margin-left:0;margin-top:0;width:186.35pt;height:110.6pt;mso-position-horizontal:center',
      ),
    )
    const colEmu = 9026 * 635
    const wEmu = (box?.widthPx ?? 0) * 9525
    expect(box?.offsetXEmu).toBe(Math.round((colEmu - wEmu) / 2))
    const right = await boxOf(
      para(
        'position:absolute;margin-left:0;margin-top:0;width:100pt;height:50pt;mso-position-horizontal:right;mso-position-horizontal-relative:page',
      ),
    )
    // page-relative right edge, measured from the column origin
    expect(right?.offsetXEmu).toBe(Math.round((11906 - 1440) * 635 - (right?.widthPx ?? 0) * 9525))
  })

  it('page-relative numeric offsets pin the box to the page like DrawingML posOffset', async () => {
    const box = await boxOf(
      para(
        'position:absolute;margin-left:36pt;margin-top:72pt;width:100pt;height:50pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page',
      ),
    )
    expect(box?.offsetXEmu).toBe(36 * 12700 - 1440 * 635)
    expect(box?.offsetYEmu).toBe(72 * 12700 - 1440 * 635)
    expect(box?.pageRelV).toBe(true)
    expect(box?.pageRelVFrom).toBe('page')
    // Word spells a numeric position out as mso-position-*:absolute
    const spelled = await boxOf(
      para(
        'position:absolute;margin-left:36pt;margin-top:72pt;width:100pt;height:50pt;mso-position-horizontal:absolute;mso-position-horizontal-relative:page;mso-position-vertical:absolute;mso-position-vertical-relative:margin',
      ),
    )
    expect(spelled?.offsetXEmu).toBe(36 * 12700 - 1440 * 635)
    expect(spelled?.offsetYEmu).toBe(72 * 12700)
    expect(spelled?.pageRelV).toBe(true)
    expect(spelled?.pageRelVFrom).toBe('margin')
  })
})

describe('VML textbox autofit', () => {
  it('mso-fit-shape-to-text drops the stale declared height', async () => {
    const xml =
      `<w:p><w:r><w:pict><v:shape ${V_NS} id="s1" type="#_x0000_t202" style="position:absolute;margin-left:0;margin-top:0;width:186.35pt;height:110.6pt">` +
      '<v:textbox style="mso-fit-shape-to-text:t"><w:txbxContent><w:p><w:r><w:t>Fit</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
      '</v:shape></w:pict></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.widthPx).toBe(248)
    expect(box?.heightPx).toBeUndefined()
    expect(box?.minHeightPx).toBeUndefined()
  })

  it('textbox table cells keep tblGrid spans and per-cell tcBorders', async () => {
    const cell = (w: number, extra: string, text: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${extra}</w:tcPr>` +
      `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
    const line = (side: string) =>
      `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="00000A"/>`
    const tbl =
      '<w:tbl><w:tblPr><w:tblW w:w="6565" w:type="dxa"/>' +
      `<w:tblBorders>${line('top')}${line('left')}</w:tblBorders></w:tblPr>` +
      '<w:tblGrid><w:gridCol w:w="2455"/><w:gridCol w:w="1897"/><w:gridCol w:w="2213"/></w:tblGrid>' +
      '<w:tr>' +
      cell(2455, `<w:tcBorders>${line('top')}${line('left')}</w:tcBorders>`, 'A') +
      cell(
        4110,
        `<w:gridSpan w:val="2"/><w:tcBorders>${line('top')}${line('right')}</w:tcBorders>`,
        'B',
      ) +
      '</w:tr><w:tr>' +
      cell(2455, `<w:tcBorders>${line('left')}</w:tcBorders>`, 'C') +
      cell(1897, '', 'D') +
      cell(2213, `<w:tcBorders>${line('right')}</w:tcBorders>`, 'E') +
      '</w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: vmlTextboxParagraph(tbl) }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.readOnly).toBe(true)
    const rows = box!.paras.map((p) => p.cells!)
    expect(rows.map((r) => r.map((c) => c.widthTwips))).toEqual([
      [2455, 4110],
      [2455, 1897, 2213],
    ])
    expect(rows[0][0].borders?.top?.style).toBe('single')
    expect(rows[0][0].borders?.left?.style).toBe('single')
    expect(rows[0][1].borders?.top?.style).toBe('single')
    expect(rows[0][1].borders?.right?.style).toBe('single')
    expect(rows[1][0].borders?.left?.style).toBe('single')
    expect(rows[1][1].borders).toBeUndefined()
    expect(rows[1][2].borders?.right?.style).toBe('single')
    expect(box!.paras.every((p) => p.runs.length === 0)).toBe(true)
  })
})
