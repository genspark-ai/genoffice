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

  it('renders tables inside the textbox as one line per row, recursing into sdt-wrapped nested tables', async () => {
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
    const lines = box?.paras.map((p) => p.runs.map((r) => r.text).join('')) ?? []
    expect(lines[0]).toBe('Rating\u2002\u2002Buy')
    // nested table row becomes its own line; multiple cell paragraphs joined with a space
    expect(lines).toContain('Alice alice@example.com')
    // flattened rows don't map 1:1 to w:p children — the box must not be editable,
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

  it('still protects paragraphs whose own runs carry field codes', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>',
      }),
    )
    expect(doc.blocks[0].type).toBe('passthrough')
    expect(doc.blocks[0].label).not.toBe('Text box')
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
