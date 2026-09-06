import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** header layout tables: borders, fixed widths, row height, cell margins, compat-mode geometry */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const HDR_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'

function settingsPart(compatibilityMode: number) {
  return {
    path: 'word/settings.xml',
    xml:
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}><w:compat>` +
      `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${compatibilityMode}"/>` +
      '</w:compat></w:settings>',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  }
}

const BORDERS =
  '<w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((s) => `<w:${s} w:val="single" w:sz="12" w:space="0" w:color="auto"/>`)
    .join('') +
  '</w:tblBorders>'

function cell(width: number, inner: string, tcPrExtra = ''): string {
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${tcPrExtra}</w:tcPr>${inner}</w:tc>`
}

const PAGE_CELL_PARA =
  '<w:p><w:pPr><w:tabs><w:tab w:val="clear" w:pos="4252"/><w:tab w:val="clear" w:pos="8504"/></w:tabs>' +
  '<w:ind w:left="142"/></w:pPr>' +
  '<w:r><w:t xml:space="preserve">Página 1 de 1</w:t></w:r>' +
  '<w:r><w:tab/><w:t xml:space="preserve"> de  </w:t></w:r></w:p>'

const FIXED_TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="70" w:type="dxa"/>' +
  BORDERS +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar>' +
  '</w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3544"/><w:gridCol w:w="4376"/><w:gridCol w:w="1980"/></w:tblGrid>' +
  '<w:tr><w:trPr><w:trHeight w:val="1428"/></w:trPr>' +
  cell(3544, '<w:p><w:r><w:t>LOGO</w:t></w:r></w:p>', '<w:vAlign w:val="center"/>') +
  cell(
    4376,
    '<w:p><w:pPr><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/></w:pPr><w:r><w:t>TITLE</w:t></w:r></w:p>',
    '<w:vAlign w:val="center"/>',
  ) +
  cell(1980, PAGE_CELL_PARA, '<w:vAlign w:val="center"/>') +
  '</w:tr></w:tbl>'

async function headerParas(tableXml: string, compatibilityMode: number) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    extraParts: [
      {
        path: 'word/header1.xml',
        xml: `<w:hdr ${W}>${tableXml}<w:p/></w:hdr>`,
        contentType: HDR_CT,
      },
      settingsPart(compatibilityMode),
    ],
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId60"/>',
  })
  return (await parseDocx(bytes)).headerParas!
}

describe('header layout table geometry', () => {
  it('keeps fixed column widths, borders, row height, margins and vAlign', async () => {
    const [row] = await headerParas(FIXED_TABLE, 12)
    expect(row.cells).toHaveLength(3)
    expect(row.cells!.map((c) => c.widthTwips)).toEqual([3544, 4376, 1980])
    expect(row.row).toMatchObject({ heightTwips: 1428, heightRule: 'atLeast' })
    for (const c of row.cells!) {
      expect(c.vAlign).toBe('center')
      expect(c.marTwips).toEqual({ left: 70, right: 70 })
    }
    // shared lines paint once: inside verticals on the left cell's right edge
    const single = { style: 'single', szEighths: 12, color: 'auto' }
    expect(row.cells![0].borders).toEqual({
      top: single,
      left: single,
      bottom: single,
      right: single,
    })
    expect(row.cells![1].borders).toEqual({ top: single, bottom: single, right: single })
    expect(row.cells![2].borders).toEqual({ top: single, bottom: single, right: single })
  })

  it('carries cell paragraph props: cleared tab stops, indent, spacing, alignment', async () => {
    const [row] = await headerParas(FIXED_TABLE, 12)
    expect(row.cells![1].paraProps?.[0]).toMatchObject({
      align: 'center',
      spaceBefore: 240,
      spaceAfter: 120,
    })
    const pageProps = row.cells![2].paraProps?.[0]
    expect(pageProps?.indentLeft).toBe(142)
    expect(pageProps?.tabStops).toBeUndefined()
  })

  it('legacy layout pulls tblInd back by the cell margin and clips tab overflow', async () => {
    const [row] = await headerParas(FIXED_TABLE, 12)
    expect(row.row).toMatchObject({ indentTwips: 0, tabOverflow: 'clip' })
  })

  it('Word 2013+ layout keeps tblInd at the border and wraps tab overflow', async () => {
    const [row] = await headerParas(FIXED_TABLE, 15)
    expect(row.row).toMatchObject({ indentTwips: 70, tabOverflow: 'wrap' })
  })

  it('marks a w:bidiVisual table so its cells run right to left', async () => {
    const rtl = FIXED_TABLE.replace(
      '<w:tblLayout w:type="fixed"/>',
      '<w:tblLayout w:type="fixed"/><w:bidiVisual/>',
    )
    const [row] = await headerParas(rtl, 15)
    expect(row.row?.bidiVisual).toBe(true)
    // the first markup cell sits at the visual right: outer right edge, no left line
    const single = { style: 'single', szEighths: 12, color: 'auto' }
    expect(row.cells![0].borders).toEqual({ top: single, bottom: single, right: single })
    expect(row.cells![1].borders).toEqual({ top: single, bottom: single, right: single })
    expect(row.cells![2].borders).toEqual({
      top: single,
      left: single,
      bottom: single,
      right: single,
    })
  })

  it('a pct-sized table stays proportional (no absolute widths)', async () => {
    const pct =
      '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr>' +
      cell(4000, '<w:p><w:r><w:t>a</w:t></w:r></w:p>') +
      cell(4000, '<w:p><w:r><w:t>b</w:t></w:r></w:p>') +
      '</w:tr></w:tbl>'
    const [row] = await headerParas(pct, 15)
    expect(row.cells!.map((c) => c.widthTwips)).toEqual([undefined, undefined])
    expect(row.cells!.map((c) => Math.round(c.widthPct!))).toEqual([50, 50])
    expect(row.cells![0].borders).toBeUndefined()
  })

  it('vertically merged cells keep no line through the merge; tcBorders none wins', async () => {
    const merged =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' +
      BORDERS +
      '</w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr>' +
      cell(2000, '<w:p><w:r><w:t>tall</w:t></w:r></w:p>', '<w:vMerge w:val="restart"/>') +
      cell(
        2000,
        '<w:p><w:r><w:t>r1</w:t></w:r></w:p>',
        '<w:tcBorders><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tcBorders>',
      ) +
      '</w:tr><w:tr>' +
      cell(2000, '<w:p/>', '<w:vMerge/>') +
      cell(2000, '<w:p><w:r><w:t>r2</w:t></w:r></w:p>') +
      '</w:tr></w:tbl>'
    const rows = (await headerParas(merged, 15)).filter((p) => p.cells)
    expect(rows).toHaveLength(2)
    expect(rows[0].cells![0].borders?.bottom).toBeUndefined()
    expect(rows[0].cells![1].borders?.bottom).toBeUndefined()
    expect(rows[1].cells![0].borders).toEqual({
      left: { style: 'single', szEighths: 12, color: 'auto' },
      bottom: { style: 'single', szEighths: 12, color: 'auto' },
      right: { style: 'single', szEighths: 12, color: 'auto' },
    })
  })
})

describe('header table cells with a VML horizontal rule', () => {
  it('models the rule run even when the part has no pictures (no media map)', async () => {
    const hrRun =
      '<w:r><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office" style="width:0;height:1.5pt" ' +
      'o:hr="t" fillcolor="#a0a0a0" stroked="f"/></w:pict></w:r>'
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr>' +
      cell(5000, `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Title</w:t></w:r>${hrRun}</w:p>`) +
      '</w:tr></w:tbl>'
    const [row] = await headerParas(table, 15)
    const runs = row.cells![0].paras[0]
    expect(runs.map((r) => r.text)).toEqual(['Title', ''])
    expect(runs[1].image).toMatchObject({
      dataUrl: '',
      rule: { colorHex: 'A0A0A0', thicknessPx: 2 },
    })
    expect(runs[1].image?.xml).toContain('o:hr="t"')
  })
})
