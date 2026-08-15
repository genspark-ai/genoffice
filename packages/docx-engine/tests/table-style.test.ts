import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const GRID_STYLE =
  '<w:style w:type="table" w:styleId="GridBlue"><w:name w:val="Grid Table Blue"/>' +
  '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>' +
  '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>' +
  '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>' +
  '</w:style>' +
  '<w:style w:type="table" w:styleId="GridBlueChild"><w:name w:val="Child"/><w:basedOn w:val="GridBlue"/>' +
  '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr></w:tblStylePr>' +
  '</w:style>'

const row = (a: string, b: string, shd = '') =>
  `<w:tr><w:tc>${shd ? `<w:tcPr>${shd}</w:tcPr>` : ''}<w:p><w:r><w:t>${a}</w:t></w:r></w:p></w:tc>` +
  `<w:tc><w:p><w:r><w:t>${b}</w:t></w:r></w:p></w:tc></w:tr>`

function styledTable(look: string, extra = ''): string {
  return (
    `<w:tbl><w:tblPr><w:tblStyle w:val="GridBlue"/>${look}</w:tblPr>` +
    row('表头A', '表头B') +
    row('一1', '一2', extra) +
    row('二1', '二2') +
    row('三1', '三2') +
    '</w:tbl>'
  )
}

describe('table styles (tblStyle display)', () => {
  it('parses table styles with firstRow / banding and applies them per tblLook', async () => {
    const bytes = await buildDocx({
      bodyXml: styledTable('<w:tblLook w:firstRow="1" w:noHBand="0"/>'),
      extraStylesXml: GRID_STYLE,
    })
    const doc = await parseDocx(bytes)
    const info = doc.styles.get('GridBlue')!
    expect(info.type).toBe('table')
    expect(info.tableDisplay).toMatchObject({
      firstRow: { fill: '4472C4', bold: true, color: 'FFFFFF' },
      band1Fill: 'D9E2F3',
    })

    const rows = doc.blocks[0].table!.rows
    expect(rows[0][0]).toMatchObject({ fill: '4472C4', bold: true, color: 'FFFFFF' })
    // body rows: band1 / none / band1
    expect(rows[1][0].fill).toBe('D9E2F3')
    expect(rows[2][0].fill).toBeUndefined()
    expect(rows[3][0].fill).toBe('D9E2F3')
  })

  it('honors noHBand and firstRow=0 flags', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: styledTable('<w:tblLook w:firstRow="0" w:noHBand="1"/>'),
        extraStylesXml: GRID_STYLE,
      }),
    )
    const rows = doc.blocks[0].table!.rows
    expect(rows[0][0].fill).toBeUndefined()
    expect(rows[0][0].bold).toBeUndefined()
    expect(rows[1][0].fill).toBeUndefined()
  })

  it('explicit cell shading beats the style', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: styledTable(
          '<w:tblLook w:firstRow="1" w:noHBand="0"/>',
          '<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>',
        ),
        extraStylesXml: GRID_STYLE,
      }),
    )
    expect(doc.blocks[0].table!.rows[1][0].fill).toBe('FF0000')
    expect(doc.blocks[0].table!.rows[1][1].fill).toBe('D9E2F3')
  })

  it('resolves basedOn chains between table styles', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraStylesXml: GRID_STYLE,
      }),
    )
    const child = doc.styles.get('GridBlueChild')!
    expect(child.tableDisplay).toMatchObject({
      firstRow: { fill: '4472C4', bold: true },
      band1Fill: 'EEEEEE',
    })
  })

  it('deep-merges paraSpacing through basedOn (partial child keeps parent values)', async () => {
    const spacedStyles =
      '<w:style w:type="table" w:styleId="SpacedParent"><w:name w:val="Spaced"/>' +
      '<w:pPr><w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>' +
      '<w:style w:type="table" w:styleId="SpacedChild"><w:name w:val="Spaced Child"/><w:basedOn w:val="SpacedParent"/>' +
      '<w:pPr><w:spacing w:after="60"/></w:pPr></w:style>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraStylesXml: spacedStyles,
      }),
    )
    expect(doc.styles.get('SpacedChild')!.tableDisplay?.paraSpacing).toMatchObject({
      beforeTwips: 120,
      afterTwips: 60,
      lineRawTwips: 240,
      lineSpacing: 1,
    })
  })

  it('style-derived display keeps untouched tables byte-identical', async () => {
    const source = await buildDocx({
      bodyXml: styledTable('<w:tblLook w:firstRow="1" w:noHBand="0"/>'),
      extraStylesXml: GRID_STYLE,
    })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }])
    expect(saved).toEqual(source)
  })
})
