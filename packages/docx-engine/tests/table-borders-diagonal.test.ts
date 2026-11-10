import { describe, expect, it } from 'vitest'
import { generateTableModelXml, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** w:tl2br / w:tr2bl are the cell diagonals Word calls "Diagonal Down/Up Border" */
describe('cell diagonal borders (w:tl2br / w:tr2bl)', () => {
  it('writes both diagonals in CT_TcBorders order, after the four edges', () => {
    const xml = generateTableModelXml({
      rows: [
        [
          {
            paras: ['a'],
            borders: {
              top: { style: 'single', szEighths: 8, color: 'FF0000' },
              tl2br: { style: 'single', szEighths: 12, color: '0000FF' },
              tr2bl: { style: 'dashed', szEighths: 18, color: '00FF00' },
            },
          },
          { paras: ['b'] },
        ],
      ],
    })
    const block = /<w:tcBorders>[\s\S]*?<\/w:tcBorders>/.exec(xml)![0]
    expect(block).toContain('<w:top w:val="single" w:sz="8" w:space="0" w:color="FF0000"/>')
    expect(block).toContain('<w:tl2br w:val="single" w:sz="12" w:space="0" w:color="0000FF"/>')
    expect(block).toContain('<w:tr2bl w:val="dashed" w:sz="18" w:space="0" w:color="00FF00"/>')
    // the sequence is top, left, bottom, right, tl2br, tr2bl
    expect(block.indexOf('<w:top')).toBeLessThan(block.indexOf('<w:tl2br'))
    expect(block.indexOf('<w:tl2br')).toBeLessThan(block.indexOf('<w:tr2bl'))
  })

  it('omits width and color for a cleared diagonal, like the edges', () => {
    const xml = generateTableModelXml({
      rows: [[{ paras: ['a'], borders: { tl2br: { style: 'none' } } }]],
    })
    const block = /<w:tcBorders>[\s\S]*?<\/w:tcBorders>/.exec(xml)![0]
    expect(block).toContain('<w:tl2br w:val="none"/>')
    expect(block).not.toContain('w:sz=')
  })

  it('leaves a cell without diagonals byte-identical to before', () => {
    const xml = generateTableModelXml({
      rows: [[{ paras: ['a'], borders: { top: { style: 'single' } } }]],
    })
    const block = /<w:tcBorders>[\s\S]*?<\/w:tcBorders>/.exec(xml)![0]
    expect(block).not.toContain('tl2br')
    expect(block).not.toContain('tr2bl')
  })

  it('reads both diagonals back from a saved document', async () => {
    const tbl =
      '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>' +
      '<w:tcBorders>' +
      '<w:top w:val="single" w:sz="8" w:color="FF0000"/>' +
      '<w:tl2br w:val="single" w:sz="12" w:color="0000FF"/>' +
      '<w:tr2bl w:val="dashed" w:sz="18" w:color="00FF00"/>' +
      '</w:tcBorders></w:tcPr>' +
      '<w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: tbl }))
    const model = parsed.blocks.find((b) => b.type === 'table')!.table!
    expect(model.rows[0][0].borders).toEqual({
      top: { style: 'single', szEighths: 8, color: 'FF0000' },
      tl2br: { style: 'single', szEighths: 12, color: '0000FF' },
      tr2bl: { style: 'dashed', szEighths: 18, color: '00FF00' },
    })
  })
})
