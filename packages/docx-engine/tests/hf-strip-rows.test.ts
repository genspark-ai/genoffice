import { describe, expect, it } from 'vitest'
import { PAGE_MARK, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

const HEADER_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
  '</Relationships>'

const LOGO = (docPrId: string): string =>
  '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>781050</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>26670</wp:posOffset></wp:positionV>' +
  `<wp:extent cx="381000" cy="190500"/><wp:wrapNone/><wp:docPr id="${docPrId}" name="Logo"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'

const PAGE_FIELD =
  '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve">PAGE   \\* MERGEFORMAT</w:instrText></w:r>' +
  '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:rPr><w:b/><w:noProof/><w:sz w:val="32"/></w:rPr><w:t>13</w:t></w:r>' +
  '<w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'

async function build(headerInner: string, extraStylesXml = ''): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
    withImage: true,
    extraStylesXml,
    extraRels:
      '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
    extraParts: [
      {
        path: 'word/header1.xml',
        xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}>${headerInner}</w:hdr>`,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      },
      {
        path: 'word/_rels/header1.xml.rels',
        xml: HEADER_RELS,
        contentType: 'application/vnd.openxmlformats-package.relationships+xml',
      },
    ],
  })
}

const LAYOUT_STYLE =
  '<w:style w:type="table" w:styleId="Layout1"><w:name w:val="Layout 1"/>' +
  '<w:pPr><w:spacing w:before="60" w:after="40"/></w:pPr>' +
  '<w:tblPr><w:tblCellMar><w:left w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Tight"><w:name w:val="Tight"/>' +
  '<w:pPr><w:spacing w:before="0"/></w:pPr></w:style>'

const ROW = (cells: string): string =>
  '<w:tbl><w:tblPr><w:tblStyle w:val="Layout1"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="7860"/><w:gridCol w:w="2940"/></w:tblGrid>' +
  `<w:tr>${cells}</w:tr></w:tbl>`

describe('header layout-table rows', () => {
  it('applies the table style w:pPr spacing under the paragraph style and direct pPr', async () => {
    const doc = await parseDocx(
      await build(
        ROW(
          '<w:tc><w:p><w:r><w:t>Title</w:t></w:r></w:p>' +
            '<w:p><w:pPr><w:spacing w:before="200"/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p>' +
            '<w:p><w:pPr><w:pStyle w:val="Tight"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p></w:tc>' +
            '<w:tc><w:p><w:r><w:t>Right</w:t></w:r></w:p></w:tc>',
        ),
        LAYOUT_STYLE,
      ),
    )
    const cell = doc.headerParas!.find((p) => p.cells)!.cells![0]
    expect(cell.paraProps![0]).toMatchObject({ spaceBefore: 60, spaceAfter: 40 })
    expect(cell.paraProps![1]).toMatchObject({ spaceBefore: 200, spaceAfter: 40 })
    expect(cell.paraProps![2]!.spaceBefore).toBeUndefined()
    expect(cell.paraProps![2]!.spaceAfter).toBe(40)
  })

  it('binds a cell-anchored picture to its row (paragraph-relative offsets measure from the row top)', async () => {
    const doc = await parseDocx(
      await build(
        '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>' +
          ROW(
            '<w:tc><w:p><w:r><w:t>Title</w:t></w:r></w:p></w:tc>' +
              `<w:tc><w:p>${LOGO('4')}</w:p></w:tc>`,
          ) +
          `<w:p>${LOGO('5')}</w:p>`,
      ),
    )
    const [inRow, inPara] = doc.headerImages!.filter((i) => i.floating)
    expect(inRow.posVRel).toBe('paragraph')
    expect(inRow.posYPx).toBe(3)
    expect(inRow.anchorPara).toBe(1)
    expect(doc.headerParas![1].cells).toHaveLength(2)
    // a plain-paragraph anchor keeps the strip-top origin
    expect(inPara.anchorPara).toBeUndefined()
    const part = Object.values(doc.hfParts ?? {})[0]
    expect(part.images!.find((i) => i.floating)!.anchorPara).toBe(1)
  })
})

describe('header/footer PAGE field formatting', () => {
  it('formats the page number with the result run rPr, not the field-code run', async () => {
    const doc = await parseDocx(
      await build(`<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${PAGE_FIELD}</w:p>`),
    )
    const run = doc.headerParas![0].runs.find((r) => r.text.includes(PAGE_MARK))!
    expect(run.sizeHalfPoints).toBe(32)
    expect(run.bold).toBe(true)
  })
})
