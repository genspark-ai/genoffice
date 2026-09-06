import { describe, expect, it } from 'vitest'
import { fieldDisplayOf } from '../src/parse-fields'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const DEL = '<w:del w:id="9" w:author="A" w:date="2024-08-18T21:35:00Z">'

const DELETED_TOC_ENTRY =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:rPr><w:del w:id="8" w:author="A" w:date="2024-08-18T21:35:00Z"/><w:noProof/></w:rPr></w:pPr>' +
  `${DEL}<w:r><w:delText>Introduction</w:delText></w:r><w:r><w:tab/><w:delText>1</w:delText></w:r></w:del></w:p>`

const LIVE_TOC_ENTRY_WITH_DELETED_PAGE =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:rPr><w:ins w:id="1" w:author="A"/></w:rPr></w:pPr>' +
  '<w:hyperlink w:anchor="_Toc1"><w:r><w:t>Basic overview</w:t></w:r><w:r><w:tab/></w:r>' +
  `${DEL}<w:r><w:delText>6</w:delText></w:r></w:del>` +
  '<w:ins w:id="10" w:author="A"><w:r><w:t>7</w:t></w:r></w:ins></w:hyperlink></w:p>'

describe('fieldDisplayOf with tracked deletions', () => {
  it('flags a TOC entry whose every run is deleted, keeping its text for the struck-through view', () => {
    const field = fieldDisplayOf(DELETED_TOC_ENTRY)
    expect(field).toMatchObject({
      kind: 'tocLine',
      left: 'Introduction',
      right: '1',
      deleted: true,
      markDeleted: true,
    })
  })

  it('shows only the live result of an entry whose old page number was deleted', () => {
    const field = fieldDisplayOf(LIVE_TOC_ENTRY_WITH_DELETED_PAGE)
    expect(field).toMatchObject({ kind: 'tocLine', left: 'Basic overview', right: '7' })
    expect(field?.deleted).toBeUndefined()
    expect(field?.markDeleted).toBeUndefined()
  })
})

describe('parseComments display text', () => {
  it('drops field instructions and deleted runs, keeping paragraph breaks', async () => {
    const commentsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:comment w:id="0" w:author="Author" w:initials="SC ED">' +
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE \\# "\'Page: \'#\'" </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:annotationRef/></w:r><w:r><w:t>Remark: first</w:t></w:r>' +
      `${DEL}<w:r><w:delText> gone</w:delText></w:r></w:del></w:p>` +
      '<w:p><w:r><w:t>second</w:t></w:r></w:p>' +
      '</w:comment></w:comments>'
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>body</w:t></w:r><w:commentRangeEnd w:id="0"/>' +
        '<w:r><w:commentReference w:id="0"/></w:r></w:p>',
      extraParts: [
        {
          path: 'word/comments.xml',
          xml: commentsXml,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        },
      ],
    })
    const doc = await parseDocx(bytes)
    expect(doc.comments).toHaveLength(1)
    expect(doc.comments[0]).toMatchObject({
      initials: 'SC ED',
      text: 'Remark: first\nsecond',
    })
  })
})
