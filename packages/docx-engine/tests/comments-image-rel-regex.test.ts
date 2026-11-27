import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const COMMENTS_XML =
  XML_DECL +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<w:comment w:id="1" w:author="Alice" w:initials="A">' +
  '<w:p><w:r><w:drawing><a:blip r:embed="rId(1)"/></w:drawing></w:r></w:p>' +
  '</w:comment>' +
  '</w:comments>'

const COMMENT_RELS =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId(1)" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
  'Target="media/image1.png"/>' +
  '</Relationships>'

/** a comment picture whose r:embed carries regex metacharacters */
async function buildCommentedDocx(): Promise<Uint8Array> {
  const doc = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>hello</w:t></w:r></w:p>',
    extraParts: [
      {
        path: 'word/comments.xml',
        xml: COMMENTS_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
      },
    ],
    binaryParts: [
      {
        path: 'word/media/image1.png',
        base64: TINY_PNG_BASE64,
        extension: 'png',
        contentType: 'image/png',
      },
    ],
  })
  const zip = await JSZip.loadAsync(doc)
  zip.file('word/_rels/comments.xml.rels', COMMENT_RELS)
  return zip.generateAsync({ type: 'uint8array' })
}

describe('comment picture relationship lookup', () => {
  it('parses a comment whose relationship id holds regex metacharacters', async () => {
    // unescaped, `new RegExp('…Id="rId(1)"…')` throws and fails the whole open;
    // the document must parse and the comment must survive the lookup
    const doc = await parseDocx(await buildCommentedDocx())
    expect(doc.comments.map((c) => c.id)).toContain('1')
    expect(doc.comments[0]!.author).toBe('Alice')
  })
})
