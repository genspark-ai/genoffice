import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { LegacyDocError, convertLegacyDocToDocx, readLegacyDoc } from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodePapxFkp,
  encodeTableCellPapx,
  encodeTableRowPapx,
  type SyntheticComment,
  type SyntheticPapxEntry,
} from './helpers/synthetic-doc'

async function zipText(docx: Uint8Array, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(docx)
  const file = zip.file(path)
  return file ? file.async('string') : undefined
}

const REFERENCE_RUN =
  '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="ID"/></w:r>'
const referenceRun = (id: number): string => REFERENCE_RUN.replace('ID', String(id))

/** DTTM for 2024-01-09 17:20 */
const DTTM = (124 << 20) | (1 << 16) | (9 << 11) | (17 << 6) | 20

/** "Alpha beta<ref> gamma" with a ranged comment on "beta", then "Second<ref>" with a point comment */
const BODY = 'Alpha beta\u0005 gamma\rSecond\u0005\r'
const COMMENTS: SyntheticComment[] = [
  {
    text: 'Note one',
    author: 'Ann Author',
    initials: 'AA',
    referenceCp: 10,
    range: [6, 10],
    dttm: DTTM,
  },
  { text: ['Point', 'second para'], author: 'Bob', initials: 'B', referenceCp: 24 },
]

describe('comments', () => {
  it('reads author, initials, date, text and range of each comment', () => {
    const { body } = readLegacyDoc(
      buildSyntheticDoc({ pieces: [{ text: BODY }], comments: COMMENTS }),
    )
    expect(body.comments).toHaveLength(2)
    const [ranged, point] = body.comments!
    expect(ranged).toMatchObject({
      id: 0,
      author: 'Ann Author',
      initials: 'AA',
      date: '2024-01-09T17:20:00Z',
      referenceCp: 10,
      rangeStartCp: 6,
      rangeEndCp: 10,
      paragraphs: ['Note one'],
    })
    expect(ranged!.formatted?.map((paragraph) => paragraph.text)).toEqual(['Note one'])
    expect(point).toMatchObject({
      id: 1,
      author: 'Bob',
      initials: 'B',
      referenceCp: 24,
      paragraphs: ['Point', 'second para'],
    })
    expect(point!.date).toBeUndefined()
    expect(point!.rangeStartCp).toBeUndefined()
    expect(body.warnings.filter((warning) => warning.includes('comment'))).toEqual([])
  })

  it('writes comments.xml and the range / reference markers at the anchor positions', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(
      buildSyntheticDoc({ pieces: [{ text: BODY }], comments: COMMENTS }),
    )
    expect(warnings.filter((warning) => warning.includes('comment'))).toEqual([])
    const comments = (await zipText(docx, 'word/comments.xml'))!
    expect(comments).toContain(
      '<w:comment w:id="0" w:author="Ann Author" w:date="2024-01-09T17:20:00Z" w:initials="AA">',
    )
    expect(comments).toContain('<w:comment w:id="1" w:author="Bob" w:initials="B">')
    expect(comments).toContain(
      '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>',
    )
    expect(comments).toContain('>Note one</w:t>')
    expect(comments).toContain('>second para</w:t>')
    expect(await zipText(docx, '[Content_Types].xml')).toContain(
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
    )
    expect(await zipText(docx, 'word/_rels/document.xml.rels')).toMatch(
      /Type="[^"]*\/comments" Target="comments.xml"/,
    )
    const document = (await zipText(docx, 'word/document.xml'))!
    expect(document).toMatch(
      new RegExp(
        'Alpha </w:t></w:r><w:commentRangeStart w:id="0"/><w:r>(?:(?!</w:r>).)*>beta</w:t></w:r>' +
          '<w:commentRangeEnd w:id="0"/>' +
          referenceRun(0) +
          '<w:r>(?:(?!</w:r>).)*> gamma</w:t></w:r>',
      ),
    )
    expect(document).toContain('>Second</w:t></w:r>' + referenceRun(1) + '</w:p>')
    expect(document).not.toContain('<w:commentRangeStart w:id="1"/>')
    const parsed = await parseDocx(docx)
    expect(parsed.comments.map((comment) => [comment.id, comment.author, comment.text])).toEqual([
      ['0', 'Ann Author', 'Note one'],
      ['1', 'Bob', 'Point\nsecond para'],
    ])
  })

  it('anchors a comment inside a table cell', async () => {
    const text = 'Cell\u0005\u0007Other\u0007\u0007After\r'
    const entries: SyntheticPapxEntry[] = [
      {
        fcStart: SYNTHETIC_TEXT_OFFSET,
        fcEnd: SYNTHETIC_TEXT_OFFSET + 12,
        grpprl: encodeTableCellPapx(),
      },
      {
        fcStart: SYNTHETIC_TEXT_OFFSET + 12,
        fcEnd: SYNTHETIC_TEXT_OFFSET + 24,
        grpprl: encodeTableCellPapx(),
      },
      {
        fcStart: SYNTHETIC_TEXT_OFFSET + 24,
        fcEnd: SYNTHETIC_TEXT_OFFSET + 26,
        grpprl: encodeTableRowPapx({ gridWidths: [2000, 2000], cells: [{}, {}] }),
      },
      { fcStart: SYNTHETIC_TEXT_OFFSET + 26, fcEnd: SYNTHETIC_TEXT_OFFSET + text.length * 2 },
    ]
    const { docx } = await convertLegacyDocToDocx(
      buildSyntheticDoc({
        pieces: [{ text }],
        papxFkpPages: [encodePapxFkp(entries)],
        comments: [
          { text: 'In a cell', author: 'Cy', initials: 'C', referenceCp: 4, range: [0, 4] },
        ],
      }),
    )
    const document = (await zipText(docx, 'word/document.xml'))!
    expect(document).toMatch(
      new RegExp(
        '<w:tc>(?:(?!</w:tc>).)*<w:commentRangeStart w:id="0"/><w:r>(?:(?!</w:r>).)*>Cell</w:t></w:r>' +
          '<w:commentRangeEnd w:id="0"/>' +
          referenceRun(0) +
          '</w:p></w:tc>',
      ),
    )
    const parsed = await parseDocx(docx)
    expect(parsed.comments.map((comment) => comment.text)).toEqual(['In a cell'])
  })

  it('degrades to a warning when the comment text table is missing', async () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: BODY }],
      comments: COMMENTS,
      omitPlcfandTxt: true,
    })
    expect(readLegacyDoc(bytes).body.comments).toBeUndefined()
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings).toContain(
      'the document declares comment text but has no comment table; comments are omitted',
    )
    expect(await zipText(docx, 'word/comments.xml')).toBeUndefined()
    expect(await zipText(docx, 'word/document.xml')).not.toContain('commentReference')
  })

  it('skips comments anchored outside the main text with a warning', () => {
    const { body } = readLegacyDoc(
      buildSyntheticDoc({
        pieces: [{ text: BODY }],
        comments: [{ text: 'In a header', author: 'Hy', initials: 'H', referenceCp: 40 }],
      }),
    )
    expect(body.comments).toBeUndefined()
    expect(body.warnings).toContain(
      '1 comment(s) anchored in headers, footers or footnotes are not converted yet; they were omitted',
    )
  })

  it('reports a structurally invalid comment reference table as bad_atn', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: BODY }],
      comments: COMMENTS,
      lcbPlcfandRef: 7,
    })
    let caught: unknown
    try {
      readLegacyDoc(bytes)
    } catch (error) {
      caught = error
    }
    expect(caught).toBeInstanceOf(LegacyDocError)
    expect((caught as LegacyDocError).code).toBe('bad_atn')
  })
})

describe('comments next to notes', () => {
  /** "Ab<fn> cd<ref> ef" with a footnote on cp 2, an endnote on cp 9 and a ranged comment on "cd" */
  const BODY = 'Ab\u0002 cd\u0005 e\u0002f\r'
  const bytes = buildSyntheticDoc({
    pieces: [{ text: BODY }],
    footnotes: [{ cp: 2, text: '\u0002Foot' }],
    endnotes: [{ cp: 9, text: '\u0002End' }],
    comments: [{ text: 'Remark', author: 'Ann', initials: 'A', referenceCp: 6, range: [4, 6] }],
  })

  it('reads the comment story placed between the header and endnote stories', () => {
    const { body } = readLegacyDoc(bytes)
    expect(body.notes!.footnotes.map((note) => note.paragraphs[0]!.text)).toEqual(['Foot'])
    expect(body.notes!.endnotes.map((note) => note.paragraphs[0]!.text)).toEqual(['End'])
    expect(body.comments!.map((comment) => comment.paragraphs)).toEqual([['Remark']])
  })

  it('writes note marks and comment marks without mixing their markers', async () => {
    const { docx } = await convertLegacyDocToDocx(bytes)
    const document = (await zipText(docx, 'word/document.xml'))!
    expect(document).toContain('<w:footnoteReference w:id="1"/>')
    expect(document).toContain('<w:endnoteReference w:id="1"/>')
    expect(document).toContain('<w:commentRangeStart w:id="0"/>')
    expect(document).toContain('<w:commentRangeEnd w:id="0"/>')
    expect(document).toContain(referenceRun(0))
    expect(document).not.toMatch(/[-]/)
    expect(await zipText(docx, 'word/comments.xml')).toContain('Remark')
    expect(await zipText(docx, 'word/footnotes.xml')).toContain('Foot')
    expect(await zipText(docx, 'word/endnotes.xml')).toContain('End')
    expect((await parseDocx(docx)).comments.map((comment) => comment.text)).toEqual(['Remark'])
  })
})
