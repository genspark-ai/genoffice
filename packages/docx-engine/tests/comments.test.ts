import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const COMMENTS_XML =
  XML_DECL +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:comment w:id="1" w:author="Alice" w:initials="A" w:date="2026-07-01T10:00:00Z">' +
  '<w:p><w:r><w:t>第一条批注</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>第二段</w:t></w:r></w:p>' +
  '</w:comment>' +
  '<w:comment w:id="2" w:author="Bob">' +
  '<w:p><w:r><w:t>cross-para note</w:t></w:r></w:p>' +
  '</w:comment>' +
  '</w:comments>'

/** paragraph with an in-paragraph comment range over its middle run */
const COMMENTED_P =
  '<w:p><w:r><w:t>before </w:t></w:r>' +
  '<w:commentRangeStart w:id="1"/>' +
  '<w:r><w:t>marked</w:t></w:r>' +
  '<w:commentRangeEnd w:id="1"/>' +
  '<w:r><w:commentReference w:id="1"/></w:r>' +
  '<w:r><w:t> after</w:t></w:r></w:p>'

/** comment 2 spans two paragraphs: start in one, end in the next */
const CROSS_P_A = '<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>range opens here</w:t></w:r></w:p>'
const CROSS_P_B =
  '<w:p><w:r><w:t>range closes here</w:t></w:r><w:commentRangeEnd w:id="2"/>' +
  '<w:r><w:commentReference w:id="2"/></w:r></w:p>'

async function buildCommentedDocx(): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: COMMENTED_P + CROSS_P_A + CROSS_P_B,
    extraParts: [
      {
        path: 'word/comments.xml',
        xml: COMMENTS_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
      },
    ],
  })
}

describe('comment parsing', () => {
  it('reads word/comments.xml into the display list', async () => {
    const doc = await parseDocx(await buildCommentedDocx())
    expect(doc.comments).toEqual([
      {
        id: '1',
        author: 'Alice',
        initials: 'A',
        date: '2026-07-01T10:00:00Z',
        text: '第一条批注\n第二段',
      },
      { id: '2', author: 'Bob', initials: undefined, date: undefined, text: 'cross-para note' },
    ])
  })

  it('has no comments array entries for documents without comments.xml', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }))
    expect(doc.comments).toEqual([])
  })

  it('keeps commented paragraphs editable and marks covered runs', async () => {
    const doc = await parseDocx(await buildCommentedDocx())
    const p = doc.blocks[0]
    expect(p.type).toBe('paragraph')
    expect(p.runs).toEqual([
      { text: 'before ' },
      { text: 'marked', commentIds: ['1'] },
      { text: ' after' },
    ])
  })

  it('leaves runs of a cross-paragraph range unmarked (markers preserved by passthrough)', async () => {
    const doc = await parseDocx(await buildCommentedDocx())
    expect(doc.blocks[1].type).toBe('paragraph')
    expect(doc.blocks[1].runs).toEqual([{ text: 'range opens here' }])
    expect(doc.blocks[2].runs).toEqual([{ text: 'range closes here' }])
    // Cross-paragraph endpoints go into the block-level passthrough fields
    expect(doc.blocks[1].commentStarts).toEqual(['2'])
    expect(doc.blocks[1].commentEnds).toBeUndefined()
    expect(doc.blocks[2].commentEnds).toEqual(['2'])
    // A range fully inside one paragraph stays out of the cross-paragraph fields
    // (handled by run.commentIds)
    expect(doc.blocks[0].commentStarts).toBeUndefined()
    expect(doc.blocks[0].commentEnds).toBeUndefined()
  })

  it('editing a paragraph inside a cross-paragraph range re-emits its marker (no orphan end)', async () => {
    const bytes = await buildCommentedDocx()
    const doc = await parseDocx(bytes)
    const finalBlocks: SaveBlock[] = doc.blocks
      .filter((b) => !b.hidden)
      .map((b) =>
        b.docxIndex === 1
          ? {
              kind: 'generated',
              block: {
                type: 'paragraph',
                commentStarts: b.commentStarts,
                commentEnds: b.commentEnds,
                runs: [{ text: 'range opens here (edited)' }],
              },
            }
          : { kind: 'original', docxIndex: b.docxIndex! },
      )
    const saved = await saveDocx(doc, finalBlocks)
    const docXml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    // Before the fix: the edited paragraph lost its commentRangeStart, leaving an
    // orphaned commentRangeEnd in the later paragraph
    expect(docXml).toContain('<w:commentRangeStart w:id="2"/>')
    expect(docXml).toContain('range opens here (edited)')
    expect(docXml).toContain('<w:commentRangeEnd w:id="2"/>')
    // Editing the closing paragraph → end + reference are written back together
    const doc2 = await parseDocx(saved)
    const finalBlocks2: SaveBlock[] = doc2.blocks
      .filter((b) => !b.hidden)
      .map((b) =>
        b.runs?.some((r) => r.text.includes('range closes here'))
          ? {
              kind: 'generated',
              block: {
                type: 'paragraph',
                commentStarts: b.commentStarts,
                commentEnds: b.commentEnds,
                runs: [{ text: 'closes edited' }],
              },
            }
          : { kind: 'original', docxIndex: b.docxIndex! },
      )
    const saved2 = await saveDocx(doc2, finalBlocks2)
    const docXml2 = await (await JSZip.loadAsync(saved2)).file('word/document.xml')!.async('string')
    expect(docXml2).toContain('closes edited')
    expect(docXml2).toContain(
      '<w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r>',
    )
  })
})

describe('comment save fidelity', () => {
  it('no edits -> byte-identical output including comments.xml', async () => {
    const bytes = await buildCommentedDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden).map((b) => b.docxIndex!)
    const saved = await saveDocx(
      doc,
      visible.map((docxIndex): SaveBlock => ({ kind: 'original', docxIndex })),
    )
    expect(saved).toBe(bytes)
  })

  it('editing a commented paragraph re-emits its range markers and reference', async () => {
    const bytes = await buildCommentedDocx()
    const doc = await parseDocx(bytes)
    const finalBlocks: SaveBlock[] = doc.blocks
      .filter((b) => !b.hidden)
      .map((b) =>
        b.docxIndex === 0
          ? {
              kind: 'generated',
              block: {
                type: 'paragraph',
                runs: [
                  { text: 'BEFORE ' },
                  { text: 'marked edited', commentIds: ['1'] },
                  { text: ' AFTER' },
                ],
              },
            }
          : { kind: 'original', docxIndex: b.docxIndex! },
      )
    const saved = await saveDocx(doc, finalBlocks)
    const zip = await JSZip.loadAsync(saved)
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).toContain(
      '<w:commentRangeStart w:id="1"/><w:r><w:t xml:space="preserve">marked edited</w:t></w:r>' +
        '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>',
    )
    // comments.xml itself is untouched
    expect(await zip.file('word/comments.xml')!.async('string')).toBe(COMMENTS_XML)
  })

  it('editing one comment keeps the other rich comment entry byte-identical (surgical rebuild)', async () => {
    const RICH_COMMENTS =
      XML_DECL +
      '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:comment w:id="1" w:author="Alice" w:initials="A">' +
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>加粗观点</w:t></w:r>' +
      '<w:r><w:t>与普通文字</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:i/><w:highlight w:val="yellow"/></w:rPr><w:t>第二段斜体高亮</w:t></w:r></w:p>' +
      '</w:comment>' +
      '<w:comment w:id="2" w:author="Bob"><w:p><w:r><w:t>待编辑的批注</w:t></w:r></w:p></w:comment>' +
      '</w:comments>'
    const bytes = await buildDocx({
      bodyXml: COMMENTED_P + CROSS_P_A + CROSS_P_B,
      extraParts: [
        {
          path: 'word/comments.xml',
          xml: RICH_COMMENTS,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        },
      ],
      extraRels:
        '<Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
    })
    const doc = await parseDocx(bytes)
    // Only comment 2's text changes; comment 1 (rich formatting) stays untouched
    const edited = doc.comments.map((c) => (c.id === '2' ? { ...c, text: '已改写的批注' } : c))
    const saved = await saveDocx(
      doc,
      doc.blocks
        .filter((b) => !b.hidden)
        .map((b): SaveBlock => ({ kind: 'original', docxIndex: b.docxIndex! })),
      { comments: edited },
    )
    const commentsXml = await (
      await JSZip.loadAsync(saved)
    )
      .file('word/comments.xml')!
      .async('string')
    // Before the surgical fix: comment 1 was rebuilt as plain text, losing its
    // bold/italic/highlight
    expect(commentsXml).toContain(
      '<w:comment w:id="1" w:author="Alice" w:initials="A">' +
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>加粗观点</w:t></w:r>' +
        '<w:r><w:t>与普通文字</w:t></w:r></w:p>' +
        '<w:p><w:r><w:rPr><w:i/><w:highlight w:val="yellow"/></w:rPr><w:t>第二段斜体高亮</w:t></w:r></w:p>' +
        '</w:comment>',
    )
    expect(commentsXml).toContain('已改写的批注')
    expect(commentsXml).not.toContain('待编辑的批注')
  })
})

describe('comment replies/resolution (commentsExtended)', () => {
  it('round-trips new replies and the resolved state', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>被批注文字</w:t></w:r>' +
          '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>',
      }),
    )
    const blocks = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    const comments = [
      { id: '1', author: '甲', text: '主批注' },
      { id: '2', author: '乙', text: '回复一下', parentId: '1' },
    ]
    const saved = await saveDocx(parsed, blocks, { comments })
    const reparsed = await parseDocx(saved)
    expect(reparsed.comments).toHaveLength(2)
    const reply = reparsed.comments.find((c) => c.id === '2')!
    expect(reply.parentId).toBe('1')
    expect(reply.paraId).toBeTruthy()
    // Resolve the main comment → done round-trips
    const resolved = reparsed.comments.map((c) => ({ ...c, done: true }))
    const saved2 = await saveDocx(await parseDocx(saved), blocks, { comments: resolved })
    const reparsed2 = await parseDocx(saved2)
    expect(reparsed2.comments.every((c) => c.done)).toBe(true)
    // Part/relationship/ContentType wired together
    const zip = await (await import('jszip')).default.loadAsync(saved2)
    expect(zip.file('word/commentsExtended.xml')).toBeTruthy()
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('commentsExtended.xml')
    const ct = await zip.file('[Content_Types].xml')!.async('string')
    expect(ct).toContain('commentsExtended+xml')
  })
})

describe('reference-only comments (bare w:commentReference, no range markers)', () => {
  const refDocx = (bodyXml: string) =>
    buildDocx({
      bodyXml,
      extraParts: [
        {
          path: 'word/comments.xml',
          xml: COMMENTS_XML,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        },
      ],
    })

  it('anchors on the preceding run', async () => {
    const doc = await parseDocx(
      await refDocx(
        '<w:p><w:r><w:t>hello</w:t></w:r><w:r><w:commentReference w:id="1"/></w:r></w:p>',
      ),
    )
    expect(doc.blocks[0].runs).toEqual([{ text: 'hello', commentIds: ['1'] }])
  })

  it('anchors a leading reference on the following run', async () => {
    const doc = await parseDocx(
      await refDocx(
        '<w:p><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t>after</w:t></w:r></w:p>',
      ),
    )
    expect(doc.blocks[0].runs).toEqual([{ text: 'after', commentIds: ['1'] }])
  })

  it('keeps an empty paragraph with only a reference parseable (comment stays in the list)', async () => {
    const doc = await parseDocx(
      await refDocx('<w:p><w:r><w:commentReference w:id="1"/></w:r></w:p>'),
    )
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].runs).toEqual([])
    expect(doc.comments.map((c) => c.id)).toContain('1')
  })

  it('does not anchor a reference whose id has range markers elsewhere', async () => {
    const doc = await parseDocx(await refDocx(CROSS_P_A + CROSS_P_B))
    expect(doc.blocks[1].runs).toEqual([{ text: 'range closes here' }])
  })
})
