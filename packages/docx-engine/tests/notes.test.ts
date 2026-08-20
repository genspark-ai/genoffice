import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  generateCaptionXml,
  generateIndexFieldXml,
  nextNoteId,
  parseDocx,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const FOOTNOTES_XML =
  XML_DECL +
  '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>这是脚注正文</w:t></w:r></w:p></w:footnote>' +
  '</w:footnotes>'

const FOOTNOTE_P =
  '<w:p><w:r><w:t>正文</w:t></w:r>' +
  '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>' +
  '<w:r><w:t>继续</w:t></w:r></w:p>'

const FOOTNOTES_REL =
  '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'

async function buildFootnoteDocx(): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: FOOTNOTE_P,
    extraRels: FOOTNOTES_REL,
    extraParts: [
      {
        path: 'word/footnotes.xml',
        xml: FOOTNOTES_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
      },
    ],
  })
}

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

describe('footnotes / endnotes', () => {
  it('parses word/footnotes.xml and keeps the reference paragraph editable', async () => {
    const doc = await parseDocx(await buildFootnoteDocx())
    expect(doc.footnotes).toEqual([{ id: '2', text: '这是脚注正文' }])
    expect(doc.endnotes).toEqual([])
    const p = doc.blocks[0]
    expect(p.type).toBe('paragraph')
    expect(p.runs).toEqual([
      { text: '正文' },
      { text: '1', noteRef: { kind: 'footnote', id: '2' } },
      { text: '继续' },
    ])
  })

  it('adding a note keeps the untouched rich footnote entry byte-identical (surgical rebuild)', async () => {
    // Existing footnote with rich formatting (bold run): after adding a new footnote and
    // regenerating the part, the original entry keeps its bytes
    const RICH_FOOTNOTES =
      XML_DECL +
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r>' +
      '<w:r><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr><w:t>加粗红字脚注</w:t></w:r></w:p></w:footnote>' +
      '</w:footnotes>'
    const bytes = await buildDocx({
      bodyXml: FOOTNOTE_P,
      extraRels: FOOTNOTES_REL,
      extraParts: [
        {
          path: 'word/footnotes.xml',
          xml: RICH_FOOTNOTES,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        },
      ],
    })
    const doc = await parseDocx(bytes)
    const out = await saveDocx(doc, originalOrder(doc), {
      footnotes: [...doc.footnotes, { id: '3', text: '新脚注' }],
    })
    const fnXml = await (await JSZip.loadAsync(out)).file('word/footnotes.xml')!.async('string')
    // Before the surgical fix: the original entry was rebuilt as plain text, losing
    // <w:b/> and the color
    expect(fnXml).toContain(
      '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r>' +
        '<w:r><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr><w:t>加粗红字脚注</w:t></w:r></w:p></w:footnote>',
    )
    expect(fnXml).toContain('新脚注')
  })

  it('untouched footnote paragraphs still save byte-identical', async () => {
    const bytes = await buildFootnoteDocx()
    const doc = await parseDocx(bytes)
    const out = await saveDocx(doc, originalOrder(doc))
    expect(out).toBe(doc.internal.originalBytes)
  })

  it('regenerates footnotes.xml preserving separators, and round-trips a new note', async () => {
    const doc = await parseDocx(await buildFootnoteDocx())
    const newId = nextNoteId(doc.footnotes)
    expect(newId).toBe('3')
    const notes = [...doc.footnotes, { id: newId, text: '新脚注' }]
    const out = await saveDocx(
      doc,
      [
        ...originalOrder(doc),
        {
          kind: 'generated',
          block: {
            type: 'paragraph',
            runs: [{ text: '新段落' }, { text: '2', noteRef: { kind: 'footnote', id: newId } }],
          },
        },
      ],
      { footnotes: notes },
    )
    const zip = await JSZip.loadAsync(out)
    const fnXml = await zip.file('word/footnotes.xml')!.async('string')
    expect(fnXml).toContain('w:type="separator"')
    expect(fnXml).toContain('新脚注')

    const reparsed = await parseDocx(out)
    expect(reparsed.footnotes).toEqual([
      { id: '2', text: '这是脚注正文' },
      { id: '3', text: '新脚注' },
    ])
    const last = reparsed.blocks.filter((b) => !b.hidden).at(-1)!
    expect(last.runs).toEqual([
      { text: '新段落' },
      { text: '2', noteRef: { kind: 'footnote', id: '3' } },
    ])
  })

  it('creates word/endnotes.xml with rel + content type when the part is missing', async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>正文</w:t></w:r></w:p>' })
    const doc = await parseDocx(bytes)
    const out = await saveDocx(
      doc,
      [
        {
          kind: 'generated',
          block: {
            type: 'paragraph',
            runs: [{ text: '正文' }, { text: '1', noteRef: { kind: 'endnote', id: '2' } }],
          },
        },
      ],
      { endnotes: [{ id: '2', text: '尾注内容' }] },
    )
    const zip = await JSZip.loadAsync(out)
    const enXml = await zip.file('word/endnotes.xml')!.async('string')
    expect(enXml).toContain('尾注内容')
    expect(enXml).toContain('w:type="separator"')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('endnotes.xml')
    const ct = await zip.file('[Content_Types].xml')!.async('string')
    expect(ct).toContain('/word/endnotes.xml')

    const reparsed = await parseDocx(out)
    expect(reparsed.endnotes).toEqual([{ id: '2', text: '尾注内容' }])
    expect(reparsed.blocks[0].runs).toEqual([
      { text: '正文' },
      { text: '1', noteRef: { kind: 'endnote', id: '2' } },
    ])
  })
})

describe('index entries (XE)', () => {
  const XE_P =
    '<w:p><w:r><w:t>人工智能是研究热点</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> XE "人工智能" </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r><w:t>,继续。</w:t></w:r></w:p>'

  it('keeps XE-only field paragraphs editable and parses the term', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: XE_P }))
    const p = doc.blocks[0]
    expect(p.type).toBe('paragraph')
    expect(p.runs).toEqual([
      { text: '人工智能是研究热点' },
      { text: '', xeTerm: '人工智能' },
      { text: ',继续。' },
    ])
  })

  it('regenerates the XE field when the paragraph is edited', async () => {
    const bytes = await buildDocx({ bodyXml: XE_P })
    const doc = await parseDocx(bytes)
    const out = await saveDocx(doc, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: '改过的文本' }, { text: '', xeTerm: '人工智能' }],
        },
      },
    ])
    const reparsed = await parseDocx(out)
    expect(reparsed.blocks[0].runs).toEqual([
      { text: '改过的文本' },
      { text: '', xeTerm: '人工智能' },
    ])
  })

  it('PAGE field paragraphs collapse into an editable inline field (no longer whole-paragraph protected)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
      }),
    )
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].runs?.[0]?.instrField).toBe('PAGE')
  })

  it('generateIndexFieldXml sorts, dedupes and wraps in one dirty INDEX field', () => {
    const paras = generateIndexFieldXml(['术语B', '术语A', '术语A'])
    expect(paras).toHaveLength(2)
    expect(paras[0]).toContain('fldCharType="begin" w:dirty="true"')
    expect(paras[0]).toContain('INDEX')
    expect(paras[0]).toContain('术语A')
    expect(paras[1]).toContain('术语B')
    expect(paras[1]).toContain('fldCharType="end"')
  })
})

describe('captions', () => {
  it('generateCaptionXml emits a dirty SEQ field with the static number', () => {
    const xml = generateCaptionXml('图', 3, '系统架构')
    expect(xml).toContain(' SEQ 图 \\* ARABIC ')
    expect(xml).toContain('w:dirty="true"')
    expect(xml).toContain('<w:t>3</w:t>')
    expect(xml).toContain('系统架构')
  })

  it('caption paragraphs parse as protected SEQ field blocks with visible text', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: generateCaptionXml('图', 1, '测试') }))
    const p = doc.blocks[0]
    expect(p.type).toBe('passthrough')
    expect(p.fieldDisplay).toEqual({ kind: 'text', left: '图 1 测试' })
  })
})

describe('rich-text footnote display runs', () => {
  it('parses bold/italic/color/size runs; plain-text footnotes carry no richParas', async () => {
    const footnotesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="1"><w:p>' +
      '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>' +
      '<w:r><w:t xml:space="preserve"> 见</w:t></w:r>' +
      '<w:r><w:rPr><w:b/><w:color w:val="c00000"/></w:rPr><w:t>重要文献</w:t></w:r>' +
      '<w:r><w:rPr><w:i/><w:sz w:val="16"/></w:rPr><w:t>(斜体小字)</w:t></w:r>' +
      '</w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>纯文本</w:t></w:r></w:p></w:footnote>' +
      '</w:footnotes>'
    const { parseNotesXml } = await import('../src/notes')
    const notes = parseNotesXml(footnotesXml, 'footnote')
    expect(notes).toHaveLength(2)
    const rich = notes[0]
    expect(rich.richParas?.[0]).toEqual([
      { text: '见' },
      { text: '重要文献', bold: true, color: 'C00000' },
      { text: '(斜体小字)', italic: true, sizeHalfPoints: 16 },
    ])
    expect(notes[1].richParas).toBeUndefined()
    expect(notes[1].text).toBe('纯文本')
  })

  it('serializes richParas runs with size/font formatting for fresh notes (P17)', async () => {
    const { buildNotesXml, parseNotesXml } = await import('../src/notes')
    const xml = buildNotesXml(
      'footnote',
      [
        {
          id: '201',
          text: 'small note tail',
          richParas: [
            [
              { text: 'small note', sizeHalfPoints: 16, fontAscii: 'Arial', bold: true },
              { text: ' tail', sizeHalfPoints: 16, color: '1F4E79' },
            ],
          ],
        },
      ],
      null,
    )
    expect(xml).toContain('<w:sz w:val="16"/>')
    expect(xml).toContain('w:ascii="Arial"')
    expect(xml).toContain('<w:b/>')
    expect(xml).toContain('<w:color w:val="1F4E79"/>')
    // the self-reference mark + spacer shrink to the note's own size
    expect(xml).toContain('<w:vertAlign w:val="superscript"/><w:sz w:val="16"/>')
    // rich rebuilds pin single spacing so template docDefaults cannot inflate the area
    expect(xml).toContain('<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
    const notes = parseNotesXml(xml, 'footnote')
    expect(notes[0]!.text).toBe('small note tail')
    expect(notes[0]!.richParas?.[0]).toEqual([
      { text: 'small note', bold: true, sizeHalfPoints: 16 },
      { text: ' tail', color: '1F4E79', sizeHalfPoints: 16 },
    ])
  })
})
