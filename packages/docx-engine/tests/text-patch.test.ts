import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { patchParagraphTexts } from '../src/text-patch'
import { buildDocx } from './helpers/build-docx'

describe('patchParagraphTexts', () => {
  const COMMENT =
    '<w:comment w:id="1" w:author="甲">' +
    '<w:p><w:r><w:t xml:space="preserve">前缀</w:t></w:r>' +
    '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>红色加粗</w:t></w:r>' +
    '<w:r><w:t xml:space="preserve">后缀</w:t></w:r></w:p>' +
    '<w:p w14:paraId="AA00"><w:hyperlink r:id="rId9"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>链接</w:t></w:r></w:hyperlink></w:p>' +
    '</w:comment>'

  it('editing one paragraph keeps the other paragraph and untouched runs byte-identical', () => {
    const out = patchParagraphTexts(COMMENT, '前缀红色加粗改后缀\n链接')
    expect(out).not.toBeNull()
    // Untouched paragraph keeps its original bytes
    expect(out).toContain('<w:p w14:paraId="AA00"><w:hyperlink r:id="rId9">')
    // The touched run keeps its rPr
    expect(out).toContain('<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>')
    expect(out).toContain('红色加粗改')
  })

  it('inserting text mid-paragraph preserves the runs on both sides', () => {
    const out = patchParagraphTexts(COMMENT, '前缀X红色加粗后缀\n链接')
    expect(out).not.toBeNull()
    expect(out).toContain('前缀X')
    expect(out).toContain('<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>')
  })

  it('cross-run deletion empties the middle run but keeps its formatting shell', () => {
    const out = patchParagraphTexts(COMMENT, '前后缀\n链接')
    expect(out).not.toBeNull()
    const decoded = out!
    expect(decoded).toContain('<w:t xml:space="preserve">前</w:t>')
    // Full re-extracted text matches (guaranteed by the self-check); the red run's
    // content is emptied but its structure remains
    expect(decoded).toContain('<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>')
  })

  it('escapes XML special characters in text correctly', () => {
    const out = patchParagraphTexts(COMMENT, '前缀<A&B>后缀\n链接')
    expect(out).not.toBeNull()
    expect(out).toContain('&lt;A&amp;B&gt;')
  })

  it('returns null when the paragraph count changes (caller falls back to a rebuild)', () => {
    expect(patchParagraphTexts(COMMENT, '只剩一段')).toBeNull()
    expect(patchParagraphTexts(COMMENT, '一\n二\n三')).toBeNull()
  })

  it('treats the footnote first-paragraph self-reference space as an immutable prefix', () => {
    const NOTE =
      '<w:footnote w:id="2">' +
      '<w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>' +
      '<w:r><w:t xml:space="preserve"> </w:t></w:r>' +
      '<w:r><w:rPr><w:i/></w:rPr><w:t>斜体脚注</w:t></w:r></w:p>' +
      '</w:footnote>'
    const out = patchParagraphTexts(NOTE, '斜体脚注改', { stripFirstParaLeadingSpace: true })
    expect(out).not.toBeNull()
    expect(out).toContain('<w:footnoteRef/>')
    expect(out).toContain('<w:t xml:space="preserve"> </w:t>')
    expect(out).toContain('<w:rPr><w:i/></w:rPr>')
    expect(out).toContain('斜体脚注改')
  })

  it('leaves entity-encoded text byte-identical when nothing changed', () => {
    const entry = '<w:comment w:id="1"><w:p><w:r><w:t>A &#8212; B</w:t></w:r></w:p></w:comment>'
    expect(patchParagraphTexts(entry, 'A — B')).toBe(entry)
  })
})

describe('rich-text comment/footnote edits use the surgical patch', () => {
  const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
    doc.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

  it('editing comment text keeps the bold run and hyperlink structure inside the comment', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>正文</w:t></w:r>' +
        '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId31" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
      extraParts: [
        {
          path: 'word/comments.xml',
          xml:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
            '<w:comment w:id="1" w:author="甲"><w:p w14:paraId="11112222">' +
            '<w:r><w:rPr><w:b/></w:rPr><w:t>加粗</w:t></w:r>' +
            '<w:r><w:t xml:space="preserve">批注</w:t></w:r>' +
            '</w:p></w:comment></w:comments>',
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        },
      ],
    })
    const doc = await parseDocx(bytes)
    expect(doc.comments[0].text).toBe('加粗批注')
    const edited = [{ ...doc.comments[0], text: '加粗的批注' }]
    const out = await saveDocx(doc, originalOrder(doc), { comments: edited })
    const xml = await (await JSZip.loadAsync(out)).file('word/comments.xml')!.async('string')
    expect(xml).toContain('<w:rPr><w:b/></w:rPr>')
    expect(xml).toContain('w14:paraId="11112222"')
    const reparsed = await parseDocx(out)
    expect(reparsed.comments[0].text).toBe('加粗的批注')
  })

  it('editing footnote text keeps the italic formatting inside the footnote', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>正文</w:t></w:r>' +
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId32" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      extraParts: [
        {
          path: 'word/footnotes.xml',
          xml:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
            '<w:footnote w:id="2"><w:p>' +
            '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>' +
            '<w:r><w:t xml:space="preserve"> </w:t></w:r>' +
            '<w:r><w:rPr><w:i/></w:rPr><w:t>斜体脚注</w:t></w:r>' +
            '</w:p></w:footnote></w:footnotes>',
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        },
      ],
    })
    const doc = await parseDocx(bytes)
    expect(doc.footnotes[0].text).toBe('斜体脚注')
    const edited = [{ ...doc.footnotes[0], text: '斜体脚注(改)' }]
    const out = await saveDocx(doc, originalOrder(doc), { footnotes: edited })
    const xml = await (await JSZip.loadAsync(out)).file('word/footnotes.xml')!.async('string')
    expect(xml).toContain('<w:rPr><w:i/></w:rPr>')
    expect(xml).toContain('<w:footnoteRef/>')
    const reparsed = await parseDocx(out)
    expect(reparsed.footnotes[0].text).toBe('斜体脚注(改)')
    expect(reparsed.footnotes[0].richParas?.[0]?.some((r) => r.italic)).toBe(true)
  })
})
