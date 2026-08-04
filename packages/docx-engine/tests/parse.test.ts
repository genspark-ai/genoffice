import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildKitchenSinkDocx } from './helpers/build-docx'

describe('parseDocx', () => {
  it('parses the kitchen-sink document into anchored blocks', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)

    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => b.type)).toEqual([
      'heading',
      'paragraph',
      'heading',
      'listItem',
      'listItem',
      'listItem',
      'paragraph',
      'table',
      'image',
      'passthrough',
      'paragraph',
    ])

    // every block anchored to its top-level element index with exact XML slice
    doc.blocks.forEach((block, i) => {
      expect(block.docxIndex).toBe(i)
      expect(block.originalXml).toBeTruthy()
      expect(doc.internal.documentXml).toContain(block.originalXml!)
    })

    const h1 = visible[0]
    expect(h1.level).toBe(1)
    expect(h1.runs?.[0].text).toBe('第一章 概述')

    const rich = visible[1]
    const runs = rich.runs!
    expect(runs.find((r) => r.bold)?.text).toBe('加粗')
    expect(runs.find((r) => r.italic)?.text).toBe('斜体')
    expect(runs.find((r) => r.underline)?.text).toBe('下划线')
    expect(runs.find((r) => r.strike)?.text).toBe('删除线')
    const colored = runs.find((r) => r.color === 'FF0000')
    expect(colored?.text).toBe('红色14号')
    expect(colored?.sizeHalfPoints).toBe(28)
    // XML entities decoded
    expect(runs[runs.length - 1].text).toContain('& < >')

    const bullets = visible.filter((b) => b.type === 'listItem')
    expect(bullets[0].list).toEqual({ kind: 'bullet', numId: '1', ilvl: 0 })
    expect(bullets[2].list).toEqual({ kind: 'ordered', numId: '2', ilvl: 0 })

    const linkPara = visible[6]
    const linkRun = linkPara.runs!.find((r) => r.link)
    expect(linkRun?.link?.href).toBe('https://example.com/')
    expect(linkRun?.link?.rId).toBe('rId20')

    const table = visible[7]
    expect(table.label).toBe('Table 2×2')
    expect(table.previewText).toContain('A1')

    const image = visible[8]
    expect(image.imageDataUrl).toMatch(/^data:image\/png;base64,/)

    const math = visible[9]
    expect(math.label).toBe('Equation')

    // trailing sectPr is hidden
    const hidden = doc.blocks.filter((b) => b.hidden)
    expect(hidden).toHaveLength(1)
    expect(hidden[0].originalXml).toContain('<w:sectPr>')

    // style metadata for generation
    expect(doc.headingStyleIds.get(1)).toBe('Heading1')
    expect(doc.headingStyleIds.get(2)).toBe('Heading2')
    expect(doc.listParagraphStyleId).toBe('ListParagraph')
  })

  it('protects structural paragraphs (section break / fields) as passthrough', async () => {
    const { buildDocx } = await import('./helpers/build-docx')
    const bytes = await buildDocx({
      bodyXml: [
        // mid-document section break lives inside w:pPr
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:pPr></w:p>',
        // TOC field (complex fields still protect the whole paragraph; simple inline
        // fields PAGE/DATE are already editable)
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
        '<w:p><w:r><w:t>带脚注的段落</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
        // comments stay editable: they are tracked at run level, not as passthrough
        '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>有批注</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>',
        '<w:p><w:r><w:t>普通可编辑段落</w:t></w:r></w:p>',
      ].join(''),
    })
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => b.type)).toEqual([
      'passthrough',
      'passthrough',
      'paragraph',
      'paragraph',
      'paragraph',
    ])
    expect(visible[0].label).toBe('Section break paragraph')
    expect(visible[1].label).toBe('Auto TOC (updates when opened in Word)')
    // footnote references stay editable as atomic noteRef runs
    expect(visible[2].runs).toEqual([
      { text: '带脚注的段落' },
      { text: '*', noteRef: { kind: 'footnote', id: '2' } },
    ])
    expect(visible[3].runs).toEqual([{ text: '有批注', commentIds: ['0'] }])
  })

  it('detects headings by effective outline level, not only Heading1-style paragraphs', async () => {
    const { buildDocx } = await import('./helpers/build-docx')
    const p = (pPr: string, text: string) =>
      `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    const bytes = await buildDocx({
      extraStylesXml:
        // custom style with no own outlineLvl: inherits level 2 through basedOn
        '<w:style w:type="paragraph" w:styleId="MySub"><w:name w:val="My Sub"/><w:basedOn w:val="Heading2"/></w:style>' +
        // Word's TOCHeading pattern: basedOn Heading1 but outlineLvl 9 = body text
        '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/>' +
        '<w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>',
      bodyXml: [
        p('<w:pStyle w:val="Heading1"/>', 'h1 via built-in style'),
        p('<w:pStyle w:val="Heading2"/>', 'h2 via built-in style'),
        p('<w:pStyle w:val="Heading3"/>', 'h3 via undefined built-in styleId'),
        p('<w:outlineLvl w:val="2"/>', 'h3 via direct outlineLvl'),
        p('<w:pStyle w:val="MySub"/>', 'h2 inherited through basedOn'),
        p('<w:pStyle w:val="TOCHeading"/>', 'body text: style outlineLvl 9'),
        p(
          '<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="9"/>',
          'body text: direct outlineLvl 9',
        ),
      ].join(''),
    })
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => [b.type, b.level])).toEqual([
      ['heading', 1],
      ['heading', 2],
      ['heading', 3],
      ['heading', 3],
      ['heading', 2],
      ['paragraph', undefined],
      ['paragraph', undefined],
    ])
  })
})
