import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { docxToText } from '../src/docx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
// entity-decoded characters, built from code points so the source stays ASCII
const CJK = String.fromCodePoint(0x4e2d)
const WINGDING = String.fromCodePoint(0xf0e0)
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'

async function docx(
  body: string,
  extra: Record<string, string> = {},
  opts: { mainPart?: string; styles?: string } = {},
): Promise<Uint8Array> {
  const zip = new JSZip()
  const main = opts.mainPart ?? 'word/document.xml'
  zip.file(
    '_rels/.rels',
    `<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/${main}"/></Relationships>`,
  )
  zip.file(
    main,
    `<?xml version="1.0"?><w:document ${W} ${MC}><w:body>${body}</w:body></w:document>`,
  )
  zip.file(
    'word/styles.xml',
    opts.styles ??
      `<w:styles ${W}><w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="Sub"><w:name w:val="Subhead"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>` +
        `<w:style w:type="paragraph" w:styleId="Bul"><w:name w:val="Bullets"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style>` +
        `<w:style w:type="paragraph" w:styleId="SubBul"><w:name w:val="Sub bullets"/><w:basedOn w:val="Bul"/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="MyH"><w:name w:val="My heading"/><w:basedOn w:val="1"/></w:style></w:styles>`,
  )
  for (const [path, xml] of Object.entries(extra)) zip.file(path, xml)
  return zip.generateAsync({ type: 'uint8array' })
}

const p = (text: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

describe('docxToText (XML scan)', () => {
  it('marks headings by style name, outline level and Heading<N> id; lists by numPr', async () => {
    const text = await docxToText(
      await docx(
        p('Title', '<w:pStyle w:val="1"/>') +
          p('Sub', '<w:pStyle w:val="Sub"/>') +
          p('Three', '<w:pStyle w:val="Heading3"/>') +
          p('Deep', '<w:outlineLvl w:val="8"/>') +
          p('item', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
          p('styled item', '<w:pStyle w:val="Bul"/>') +
          p('inherited item', '<w:pStyle w:val="SubBul"/>') +
          p('inherited heading', '<w:pStyle w:val="MyH"/>') +
          p('plain'),
      ),
    )
    expect(text.split('\n')).toEqual([
      '# Title',
      '## Sub',
      '### Three',
      '###### Deep',
      '- item',
      '- styled item',
      '- inherited item',
      '# inherited heading',
      'plain',
    ])
  })

  it('numId 0 switches numbering off, on the paragraph and on the style', async () => {
    const text = await docxToText(
      await docx(
        p(
          'styled but off',
          '<w:pStyle w:val="Bul"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>',
        ) +
          p('off by style', '<w:pStyle w:val="Off"/>') +
          p('on', '<w:numPr><w:numId w:val="2"/></w:numPr>'),
        {},
        {
          styles:
            `<w:styles ${W}><w:style w:type="paragraph" w:styleId="Bul"><w:name w:val="Bullets"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style>` +
            `<w:style w:type="paragraph" w:styleId="Off"><w:name w:val="Off"/><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style></w:styles>`,
        },
      ),
    )
    expect(text.split('\n')).toEqual(['styled but off', 'off by style', '- on'])
  })

  it('a self-closing paragraph inside a text box does not swallow the host paragraph', async () => {
    const text = await docxToText(
      await docx(
        `<w:p><w:r><w:t>before </w:t></w:r><w:r><w:drawing><w:txbxContent><w:p/><w:p/></w:txbxContent></w:drawing></w:r>` +
          `<w:r><w:t>after</w:t></w:r></w:p>` +
          `<w:tbl><w:tr><w:tc/><w:tc>${p('x')}</w:tc></w:tr></w:tbl>`,
      ),
    )
    expect(text).toBe('\n\nbefore after\n | x')
  })

  it('flattens tables one row per line with nested paragraphs joined, nested tables inside the cell', async () => {
    const cell = (inner: string) => `<w:tc><w:tcPr><w:tcW w:w="1"/></w:tcPr>${inner}</w:tc>`
    const nested = `<w:tbl><w:tr>${cell(p('n1'))}${cell(p('n2'))}</w:tr></w:tbl>`
    const text = await docxToText(
      await docx(
        `<w:tbl><w:tblPr><w:tblStyle w:val="x"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid>` +
          `<w:tr><w:trPr/>${cell(p('Metric'))}${cell(p('Value'))}</w:tr>` +
          `<w:tr>${cell(p('Revenue') + p('(net)'))}${cell(p('100') + nested)}</w:tr>` +
          `<w:tr>${cell('<w:p><w:r><w:t>week 1-4</w:t><w:br/><w:t>(4 weeks)</w:t></w:r></w:p>')}${cell(p('x'))}</w:tr></w:tbl>` +
          p('after'),
      ),
    )
    expect(text).toBe('Metric | Value\nRevenue (net) | 100 n1 | n2\nweek 1-4 (4 weeks) | x\nafter')
  })

  it('reads a text box once through mc:Choice, skipping the mc:Fallback twin', async () => {
    const box = (label: string) => `<w:txbxContent>${p(`box ${label}`)}</w:txbxContent>`
    const text = await docxToText(
      await docx(
        `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>${box('new')}</w:drawing></mc:Choice>` +
          `<mc:Fallback><w:pict>${box('old')}</w:pict></mc:Fallback></mc:AlternateContent></w:r>` +
          `<w:r><w:t>anchor</w:t></w:r></w:p>`,
      ),
    )
    expect(text).toBe('box new\nanchor')
  })

  it('drops hidden runs, tracked deletions and field codes; keeps field results, tabs, breaks, symbols and entities', async () => {
    const text = await docxToText(
      await docx(
        `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:rPr><w:vanish/></w:rPr></w:pPr>` +
          `<w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r>` +
          `<w:del><w:r><w:delText>gone</w:delText></w:r></w:del>` +
          `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r>` +
          `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
          `<w:r><w:tab/><w:t>a&amp;b &lt;c&gt; &#x4e2d;</w:t><w:br/><w:t>next</w:t><w:sym w:font="Wingdings" w:char="F0E0"/><w:noBreakHyphen/></w:r></w:p>` +
          `<w:p><w:pPr><w:pStyle w:val="Heading8"/></w:pPr><w:r><w:br w:type="column"/></w:r><w:r><w:t>column head</w:t></w:r></w:p>`,
      ),
    )
    expect(text).toBe(`7\ta&b <c> ${CJK}\nnext${WINGDING}-\n###### column head`)
  })

  it('appends footnotes and endnotes under headers, skipping separator notes', async () => {
    const notes = (tag: string, body: string) =>
      `<w:${tag}s ${W}><w:${tag} w:type="separator" w:id="-1">${p('sep')}</w:${tag}>${body}</w:${tag}s>`
    const text = await docxToText(
      await docx(p('body'), {
        'word/footnotes.xml': notes(
          'footnote',
          `<w:footnote w:id="1">${p('first note')}${p('more')}</w:footnote>`,
        ),
        'word/endnotes.xml': notes('endnote', `<w:endnote w:id="1">${p('an endnote')}</w:endnote>`),
      }),
    )
    expect(text).toBe('body\n## Footnotes\nfirst note more\n## Endnotes\nan endnote')
  })

  it('stops at the character budget instead of scanning the whole body', async () => {
    const body = Array.from({ length: 1000 }, (_, i) => p(`paragraph ${i}`)).join('')
    const text = await docxToText(await docx(body), { maxChars: 100 })
    expect(text.length).toBeGreaterThanOrEqual(100)
    expect(text.length).toBeLessThan(140)
    expect(text.startsWith('paragraph 0\nparagraph 1')).toBe(true)
  })

  it('finds a main part that is not word/document.xml through the package rels', async () => {
    const text = await docxToText(
      await docx(p('elsewhere'), {}, { mainPart: 'word/document2.xml' }),
    )
    expect(text).toBe('elsewhere')
  })

  it('rejects a package without a main document part', async () => {
    const zip = new JSZip()
    zip.file('_rels/.rels', '<Relationships/>')
    zip.file('word/styles.xml', `<w:styles ${W}/>`)
    await expect(docxToText(await zip.generateAsync({ type: 'uint8array' }))).rejects.toThrow(
      /main document/,
    )
  })
})
