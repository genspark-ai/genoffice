import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  bibliographyEntries,
  bibliographyFieldXml,
  citationFieldOf,
  citationInstruction,
  citationStyleOfXml,
  formatCitation,
  generateFieldTableXml,
  parseDocx,
  parseSourcesXml,
  saveDocx,
  type SaveBlock,
  type SourceInfo,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * Citations as Word keeps them: sources in the b:Sources part, each citation
 * a CITATION field whose result is the citation in the selected style, the
 * bibliography a BIBLIOGRAPHY field.
 */

const NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'
const SOURCES_XML =
  `<b:Sources SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6" xmlns:b="${NS}" xmlns="${NS}">` +
  '<b:Source><b:Tag>Mat11</b:Tag><b:SourceType>Book</b:SourceType><b:Guid>{AF584761}</b:Guid>' +
  '<b:Author><b:Author><b:NameList><b:Person><b:Last>Mattmann</b:Last><b:First>Chris</b:First></b:Person>' +
  '<b:Person><b:Last>Zitting</b:Last><b:First>Jukka</b:First></b:Person></b:NameList></b:Author></b:Author>' +
  '<b:Title>Tika in Action</b:Title><b:Year>2011</b:Year><b:City>Shelter Island</b:City>' +
  '<b:Publisher>Manning</b:Publisher><b:RefOrder>1</b:RefOrder></b:Source></b:Sources>'

const BOOK = parseSourcesXml(SOURCES_XML)[0]!
const ARTICLE: SourceInfo = {
  tag: 'Kay93',
  type: 'JournalArticle',
  author: 'Kay, Alan',
  title: 'The early history of Smalltalk',
  year: '1993',
  publisher: 'ACM SIGPLAN Notices',
}
const SOURCES = [BOOK, ARTICLE]

const fld = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`
const citationField = (tag: string, result: string): string =>
  '<w:sdt><w:sdtPr><w:id w:val="-2100173194"/><w:citation/></w:sdtPr><w:sdtContent>' +
  fld('begin') +
  `<w:r><w:instrText xml:space="preserve"> CITATION ${tag} \\l 1033 </w:instrText></w:r>` +
  fld('separate') +
  `<w:r><w:rPr><w:noProof/></w:rPr><w:t>${result}</w:t></w:r>` +
  fld('end') +
  '</w:sdtContent></w:sdt>'

const sourcesParts = [
  {
    path: 'customXml/item1.xml',
    xml: SOURCES_XML,
    contentType: 'application/xml',
  },
]

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

describe('sources', () => {
  it('reads every author of a source, and its city', () => {
    expect(BOOK.author).toBe('Mattmann, Chris; Zitting, Jukka')
    expect(BOOK.city).toBe('Shelter Island')
  })

  it('names the style of the part when it is one the picker offers', () => {
    expect(citationStyleOfXml(SOURCES_XML)).toBe('apa')
    expect(
      citationStyleOfXml(
        SOURCES_XML.replace('StyleName="APA"', 'StyleName="IEEE"').replace(
          'APASixthEditionOfficeOnline',
          'IEEE2006OfficeOnline',
        ),
      ),
    ).toBe('ieee')
    expect(
      citationStyleOfXml(
        '<b:Sources SelectedStyle="\\GB7714.xsl" StyleName="GB7714" Version="2005">',
      ),
    ).toBeNull()
  })
})

describe('citation styles', () => {
  const cite = (tag: string, style: Parameters<typeof formatCitation>[2], order: string[] = []) =>
    formatCitation(citationInstruction(tag), SOURCES, style, order)

  it('writes the in-text citation each style asks for', () => {
    expect(cite('Mat11', 'apa')).toBe('(Mattmann & Zitting, 2011)')
    expect(cite('Mat11', 'mla')).toBe('(Mattmann and Zitting)')
    expect(cite('Mat11', 'chicago')).toBe('(Mattmann and Zitting 2011)')
    expect(cite('Mat11', 'harvard')).toBe('(Mattmann & Zitting, 2011)')
    expect(cite('Kay93', 'ieee', ['Kay93', 'Mat11'])).toBe('[1]')
    expect(cite('Mat11', 'ieee', ['Kay93', 'Mat11'])).toBe('[2]')
  })

  it('a citation of a source the document does not hold keeps what it shows', () => {
    expect(cite('Nobody', 'apa')).toBeNull()
  })

  it('reads the switches of a CITATION field', () => {
    expect(citationFieldOf(' CITATION Mat11 \\p 12 \\l 1033 \\m Kay93 ')).toEqual({
      tags: ['Mat11', 'Kay93'],
      pages: '12',
    })
    expect(formatCitation('CITATION Mat11 \\p 12 \\m Kay93', SOURCES, 'apa')).toBe(
      '(Mattmann & Zitting, 2011, p. 12; Kay, 1993, p. 12)',
    )
    expect(formatCitation('CITATION Mat11 \\n', SOURCES, 'apa')).toBe('(Tika in Action, 2011)')
  })

  const lines = (style: Parameters<typeof bibliographyEntries>[1], order: string[] = []) =>
    bibliographyEntries(SOURCES, style, order).map((segs) => segs.map((s) => s.text).join(''))

  it('writes the bibliography of each style, by author or by number', () => {
    expect(lines('apa')).toEqual([
      'Kay, A. (1993). The early history of Smalltalk. ACM SIGPLAN Notices.',
      'Mattmann, C., & Zitting, J. (2011). Tika in Action. Shelter Island: Manning.',
    ])
    expect(lines('mla')[1]).toBe(
      'Mattmann, Chris and Jukka Zitting. Tika in Action. Shelter Island: Manning. 2011.',
    )
    expect(lines('chicago')[1]).toBe(
      'Mattmann, Chris, and Jukka Zitting. 2011. Tika in Action. Shelter Island: Manning.',
    )
    expect(lines('harvard')[1]).toBe(
      'Mattmann, C. & Zitting, J., 2011. Tika in Action. Shelter Island: Manning.',
    )
    expect(lines('ieee', ['Mat11'])).toEqual([
      '[1] C. Mattmann and J. Zitting, Tika in Action, Shelter Island: Manning, 2011.',
      '[2] A. Kay, “The early history of Smalltalk”, ACM SIGPLAN Notices, 1993.',
    ])
  })

  it('sets a book title and a journal name in italics', () => {
    const [article, book] = bibliographyEntries(SOURCES, 'apa')
    expect(book!.filter((s) => s.italic).map((s) => s.text)).toEqual(['Tika in Action'])
    expect(article!.filter((s) => s.italic).map((s) => s.text)).toEqual(['ACM SIGPLAN Notices'])
    const xml = bibliographyFieldXml([book!, article!])
    expect(xml).toHaveLength(2)
    expect(xml[0]).toContain(' BIBLIOGRAPHY ')
    expect(xml[0]).toContain('<w:i/><w:iCs/><w:noProof/></w:rPr><w:t xml:space="preserve">Tika')
    expect(xml[1]).toMatch(/fldCharType="end"\/><\/w:r><\/w:p>$/)
  })
})

describe('a document that arrives with sources and CITATION fields', () => {
  const body =
    '<w:p><w:r><w:t xml:space="preserve">As shown </w:t></w:r>' +
    citationField('Mat11', '(Mattmann &amp; Zitting, 2011)') +
    '<w:r><w:t>.</w:t></w:r></w:p>'

  it('keeps the citation as a field run in a paragraph that can be edited', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: body, extraParts: sourcesParts }))
    expect(doc.citationStyle).toBe('apa')
    expect(doc.sources.map((s) => s.tag)).toEqual(['Mat11'])
    const block = doc.blocks[0]!
    expect(block.type).toBe('paragraph')
    const field = block.runs!.find((r) => r.instrField !== undefined)!
    expect(field.instrField).toBe('CITATION Mat11 \\l 1033')
    expect(field.text).toBe('(Mattmann & Zitting, 2011)')
  })

  it('writes the restyled citation back as a CITATION field and the style into the sources part', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: body, extraParts: sourcesParts }))
    const runs = doc.blocks[0]!.runs!.map((r) =>
      r.instrField === undefined
        ? r
        : { ...r, text: formatCitation(r.instrField, doc.sources, 'mla')! },
    )
    const out = await saveDocx(doc, [{ kind: 'generated', block: { type: 'paragraph', runs } }], {
      sources: doc.sources,
      sourcesStyle: 'mla',
    })
    const zip = await JSZip.loadAsync(out)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toMatch(
      /fldCharType="begin"\/><\/w:r><w:r><w:instrText[^>]*> CITATION Mat11 \\l 1033 <\/w:instrText>/,
    )
    expect(xml).toContain('(Mattmann and Zitting)')
    const sources = await zip.file('customXml/item1.xml')!.async('string')
    expect(sources).toContain('SelectedStyle="\\MLASeventhEditionOfficeOnline.xsl"')
    expect(sources).toContain('StyleName="MLA"')
    // the source itself goes back as it was read, the fields the model does not hold included
    expect(sources).toContain('<b:Guid>{AF584761}</b:Guid>')
    const again = await parseDocx(out)
    expect(again.citationStyle).toBe('mla')
    expect(again.blocks[0]!.runs!.find((r) => r.instrField)!.text).toBe('(Mattmann and Zitting)')
  })

  it('a save that changes no source leaves the sources part and its style alone', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: body, extraParts: sourcesParts }))
    const out = await saveDocx(doc, originalOrder(doc), {})
    const sources = await (await JSZip.loadAsync(out)).file('customXml/item1.xml')!.async('string')
    expect(sources).toBe(SOURCES_XML)
  })
})

describe('index entry switches', () => {
  const xe = (instr: string): string =>
    fld('begin') +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    fld('end')

  it('reads what follows the entry text and writes it back', async () => {
    const body =
      '<w:p><w:r><w:t>Apples</w:t></w:r>' +
      xe('XE "Fruit:Apples" \\t "See Pomes" \\b') +
      xe('XE "Plain"') +
      '</w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const marks = doc.blocks[0]!.runs!.filter((r) => r.xeTerm !== undefined)
    expect(marks.map((r) => [r.xeTerm, r.xeSwitches])).toEqual([
      ['Fruit:Apples', '\\t "See Pomes" \\b'],
      ['Plain', undefined],
    ])
    const out = await saveDocx(doc, [
      { kind: 'generated', block: { type: 'paragraph', runs: doc.blocks[0]!.runs! } },
    ])
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    expect(xml).toContain(' XE "Fruit:Apples" \\t "See Pomes" \\b </w:instrText>')
    expect(xml).toContain(' XE "Plain" </w:instrText>')
  })
})

describe('an index with pages behind a right tab', () => {
  it('reads a bold or italic page back as a run of its own, in every line of the field', async () => {
    const pageRuns = [{ text: '1', bold: true }, { text: ', ' }, { text: '3', italic: true }]
    const xml = generateFieldTableXml(
      [
        { text: 'Apples', style: 'Index1', pageNo: '1, 3', pageRuns },
        { text: 'Fruit', style: 'Index1', pageNo: '1, 3', pageRuns },
        { text: 'Pears', style: 'Index2', indent: 240, pageNo: '2' },
      ],
      { instr: ' INDEX \\e "\t" \\c "2" ', leader: 'dot', tabPos: 4300 },
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: xml.join('') }))
    const shown = doc.blocks.slice(0, 3).map((b) => b.fieldDisplay)
    expect(shown.map((d) => [d?.kind, d?.left, d?.right, d?.rightRuns])).toEqual([
      ['tocLine', 'Apples', '1, 3', pageRuns],
      ['tocLine', 'Fruit', '1, 3', pageRuns],
      ['tocLine', 'Pears', '2', undefined],
    ])
  })
})

describe('mail-merge rule fields', () => {
  const NO_RESULT = String.fromCharCode(0x200b)
  const field = (instr: string, result = ''): string =>
    fld('begin') +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    fld('separate') +
    (result ? `<w:r><w:t>${result}</w:t></w:r>` : '') +
    fld('end')
  const nested = (code: string): string =>
    fld('begin') +
    `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>` +
    fld('separate') +
    '<w:r><w:t>x</w:t></w:r>' +
    fld('end')

  it('are field runs of a paragraph that can be edited, so the merge finds them after a reopen', async () => {
    const body =
      '<w:p><w:r><w:t xml:space="preserve">No. </w:t></w:r>' +
      field('MERGEREC', '4') +
      field('NEXT') +
      field('SET Greeting "Hello"') +
      field('ASK Reason "Why?" \\d "none" \\o') +
      field('FILLIN "Who signs?" \\d "The team"', 'The team') +
      '</w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(doc.blocks[0]!.type).toBe('paragraph')
    expect(
      doc.blocks[0]!.runs!.filter((r) => r.instrField).map((r) => [r.instrField, r.text]),
    ).toEqual([
      ['MERGEREC', '4'],
      ['NEXT', NO_RESULT],
      ['SET Greeting "Hello"', NO_RESULT],
      ['ASK Reason "Why?" \\d "none" \\o', NO_RESULT],
      ['FILLIN "Who signs?" \\d "The team"', 'The team'],
    ])
  })

  it('a field read without a result is written without one', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: `<w:p><w:r><w:t>Label</w:t></w:r>${field('NEXT')}</w:p>` }),
    )
    const out = await saveDocx(doc, [
      { kind: 'generated', block: { type: 'paragraph', runs: doc.blocks[0]!.runs! } },
    ])
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    expect(xml).toMatch(
      /fldCharType="begin"\/><\/w:r><w:r><w:instrText[^>]*> NEXT <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="end"\/>/,
    )
    expect(xml).not.toContain(NO_RESULT)
  })

  it('a NEXTIF and a SKIPIF carry their data column as the IF does', async () => {
    const rule = (keyword: string) =>
      fld('begin') +
      `<w:r><w:instrText xml:space="preserve"> ${keyword} </w:instrText></w:r>` +
      nested('MERGEFIELD City') +
      '<w:r><w:instrText xml:space="preserve"> = "London" </w:instrText></w:r>' +
      fld('end')
    const doc = await parseDocx(
      await buildDocx({ bodyXml: `<w:p>${rule('NEXTIF')}${rule('SKIPIF')}</w:p>` }),
    )
    expect(doc.blocks[0]!.type).toBe('paragraph')
    expect(doc.blocks[0]!.runs!.map((r) => r.instrField)).toEqual([
      'NEXTIF «City» = "London"',
      'SKIPIF «City» = "London"',
    ])
  })

  it('a SET that Word has run keeps its bookmark: the paragraph stays as it was read', async () => {
    const body =
      '<w:p>' +
      fld('begin') +
      '<w:r><w:instrText xml:space="preserve"> SET Greeting "Hello" </w:instrText></w:r>' +
      fld('separate') +
      '<w:bookmarkStart w:id="1" w:name="Greeting"/><w:r><w:t>Hello</w:t></w:r><w:bookmarkEnd w:id="1"/>' +
      fld('end') +
      '</w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(doc.blocks[0]!.type).toBe('passthrough')
  })
})
