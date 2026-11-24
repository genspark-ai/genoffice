import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  STYLEREF_NOT_FOUND,
  SeqCounter,
  captionNumberingOfXml,
  computeListMarkerInfos,
  generateCaptionXml,
  generateFieldTableXml,
  generateTableOfAuthoritiesFieldXml,
  generateTableOfFiguresFieldXml,
  generateTocFieldXml,
  parseDocx,
  numberingEventsOfXml,
  parseSeqInstruction,
  parseStyleRefInstruction,
  rewriteSeqResultsXml,
  saveDocx,
  seqFieldsOfXml,
  setCaptionNumberingXml,
  styleRefResult,
  unreliableSeqLabels,
  type NumberingDef,
  type StyledParagraph,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'
import { allowDirtyFields } from './helpers/dirty-fields'

const field = (instr: string, result: string, dirty = false) =>
  `<w:r><w:fldChar w:fldCharType="begin"${dirty ? ' w:dirty="true"' : ''}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

const simple = (instr: string, result: string) =>
  `<w:fldSimple w:instr=" ${instr} "><w:r><w:rPr><w:noProof/></w:rPr><w:t>${result}</w:t></w:r></w:fldSimple>`

describe('SEQ fields of a block', () => {
  it('are read from complex and simple fields, nested ones included, in order', () => {
    const xml =
      `<w:p><w:r><w:t>Figure </w:t></w:r>${field('SEQ Figure \\* ARABIC', '4', true)}` +
      `<w:r><w:t> and </w:t></w:r>${simple('SEQ &quot;My Table&quot; \\* ROMAN', 'II')}` +
      `${field('DATE', '2026')}` +
      // a SEQ inside another field's result
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF 1 = 1 </w:instrText></w:r>' +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${field('SEQ Figure', '5')}` +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const fields = seqFieldsOfXml(xml)
    expect(fields.map((f) => [f.instr, f.result, f.dirtyAttr !== null])).toEqual([
      ['SEQ Figure \\* ARABIC', '4', true],
      ['SEQ "My Table" \\* ROMAN', 'II', false],
      ['SEQ Figure', '5', false],
    ])
    for (const f of fields) expect(xml.slice(f.spans[0][0], f.spans[0][1])).toBe(f.result)
    expect(seqFieldsOfXml('<w:p><w:r><w:t>SEQ is only text here</w:t></w:r></w:p>')).toEqual([])
  })

  it('instructions give the label, the switches and the number format', () => {
    expect(parseSeqInstruction('SEQ Figure \\* ARABIC')).toMatchObject({
      id: 'Figure',
      restart: null,
      repeat: false,
      hidden: false,
      format: 'ARABIC',
      unsupported: false,
    })
    expect(parseSeqInstruction(' SEQ "My Table" \\* roman \\r 3 ')).toMatchObject({
      id: 'My Table',
      restart: 3,
      format: 'roman',
    })
    expect(parseSeqInstruction('SEQ Figure \\c \\* MERGEFORMAT')).toMatchObject({
      repeat: true,
      format: 'ARABIC',
    })
    expect(parseSeqInstruction('SEQ Figure \\h')).toMatchObject({ hidden: true })
    // a format switch shows a hidden field after all
    expect(parseSeqInstruction('SEQ Figure \\h \\* ARABIC')).toMatchObject({ hidden: false })
    expect(parseSeqInstruction('SEQ Figure \\* Ordinal')?.format).toBeNull()
    expect(parseSeqInstruction('SEQ Figure \\* Ordinal')?.formatName).toBe('Ordinal')
    expect(parseSeqInstruction('SEQ Figure \\* ARABIC \\s 2')).toMatchObject({
      headingLevel: 2,
      unsupported: false,
    })
    expect(parseSeqInstruction('SEQ Figure \\s x')?.unsupported).toBe(true)
    expect(parseSeqInstruction('REF Figure')).toBeNull()
  })

  it('are counted per label: next, repeat, restart, hidden', () => {
    const counter = new SeqCounter()
    const next = (instr: string) => counter.next(parseSeqInstruction(instr)!)
    expect(next('SEQ Figure').text).toBe('1')
    expect(next('SEQ Table \\* ROMAN').text).toBe('I')
    expect(next('SEQ Figure').text).toBe('2')
    expect(next('SEQ Figure \\c').text).toBe('2')
    expect(next('SEQ Figure \\h')).toEqual({ value: 3, text: null })
    expect(next('SEQ Figure \\r 10').text).toBe('10')
    expect(next('SEQ Figure \\* alphabetic').text).toBe('k')
    expect(next('SEQ Table \\* Ordinal')).toEqual({ value: 2, text: '2nd' })
    // a format with no rule here: the number is counted, the text is Word's
    expect(next('SEQ Other \\* THAILETTER')).toEqual({ value: 1, text: null })
    expect(next('SEQ Table \\* ROMAN').text).toBe('III')
  })

  it('take new results in place; a rewritten field loses w:dirty, the others keep theirs', () => {
    const xml =
      `<w:p>${field('SEQ Figure', '1', true)}${field('SEQ Figure', '2', true)}` +
      `${simple('SEQ Figure', '3')}</w:p>`
    const out = rewriteSeqResultsXml(xml, ['7', null, '<9>'])
    expect(seqFieldsOfXml(out).map((f) => [f.result, f.dirtyAttr !== null])).toEqual([
      ['7', false],
      ['2', true],
      ['<9>', false],
    ])
    expect(out).toContain('<w:t>&lt;9&gt;</w:t>')
    expect(rewriteSeqResultsXml(xml, [null, null, null])).toBe(xml)
  })

  it('a label is unreliable when its cached numbers are not the count in order', () => {
    const cap = (label: string, n: number) => generateCaptionXml(label, n, 'x')
    expect([...unreliableSeqLabels([cap('Figure', 1), cap('Table', 1), cap('Figure', 2)])]).toEqual(
      [],
    )
    expect([
      ...unreliableSeqLabels([cap('Figure', 1), cap('Table', 1), undefined, cap('Figure', 3)]),
    ]).toEqual(['Figure'])
    expect([...unreliableSeqLabels([`<w:p>${field('SEQ Eq \\s 1', '1')}</w:p>`])]).toEqual(['Eq'])
  })
})

describe('generated fields are saved with their result and no w:dirty', () => {
  const documentXml = async (bytes: Uint8Array) =>
    (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

  it('captions, a table of contents, of figures, of authorities and an index', async () => {
    const body =
      generateCaptionXml('Figure', 1, 'Overview') +
      generateCaptionXml('Table', 2, 'Totals', { format: 'ROMAN' }) +
      generateTocFieldXml([
        { level: 1, text: 'Scope', pageNo: 1 },
        { level: 2, text: 'Detail', pageNo: 3 },
      ]).join('') +
      generateTableOfFiguresFieldXml([{ text: 'Figure 1 Overview', pageNo: 2 }], 'Figure').join(
        '',
      ) +
      generateTableOfAuthoritiesFieldXml(
        'Cases',
        [{ text: 'Roe v. Wade', pageNo: '4, 6' }],
        1,
      ).join('') +
      generateFieldTableXml(
        [
          { text: 'A', style: 'IndexHeading', noPage: true },
          { text: 'apple', style: 'Index1', pageNo: '2, 5' },
        ],
        { instr: ' INDEX \\h "A" \\c "2" ', pageSep: ', ' },
      ).join('')
    expect(body).not.toContain('w:dirty')
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const saved = await saveDocx(
      parsed,
      parsed.blocks.map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    const xml = await documentXml(saved)
    expect(xml).not.toContain('w:dirty')
    expect(xml).toMatch(/ SEQ Figure \\\* ARABIC <\/w:instrText>.*?<w:t>1<\/w:t>/)
    expect(xml).toMatch(/ SEQ Table \\\* ROMAN <\/w:instrText>.*?<w:t>II<\/w:t>/)
    expect(xml).toMatch(/Detail<\/w:t><\/w:r>.*?<w:tab\/><\/w:r><w:r>.*?<w:t>3<\/w:t>/)
    expect(xml).toMatch(/Figure 1 Overview<\/w:t><\/w:r>.*?<w:t>2<\/w:t>/)
    expect(xml).toContain('4, 6')
    expect(xml).toContain('apple')
    expect(xml).toMatch(/, 2, 5<\/w:t>/)
  })

  it('a table built before its pages are known is the one left for Word', () => {
    allowDirtyFields()
    expect(generateTocFieldXml([{ level: 1, text: 'Scope' }])[0]).toContain('w:dirty="true"')
    // a table that shows no page numbers at all is complete as it stands
    expect(generateTocFieldXml([{ level: 1, text: 'Scope', noPage: true }])[0]).not.toContain(
      'w:dirty',
    )
  })
})

const LEVELS = new Map([
  ['Heading1', 1],
  ['Heading2', 2],
])

const heading = (level: number, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

const chaptered = (chapter: string, n: string, level = 1) =>
  `<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">Figure </w:t></w:r>` +
  `${field(`STYLEREF ${level} \\s`, chapter)}<w:r><w:noBreakHyphen/></w:r>` +
  `${field(`SEQ Figure \\* ARABIC \\s ${level}`, n)}</w:p>`

describe('SEQ \\s: the count starts again at a heading', () => {
  it('restarts after a heading of its level or above it, each field by its own switch', () => {
    const counter = new SeqCounter()
    const next = (instr: string) => counter.next(parseSeqInstruction(instr)!).text
    counter.heading(1)
    expect(next('SEQ Figure \\s 1')).toBe('1')
    expect(next('SEQ Figure \\s 1')).toBe('2')
    counter.heading(2)
    // a lower heading does not restart a count bound to level 1, but does one bound to level 2
    expect(next('SEQ Figure \\s 1')).toBe('3')
    expect(next('SEQ Table \\s 2')).toBe('1')
    expect(next('SEQ Table \\s 2')).toBe('2')
    counter.heading(1)
    expect(next('SEQ Figure \\* ROMAN \\s 1')).toBe('I')
    expect(next('SEQ Table \\s 2')).toBe('1')
    // a field without the switch goes on counting
    expect(next('SEQ Figure')).toBe('2')
    counter.heading(1)
    expect(next('SEQ Figure')).toBe('3')
    expect(next('SEQ Figure \\s 1 \\r 7')).toBe('7')
  })

  it('a file numbered by chapter counts as reliable once the heading styles are known', () => {
    const blocks = [
      heading(1, 'One'),
      chaptered('1', '1'),
      chaptered('1', '2'),
      // a heading and a caption inside a table
      `<w:tbl><w:tr><w:tc>${heading(1, 'Two')}${chaptered('2', '1')}</w:tc></w:tr></w:tbl>`,
      chaptered('2', '2'),
      '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>',
      heading(2, 'Sub'),
      chaptered('2', '3'),
    ]
    expect([...unreliableSeqLabels(blocks, LEVELS)]).toEqual([])
    // without the styles nothing tells where the count restarts
    expect([...unreliableSeqLabels(blocks)]).toEqual(['Figure'])
    blocks[4] = chaptered('2', '5')
    expect([...unreliableSeqLabels(blocks, LEVELS)]).toEqual(['Figure'])
  })

  it('a block gives its styled paragraphs and fields in order, a replaced style left out', () => {
    const xml =
      `<w:tbl><w:tr><w:tc>${chaptered('1', '4')}${heading(1, 'Two')}` +
      '<w:p><w:pPr><w:pStyle w:val="Body"/><w:pPrChange><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange></w:pPr></w:p>' +
      `${chaptered('2', '1')}</w:tc></w:tr></w:tbl>`
    expect(
      numberingEventsOfXml(xml, LEVELS).map((e) =>
        e.kind === 'field' ? `${e.field.kind} ${e.field.result}` : `${e.styleId} ${e.level}`,
      ),
    ).toEqual([
      'Caption undefined',
      'STYLEREF 1',
      'SEQ 4',
      'Heading1 1',
      'Body undefined',
      'Caption undefined',
      'STYLEREF 2',
      'SEQ 1',
    ])
  })
})

describe('STYLEREF', () => {
  const para = (over: Partial<StyledParagraph>): StyledParagraph => ({
    styleId: 'Normal',
    text: '',
    ...over,
  })
  const h1 = (text: string, marker?: string, number?: string) =>
    para({ styleId: 'Heading1', styleName: 'heading 1', headingLevel: 1, text, marker, number })
  const result = (instr: string, paragraphs: StyledParagraph[], before: number) =>
    styleRefResult(parseStyleRefInstruction(instr)!, paragraphs, before)

  it('reads the level or the style name and what to show', () => {
    expect(parseStyleRefInstruction(' STYLEREF 1 \\s ')).toEqual({
      level: 1,
      style: null,
      show: 'bare',
      context: null,
      aboveBelow: false,
      textCase: null,
      bareContext: false,
      unsupported: false,
    })
    expect(parseStyleRefInstruction('STYLEREF "Heading 1" \\* MERGEFORMAT')).toMatchObject({
      level: null,
      style: 'Heading 1',
      show: 'text',
      unsupported: false,
    })
    expect(parseStyleRefInstruction('STYLEREF Title \\n')).toMatchObject({ show: 'number' })
    expect(parseStyleRefInstruction('STYLEREF Title \\n \\t')).toMatchObject({ show: 'bare' })
    expect(parseStyleRefInstruction('STYLEREF 1 \\w')).toMatchObject({
      context: 'full',
      unsupported: false,
    })
    expect(parseStyleRefInstruction('STYLEREF 1 \\r \\p')).toMatchObject({
      context: 'relative',
      aboveBelow: true,
    })
    expect(parseStyleRefInstruction('STYLEREF 1 \\r \\t')).toMatchObject({
      context: 'relative',
      bareContext: true,
      unsupported: false,
    })
    // in the text of a document the search runs back from the field with or without it
    expect(parseStyleRefInstruction('STYLEREF 1 \\l')).toEqual(
      parseStyleRefInstruction('STYLEREF 1'),
    )
    expect(parseStyleRefInstruction('STYLEREF 1 \\q')?.unsupported).toBe(true)
    expect(parseStyleRefInstruction('STYLEREF 1 \\* Upper')).toMatchObject({
      textCase: 'Upper',
      unsupported: false,
    })
    expect(parseStyleRefInstruction('STYLEREF 1 \\* roman')?.unsupported).toBe(true)
    expect(parseStyleRefInstruction('SEQ Figure')).toBeNull()
  })

  it('takes the nearest paragraph of the style before it, else the first after it', () => {
    const paragraphs = [
      para({ text: 'Intro' }),
      h1('Alpha', 'Chapter 1.', '1'),
      para({ styleId: 'Heading2', headingLevel: 2, text: 'Sub', marker: '1.1', number: '1.1' }),
      // across a section break and a table cell the order is all that counts
      para({ text: '' }),
      h1('Beta', 'Chapter 2.', '2'),
      para({ text: 'Body' }),
    ]
    expect(result('STYLEREF 1 \\s', paragraphs, 0)).toBe('1')
    expect(result('STYLEREF 1 \\s', paragraphs, 4)).toBe('1')
    expect(result('STYLEREF 1 \\s', paragraphs, 5)).toBe('2')
    expect(result('STYLEREF 2 \\s', paragraphs, 6)).toBe('1.1')
    expect(result('STYLEREF 1 \\n', paragraphs, 6)).toBe('Chapter 2')
    expect(result('STYLEREF 1', paragraphs, 6)).toBe('Beta')
    expect(result('STYLEREF "heading 1"', paragraphs, 2)).toBe('Alpha')
    expect(result('STYLEREF Heading1', paragraphs, 2)).toBe('Alpha')
    // no numbering to ask: the number in context is not known
    expect(result('STYLEREF 1 \\w', paragraphs, 2)).toBeNull()
  })

  it('shows the number in context and where the paragraph is', () => {
    const paragraphs = [
      h1('Alpha', 'Chapter 1.', '1'),
      para({ styleId: 'Heading2', headingLevel: 2, text: 'Sub', marker: '(a)', number: 'a' }),
      para({ text: 'Body' }),
      h1('Beta', 'Chapter 2.', '2'),
    ]
    const around = {
      // without the words: one level's numbers, and nothing where two levels are joined
      numberOf: (p: StyledParagraph, context: 'relative' | 'full', bare: boolean) =>
        bare
          ? context === 'full'
            ? null
            : (p.number ?? null)
          : context === 'full'
            ? `1${p.marker}`
            : (p.marker ?? null),
      words: { above: 'above', below: 'below' },
    }
    const placed = (instr: string, before: number) =>
      styleRefResult(parseStyleRefInstruction(instr)!, paragraphs, before, around)
    expect(placed('STYLEREF 2 \\w', 3)).toBe('1(a)')
    expect(placed('STYLEREF 2 \\r', 3)).toBe('(a)')
    expect(placed('STYLEREF 2 \\r \\t', 3)).toBe('a')
    expect(placed('STYLEREF 2 \\r \\t \\p', 3)).toBe('a above')
    expect(placed('STYLEREF 2 \\w \\t', 3)).toBeNull()
    expect(placed('STYLEREF 2 \\l', 3)).toBe('Sub')
    expect(placed('STYLEREF 1 \\l \\n', 2)).toBe('Chapter 1')
    // alone it says where the paragraph is; after a number it follows it
    expect(placed('STYLEREF 2 \\p', 3)).toBe('above')
    expect(placed('STYLEREF 2 \\p', 0)).toBe('below')
    expect(placed('STYLEREF 2 \\n \\p', 3)).toBe('(a) above')
    expect(placed('STYLEREF 2 \\w \\p', 0)).toBe('1(a) below')
    // a heading without a number is 0 in context too
    expect(
      styleRefResult(parseStyleRefInstruction('STYLEREF 1 \\r')!, [h1('Plain')], 1, around),
    ).toBe('0')
    // without the words of the interface the field stays as it is
    expect(result('STYLEREF 2 \\p', paragraphs, 3)).toBeNull()
  })

  it('a heading without a number gives 0, no heading gives the error text Word saves', () => {
    expect(result('STYLEREF 1 \\s', [h1('Plain')], 1)).toBe('0')
    expect(result('STYLEREF 1 \\n', [h1('Plain')], 1)).toBe('0')
    expect(result('STYLEREF 3 \\s', [h1('Plain')], 1)).toBe(STYLEREF_NOT_FOUND)
    // a heading kept as XML: what it shows is not known, the result stays
    expect(result('STYLEREF 1 \\s', [{ ...h1(''), unknown: true }], 1)).toBeNull()
  })

  it('a list marker carries its numbers without the words around them', () => {
    const level = (lvlText: string, numFmt = 'decimal') => ({ numFmt, lvlText, start: 1 })
    const defs = new Map<string, NumberingDef>([
      [
        '1',
        {
          numId: '1',
          abstractNumId: '1',
          startOverrides: {},
          levels: [level('Chapter %1.'), level('%1.%2'), level('(%3)', 'lowerLetter')],
        } as unknown as NumberingDef,
      ],
    ])
    const markers = computeListMarkerInfos(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 1 },
        { numId: '1', ilvl: 2 },
      ],
      defs,
    )
    expect(markers.map((m) => [m?.text, m?.number])).toEqual([
      ['Chapter 1.', '1'],
      ['Chapter 2.', '2'],
      ['2.1', '2.1'],
      ['(a)', 'a'],
    ])
  })
})

describe('caption numbering', () => {
  it('a caption with a chapter number is written as Word writes it', () => {
    const xml = generateCaptionXml('Figure', 2, 'Sales', {
      chapter: { style: 1, sep: 'hyphen' },
      chapterNumber: '3',
    })
    expect(seqFieldsOfXml(xml).map((f) => [f.instr, f.result, f.dirtyAttr !== null])).toEqual([
      ['STYLEREF 1 \\s', '3', false],
      ['SEQ Figure \\* ARABIC \\s 1', '2', false],
    ])
    expect(xml).toContain('<w:noBreakHyphen/>')
    expect(captionNumberingOfXml(xml)).toEqual({
      label: 'Figure',
      numbering: { format: 'ARABIC', chapter: { style: 1, sep: 'hyphen' } },
    })
    const colon = generateCaptionXml('My Table', 1, '', {
      format: 'ROMAN',
      chapter: { style: 2, sep: 'colon' },
    })
    expect(colon).toContain('<w:t>:</w:t>')
    expect(captionNumberingOfXml(colon, 'My Table')?.numbering).toEqual({
      format: 'ROMAN',
      chapter: { style: 2, sep: 'colon' },
    })
    expect(captionNumberingOfXml(colon, 'Figure')).toBeNull()
  })

  it('a caption takes a chapter number, changes it and loses it, its results kept', () => {
    const plain = generateCaptionXml('Figure', 4, 'Sales')
    const withChapter = setCaptionNumberingXml(plain, 'Figure', {
      format: 'alphabetic',
      chapter: { style: 1, sep: 'period' },
    })!
    expect(seqFieldsOfXml(withChapter).map((f) => [f.instr, f.result])).toEqual([
      ['STYLEREF 1 \\s', '0'],
      ['SEQ Figure \\* alphabetic \\s 1', '4'],
    ])
    expect(withChapter).toContain('<w:t xml:space="preserve"> Sales</w:t>')
    const dashed = setCaptionNumberingXml(rewriteSeqResultsXml(withChapter, ['2', 'd']), 'Figure', {
      format: 'alphabetic',
      chapter: { style: 2, sep: 'enDash' },
    })!
    expect(seqFieldsOfXml(dashed).map((f) => [f.instr, f.result])).toEqual([
      ['STYLEREF 2 \\s', '2'],
      ['SEQ Figure \\* alphabetic \\s 2', 'd'],
    ])
    const back = setCaptionNumberingXml(dashed, 'Figure', { format: 'ARABIC', chapter: null })!
    expect(back).toBe(plain.replace('<w:t>4</w:t>', '<w:t>d</w:t>'))
    // nothing to change, another label, or a number arranged by hand
    expect(setCaptionNumberingXml(plain, 'Figure', { format: 'ARABIC', chapter: null })).toBeNull()
    expect(setCaptionNumberingXml(plain, 'Table', { format: 'ROMAN', chapter: null })).toBeNull()
    const restarted = `<w:p>${field('SEQ Figure \\r 5', '5')}</w:p>`
    expect(
      setCaptionNumberingXml(restarted, 'Figure', { format: 'ROMAN', chapter: null }),
    ).toBeNull()
  })

  it('a caption as Word saves it, simple fields and proofing marks included, is rewritten whole', () => {
    const word =
      '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">Figure </w:t></w:r>' +
      `${simple('STYLEREF 1 \\s', '1')}<w:r><w:noBreakHyphen/></w:r>${simple('SEQ Figure \\* ARABIC \\s 1', '1')}` +
      '<w:proofErr w:type="gramStart"/><w:r><w:t>:The</w:t></w:r></w:p>'
    expect(captionNumberingOfXml(word)?.numbering).toEqual({
      format: 'ARABIC',
      chapter: { style: 1, sep: 'hyphen' },
    })
    const plain = setCaptionNumberingXml(word, 'Figure', { format: 'ROMAN', chapter: null })!
    expect(seqFieldsOfXml(plain).map((f) => [f.instr, f.result])).toEqual([
      ['SEQ Figure \\* ROMAN', '1'],
    ])
    expect(plain).toContain('<w:rPr><w:noProof/></w:rPr>')
    expect(plain).toContain('<w:t>:The</w:t>')
    expect(plain).not.toContain('noBreakHyphen')
  })
})
