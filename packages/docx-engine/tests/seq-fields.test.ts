import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  SeqCounter,
  generateCaptionXml,
  generateFieldTableXml,
  generateTableOfAuthoritiesFieldXml,
  generateTableOfFiguresFieldXml,
  generateTocFieldXml,
  parseDocx,
  parseSeqInstruction,
  rewriteSeqResultsXml,
  saveDocx,
  seqFieldsOfXml,
  unreliableSeqLabels,
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
    expect(parseSeqInstruction('SEQ Figure \\s 1')?.unsupported).toBe(true)
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
    expect(next('SEQ Table \\* Ordinal')).toEqual({ value: 2, text: null })
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
