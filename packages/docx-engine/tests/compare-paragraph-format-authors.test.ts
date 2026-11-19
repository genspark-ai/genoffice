import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compareDocuments, parseDocx, type CompareOptions } from '../src/index'
import { applyTrackedChangesXml } from '../src/revision-xml'
import { buildDocx, type BuildDocxOptions } from './helpers/build-docx'

const LABEL = 'Rae Visor'
const DATE = '2026-11-18T00:00:00Z'
const ABE = 'w:author="Abe Origin" w:date="2026-11-01T08:00:00Z"'
const BEA = 'w:author="Bea Revise" w:date="2026-11-02T09:00:00Z"'

const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const gone = (text: string) => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`
const p = (content: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content.startsWith('<') ? content : r(content)}</w:p>`

async function compare(
  original: BuildDocxOptions,
  revised: BuildDocxOptions,
  options: CompareOptions = {},
) {
  const result = await compareDocuments(
    await parseDocx(await buildDocx(original)),
    await parseDocx(await buildDocx(revised)),
    { author: LABEL, date: DATE, ...options },
  )
  const zip = await JSZip.loadAsync(result.bytes)
  const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
  return { ...result, xml, body: /<w:body>([\s\S]*)<w:sectPr/.exec(xml)?.[1] ?? '' }
}

const acceptAll = (xml: string) => resolvedXml(xml, 'accept')
const rejectAll = (xml: string) => resolvedXml(xml, 'reject')
function resolvedXml(xml: string, mode: 'accept' | 'reject'): string {
  const out = applyTrackedChangesXml(xml, mode)
  expect(out.remaining).toBe(0)
  return out.xml
}
const textOf = (xml: string): string[] =>
  [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((m) =>
    [...m[0].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''),
  )
/** each paragraph's own w:pPr; one left empty reads as none */
const pPrs = (xml: string): string[] =>
  [...xml.matchAll(/<w:p>(<w:pPr>[\s\S]*?<\/w:pPr>(?=<w:r|<w:ins|<w:del|<\/w:p>))?/g)].map((m) =>
    (m[1] ?? '').replace(/^<w:pPr><\/w:pPr>$/, ''),
  )
const RECORD = `<w:pPrChange w:id="\\d+" w:author="${LABEL}" w:date="${DATE}">`

describe('comparison: paragraph formatting', () => {
  const STAY = p('This paragraph stays as it is.')

  it('a paragraph whose alignment alone changed carries a w:pPrChange with the original w:pPr', async () => {
    const { body, summary, bytes } = await compare(
      { bodyXml: STAY + p('Same words here.', '<w:jc w:val="left"/>') },
      { bodyXml: STAY + p('Same words here.', '<w:jc w:val="center"/>') },
    )
    expect(body).toMatch(
      new RegExp(
        `<w:p><w:pPr><w:jc w:val="center"/>${RECORD}<w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t[^>]*>Same words here\\.</w:t></w:r></w:p>`,
      ),
    )
    expect(body).not.toMatch(/<w:(?:ins|del) /)
    expect(summary).toMatchObject({
      formatChanges: 1,
      insertions: 0,
      deletions: 0,
      changedParagraphs: 1,
      unchangedParagraphs: 1,
    })
    expect(pPrs(acceptAll(body))).toEqual(['', '<w:pPr><w:jc w:val="center"/></w:pPr>'])
    expect(pPrs(rejectAll(body))).toEqual(['', '<w:pPr><w:jc w:val="left"/></w:pPr>'])
    const block = (await parseDocx(bytes)).blocks[1]!
    expect(block.format?.align).toBe('center')
    expect(block.pPrChangeInfo).toMatchObject({ author: LABEL, date: DATE, old: { align: 'left' } })
  })

  it('formatting the original did not have is recorded against an empty w:pPr', async () => {
    const added =
      '<w:keepNext/><w:pageBreakBefore/><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>' +
      '<w:spacing w:before="240" w:after="120"/><w:ind w:left="720" w:hanging="360"/>'
    const { body, summary } = await compare(
      { bodyXml: STAY + p('Same words here.') },
      { bodyXml: STAY + p('Same words here.', added) },
    )
    expect(body).toMatch(
      new RegExp(`<w:pPr>${added}${RECORD}<w:pPr></w:pPr></w:pPrChange></w:pPr>`),
    )
    expect(summary.formatChanges).toBe(1)
    expect(pPrs(acceptAll(body))[1]).toBe(`<w:pPr>${added}</w:pPr>`)
    expect(pPrs(rejectAll(body))[1]).toBe('')
  })

  it('formatting the revised document dropped comes back on reject', async () => {
    const had = '<w:keepLines/><w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="both"/>'
    const { body } = await compare(
      { bodyXml: STAY + p('Same words here.', had) },
      { bodyXml: STAY + p('Same words here.') },
    )
    expect(body).toMatch(
      new RegExp(`<w:p><w:pPr>${RECORD}<w:pPr>${had}</w:pPr></w:pPrChange></w:pPr>`),
    )
    expect(pPrs(acceptAll(body))[1]).toBe('')
    expect(pPrs(rejectAll(body))[1]).toBe(`<w:pPr>${had}</w:pPr>`)
  })

  it('a style and a list are paragraph formatting', async () => {
    const list = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'
    const { body, summary, bytes } = await compare(
      {
        withNumbering: true,
        bodyXml: STAY + p('A heading once.', '<w:pStyle w:val="Heading1"/>') + p('An item.', list),
      },
      {
        withNumbering: true,
        bodyXml:
          STAY +
          p('A heading once.') +
          p('An item.', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
      },
    )
    expect(summary.formatChanges).toBe(2)
    expect(body).toMatch(
      new RegExp(`${RECORD}<w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange>`),
    )
    expect(body).toMatch(new RegExp(`${RECORD}<w:pPr>${list}</w:pPr></w:pPrChange>`))
    expect(pPrs(rejectAll(body)).slice(1)).toEqual([
      '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>',
      `<w:pPr>${list}</w:pPr>`,
    ])
    const blocks = (await parseDocx(bytes)).blocks
    expect(blocks[1]!.pPrChangeInfo?.old).toMatchObject({ type: 'docHeading', level: 1 })
    expect(blocks[2]!.pPrChangeInfo?.old).toMatchObject({ type: 'docListItem', numId: '2' })
  })

  it('a style only the original defines is left out of the record', async () => {
    const { body } = await compare(
      {
        extraStylesXml:
          '<w:style w:type="paragraph" w:styleId="Gone"><w:name w:val="Gone"/></w:style>',
        bodyXml: STAY + p('Same words here.', '<w:pStyle w:val="Gone"/><w:jc w:val="right"/>'),
      },
      { bodyXml: STAY + p('Same words here.') },
    )
    expect(body).toMatch(new RegExp(`${RECORD}<w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>`))
  })

  it('a paragraph whose text and formatting both changed carries both', async () => {
    const { body, summary } = await compare(
      { bodyXml: STAY + p('The price is ten dollars today.') },
      { bodyXml: STAY + p('The price is twelve dollars today.', '<w:jc w:val="right"/>') },
    )
    expect(summary).toMatchObject({ formatChanges: 1, insertions: 1, deletions: 1 })
    expect(body).toMatch(
      new RegExp(`<w:pPr><w:jc w:val="right"/>${RECORD}<w:pPr></w:pPr></w:pPrChange></w:pPr>`),
    )
    expect(textOf(acceptAll(body))[1]).toBe('The price is twelve dollars today.')
    expect(pPrs(rejectAll(body))[1]).toBe('')
    expect(textOf(rejectAll(body))[1]).toBe('The price is ten dollars today.')
  })

  it('the spelling of the same properties, the paragraph mark and a section are not a change', async () => {
    const { body, summary } = await compare(
      {
        bodyXml:
          STAY +
          p('Same words here.', '<w:keepNext w:val="1"/><w:spacing w:after="120" w:before="240"/>'),
      },
      {
        bodyXml:
          STAY +
          p(
            'Same words here.',
            '<w:keepNext/><w:spacing w:before="240" w:after="120"/><w:rPr><w:b/></w:rPr>',
          ),
      },
    )
    expect(body).not.toContain('w:pPrChange')
    expect(summary).toMatchObject({ formatChanges: 0, changedParagraphs: 0 })
  })

  it('the Formatting setting switched off, a paragraph formatting difference is not a change', async () => {
    const { body, summary } = await compare(
      { bodyXml: STAY + p('Same words here.') },
      { bodyXml: STAY + p('Same words here.', '<w:jc w:val="center"/>') },
      { settings: { formatting: false } },
    )
    expect(body).not.toContain('w:pPrChange')
    expect(body).toContain('<w:jc w:val="center"/>')
    expect(summary).toMatchObject({ formatChanges: 0, changedParagraphs: 0 })
  })

  it('a record the revised document already holds is replaced by the comparison with the original', async () => {
    const { body } = await compare(
      { bodyXml: STAY + p('Same words here.', '<w:jc w:val="left"/>') },
      {
        bodyXml:
          STAY +
          p(
            'Same words here.',
            `<w:jc w:val="center"/><w:pPrChange w:id="7" ${BEA}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>`,
          ),
      },
    )
    expect(body.match(/<w:pPrChange /g)).toHaveLength(1)
    expect(body).toMatch(new RegExp(`${RECORD}<w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>`))
  })
})

describe('combining: tracked changes keep their authors', () => {
  const STAY = p('This paragraph stays as it is.')
  const combine = (original: string, revised: string) =>
    compare({ bodyXml: STAY + original }, { bodyXml: STAY + revised }, { keepRevisions: true })
  const authorsOf = (body: string) =>
    [
      ...body.matchAll(
        /<w:(ins|del) [^>]*w:author="([^"]+)"[^>]*>(?:<w:(?:ins|del) [^>]*>)?<w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:(?:t|delText)[^>]*>([^<]*)</g,
      ),
    ].map((m) => `${m[1]} ${m[2]}: ${m[3]}`)

  it('a change the revised document tracks stays under its author; only the rest is new', async () => {
    const { body, summary } = await combine(
      p('The cat sat on the mat.'),
      p(
        r('The ') +
          `<w:ins w:id="3" ${BEA}>${r('black ')}</w:ins>` +
          r('cat sat on the ') +
          `<w:del w:id="4" ${BEA}>${gone('mat')}</w:del>` +
          r('rug.'),
      ),
    )
    expect(authorsOf(body)).toEqual([
      'ins Bea Revise: black ',
      'del Bea Revise: mat',
      `ins ${LABEL}: rug`,
    ])
    expect(body).toContain('w:date="2026-11-02T09:00:00Z"')
    expect(summary).toMatchObject({ insertions: 1, deletions: 0, carried: 0 })
    expect(textOf(acceptAll(body))[1]).toBe('The black cat sat on the rug.')
    expect(textOf(rejectAll(body))[1]).toBe('The cat sat on the mat.')
  })

  it('compared instead, the same pair reads the pending changes as accepted', async () => {
    const { body } = await compare(
      { bodyXml: STAY + p('The cat sat on the mat.') },
      {
        bodyXml:
          STAY +
          p(
            r('The ') +
              `<w:ins w:id="3" ${BEA}>${r('black ')}</w:ins>` +
              r('cat sat on the ') +
              `<w:del w:id="4" ${BEA}>${gone('mat')}</w:del>` +
              r('rug.'),
          ),
      },
    )
    expect(body).not.toContain('Bea Revise')
    expect(textOf(acceptAll(body))[1]).toBe('The black cat sat on the rug.')
    expect(textOf(rejectAll(body))[1]).toBe('The cat sat on the mat.')
  })

  it('changes the original tracks are carried with their author and date', async () => {
    const { body, summary } = await combine(
      p(
        r('The ') +
          `<w:ins w:id="5" ${ABE}>${r('old ')}</w:ins>` +
          r('cat sat ') +
          `<w:del w:id="6" ${ABE}>${gone('quietly ')}</w:del>` +
          r('on the mat.'),
      ),
      p('The cat sat quietly on the mat today.'),
    )
    expect(authorsOf(body)).toEqual([
      'ins Abe Origin: old ',
      'del Abe Origin: quietly',
      `del ${LABEL}: mat.`,
      `ins ${LABEL}: mat today.`,
    ])
    expect(body).toContain('w:date="2026-11-01T08:00:00Z"')
    expect(summary.carried).toBe(2)
    expect(textOf(acceptAll(body))[1]).toBe('The old cat sat on the mat today.')
    expect(textOf(rejectAll(body))[1]).toBe('The cat sat quietly on the mat.')
    const ids = [...body.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('a tracked change inside a word is told apart from the word', async () => {
    const original = p(r('Two cat') + `<w:ins w:id="5" ${ABE}>${r('s')}</w:ins>` + r(' sat here.'))
    const same = await combine(original, p('Two cats sat here.'))
    expect(authorsOf(same.body)).toEqual(['ins Abe Origin: s'])
    expect(textOf(acceptAll(same.body))[1]).toBe('Two cats sat here.')
    expect(textOf(rejectAll(same.body))[1]).toBe('Two cat sat here.')

    const without = await combine(original, p('Two cat sat here.'))
    expect(authorsOf(without.body)).toEqual(['ins Abe Origin: s'])
    expect(without.summary).toMatchObject({ insertions: 0, deletions: 0, carried: 1 })
    expect(textOf(acceptAll(without.body))[1]).toBe('Two cats sat here.')
  })

  it('a paragraph the original inserted as a tracked change is carried, not deleted', async () => {
    const { body, summary } = await combine(
      p('First clause of the contract.') +
        `<w:p><w:pPr><w:rPr><w:ins w:id="8" ${ABE}/></w:rPr></w:pPr><w:ins w:id="9" ${ABE}>${r('A clause Abe proposed.')}</w:ins></w:p>` +
        p('Last clause of the contract.'),
      p('First clause of the contract.') + p('Last clause of the contract.'),
    )
    expect(authorsOf(body)).toEqual(['ins Abe Origin: A clause Abe proposed.'])
    expect(body).toMatch(
      /<w:pPr><w:rPr><w:ins [^>]*w:author="Abe Origin"[^>]*\/><\/w:rPr><\/w:pPr>/,
    )
    expect(summary).toMatchObject({ deletedParagraphs: 0, deletions: 0 })
    expect(body).not.toMatch(/w:id="[89]"/)
    expect(textOf(acceptAll(body))).toContain('A clause Abe proposed.')
    expect(textOf(rejectAll(body))).toEqual([
      'This paragraph stays as it is.',
      'First clause of the contract.',
      'Last clause of the contract.',
    ])
  })

  it('a paragraph only one side has is new, except the text it already tracks', async () => {
    const { body } = await combine(
      p('First clause of the contract.') +
        p(r('A clause with ') + `<w:del w:id="5" ${ABE}>${gone('struck ')}</w:del>` + r('words.')) +
        p('Last clause of the contract.'),
      p('First clause of the contract.') +
        p(r('An added note ') + `<w:ins w:id="3" ${BEA}>${r('by Bea')}</w:ins>`) +
        p('Last clause of the contract.'),
    )
    expect(authorsOf(body).sort()).toEqual([
      'del Abe Origin: struck ',
      `del ${LABEL}: A clause with `,
      `del ${LABEL}: words.`,
      'ins Bea Revise: by Bea',
      `ins ${LABEL}: An added note `,
    ])
    expect(textOf(acceptAll(body))).toEqual([
      'This paragraph stays as it is.',
      'First clause of the contract.',
      'An added note by Bea',
      'Last clause of the contract.',
    ])
    expect(textOf(rejectAll(body))).toEqual([
      'This paragraph stays as it is.',
      'First clause of the contract.',
      'A clause with struck words.',
      'Last clause of the contract.',
    ])
  })

  it('paragraph formatting each side tracks stays under its author', async () => {
    const byBea = `<w:jc w:val="center"/><w:pPrChange w:id="7" ${BEA}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>`
    const byAbe = `<w:ind w:left="720"/><w:pPrChange w:id="8" ${ABE}><w:pPr><w:ind w:left="0"/></w:pPr></w:pPrChange>`
    const { body, summary } = await combine(
      p('Centred by Bea.', '<w:jc w:val="left"/>') + p('Indented by Abe.', byAbe),
      p('Centred by Bea.', byBea) + p('Indented by Abe.', '<w:ind w:left="720"/>'),
    )
    expect(body).toContain(`<w:pPr>${byBea}</w:pPr>`)
    expect(body).toMatch(
      new RegExp(
        `<w:pPr><w:ind w:left="720"/><w:pPrChange w:id="\\d+" ${ABE}><w:pPr><w:ind w:left="0"/></w:pPr></w:pPrChange></w:pPr>`,
      ),
    )
    expect(summary).toMatchObject({ formatChanges: 0, carried: 1 })
  })

  it('a header and a table cell keep the changes both sides track', async () => {
    const cell = (content: string) =>
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>' +
      p(content) +
      '</w:tc></w:tr></w:tbl>'
    const result = await compareDocuments(
      await parseDocx(
        await buildDocx({
          bodyXml:
            STAY +
            cell(r('Price: ten ') + `<w:ins w:id="5" ${ABE}>${r('US ')}</w:ins>` + r('dollars')),
          headerXml: p(
            r('Draft ') + `<w:del w:id="6" ${ABE}>${gone('internal ')}</w:del>` + r('contract'),
          ),
        }),
      ),
      await parseDocx(
        await buildDocx({
          bodyXml:
            STAY +
            cell(
              r('Price: ') + `<w:del w:id="3" ${BEA}>${gone('ten')}</w:del>` + r('twelve dollars'),
            ),
          headerXml: p(r('Draft internal contract ') + `<w:ins w:id="4" ${BEA}>${r('v2')}</w:ins>`),
        }),
      ),
      { author: LABEL, date: DATE, keepRevisions: true },
    )
    const zip = await JSZip.loadAsync(result.bytes)
    const header = (await zip.file('word/header1.xml')?.async('string')) ?? ''
    const body = (await zip.file('word/document.xml')?.async('string')) ?? ''
    expect(authorsOf(header)).toEqual([
      'del Abe Origin: internal',
      `ins ${LABEL}:  `,
      'ins Bea Revise: v2',
    ])
    expect(textOf(acceptAll(header))).toEqual(['Draft contract v2'])
    expect(textOf(rejectAll(header))).toEqual(['Draft internal contract'])
    expect(authorsOf(body)).toEqual([
      'del Bea Revise: ten',
      `ins ${LABEL}: twelve`,
      'ins Abe Origin: US ',
    ])
    expect(textOf(acceptAll(body))[1]).toBe('Price: twelve US dollars')
    expect(textOf(rejectAll(body))[1]).toBe('Price: ten dollars')
    for (const xml of [header, body]) {
      const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1])
      expect(new Set(ids).size).toBe(ids.length)
    }
  })

  it('an empty paragraph that is only a tracked paragraph mark stays under its author', async () => {
    const mark = (who: string) =>
      `<w:p><w:pPr><w:rPr><w:ins w:id="8" ${who}/></w:rPr></w:pPr></w:p>`
    const result = await compareDocuments(
      await parseDocx(
        await buildDocx({
          bodyXml: STAY + p('First clause.') + mark(ABE) + p('Last clause.'),
          headerXml: p('Draft contract') + mark(ABE) + p('Second line'),
        }),
      ),
      await parseDocx(
        await buildDocx({
          bodyXml: STAY + p('First clause.') + p('Last clause.') + mark(BEA) + p('The end.'),
          headerXml: p('Draft contract') + p('Second line') + mark(BEA) + p('Third line'),
        }),
      ),
      { author: LABEL, date: DATE, keepRevisions: true },
    )
    // the original's two marks (body and header) are what the status line counts
    expect(result.summary.carried).toBe(2)
    const zip = await JSZip.loadAsync(result.bytes)
    for (const part of ['word/document.xml', 'word/header1.xml']) {
      const xml = (await zip.file(part)?.async('string')) ?? ''
      const marks = [
        ...xml.matchAll(
          /<w:p><w:pPr><w:rPr>((?:<w:(?:ins|del) [^>]*\/>)+)<\/w:rPr><\/w:pPr><\/w:p>/g,
        ),
      ]
      expect(marks.map((m) => /w:author="([^"]+)"/.exec(m[1]!)?.[1]).sort()).toEqual([
        'Abe Origin',
        'Bea Revise',
      ])
      expect(marks.every((m) => (m[1]!.match(/<w:/g) ?? []).length === 1)).toBe(true)
      const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1])
      expect(new Set(ids).size).toBe(ids.length)
    }
  })

  it('nothing either side tracks, combining gives what comparing gives', async () => {
    const pair = [
      { bodyXml: STAY + p('The price is ten dollars.') + p('A clause that goes.') },
      { bodyXml: STAY + p('The price is twelve dollars.', '<w:jc w:val="right"/>') },
    ] as const
    const compared = await compare(pair[0], pair[1])
    const combined = await compare(pair[0], pair[1], { keepRevisions: true })
    expect(combined.body).toBe(compared.body)
  })
})
