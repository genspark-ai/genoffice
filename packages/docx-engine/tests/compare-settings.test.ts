import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compareDocuments, parseDocx, type CompareSettings } from '../src/index'
import { applyTrackedChangesXml } from '../src/revision-xml'
import { buildDocx, type BuildDocxOptions } from './helpers/build-docx'

const AUTHOR = 'Rae Visor'
const DATE = '2026-11-18T00:00:00Z'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const p = (text: string, rPr = '') =>
  `<w:p><w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

async function compare(
  original: BuildDocxOptions,
  revised: BuildDocxOptions,
  settings: Partial<CompareSettings> = {},
) {
  const result = await compareDocuments(
    await parseDocx(await buildDocx(original)),
    await parseDocx(await buildDocx(revised)),
    { author: AUTHOR, date: DATE, settings },
  )
  const zip = await JSZip.loadAsync(result.bytes)
  const part = async (path: string) => (await zip.file(path)?.async('string')) ?? ''
  return { ...result, part, xml: await part('word/document.xml') }
}

/** the text of a part with every tracked change accepted or rejected */
function resolved(xml: string, mode: 'accept' | 'reject'): string {
  const out = applyTrackedChangesXml(xml, mode)
  expect(out.remaining).toBe(0)
  return [...out.xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('')
}

describe('comparison settings: formatting', () => {
  const original = { bodyXml: p('Plain words stay.') + p('This turns bold.') }
  const revised = { bodyXml: p('Plain words stay.') + p('This turns bold.', '<w:b/>') }

  it('text whose formatting alone changed carries a w:rPrChange with the old formatting', async () => {
    const { xml, summary } = await compare(
      { bodyXml: p('Plain words stay.') + p('This was italic.', '<w:i/>') },
      { bodyXml: p('Plain words stay.') + p('This was italic.', '<w:b/>') },
    )
    expect(xml).toMatch(
      new RegExp(
        `<w:rPr><w:b/>.*?<w:rPrChange w:id="\\d+" w:author="${AUTHOR}" w:date="${DATE}"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t[^>]*>This was italic\\.</w:t>`,
      ),
    )
    expect(xml).not.toContain('<w:del ')
    expect(xml).not.toContain('<w:ins ')
    expect(summary).toMatchObject({ formatChanges: 1, insertions: 0, deletions: 0 })
    const reread = await parseDocx(
      (
        await compareDocuments(
          await parseDocx(await buildDocx(original)),
          await parseDocx(await buildDocx(revised)),
          { author: AUTHOR, date: DATE },
        )
      ).bytes,
    )
    const run = reread.blocks[1]!.runs![0]!
    expect(run.bold).toBe(true)
    expect(run.rPrChange?.author).toBe(AUTHOR)
    expect(run.rPrChange?.old?.bold).toBeFalsy()
  })

  it('switched off, a formatting difference is not a change', async () => {
    const { xml, summary } = await compare(original, revised, { formatting: false })
    expect(xml).not.toContain('w:rPrChange')
    expect(summary).toMatchObject({ formatChanges: 0, changedParagraphs: 0 })
  })
})

describe('comparison settings: moves', () => {
  const A = 'The first clause stands here.'
  const B = 'The second clause follows it.'
  const C = 'The third clause closes the list.'
  const original = { bodyXml: p(A) + p(B) + p(C) }
  const revised = { bodyXml: p(B) + p(C) + p(A) }

  it('a paragraph that sits elsewhere with the same text is one move', async () => {
    const { xml, summary } = await compare(original, revised)
    expect(summary).toMatchObject({ moves: 1, insertedParagraphs: 0, deletedParagraphs: 0 })
    const from =
      /<w:moveFromRangeStart [^>]*w:name="([^"]+)"\/><w:moveFrom [^>]*><w:r><w:t[^>]*>([^<]*)<\/w:t>/.exec(
        xml,
      )
    const to =
      /<w:moveToRangeStart [^>]*w:name="([^"]+)"\/><w:moveTo [^>]*><w:r><w:t[^>]*>([^<]*)<\/w:t>/.exec(
        xml,
      )
    expect(from?.[2]).toBe(A)
    expect(to?.[2]).toBe(A)
    expect(from?.[1]).toBe(to?.[1])
    expect(xml).not.toContain('<w:delText')
    expect(xml.indexOf('<w:moveFrom ')).toBeLessThan(xml.indexOf(B))
    expect(xml.indexOf('<w:moveTo ')).toBeGreaterThan(xml.indexOf(C))
    const reread = await parseDocx((await compare(original, revised)).bytes)
    expect(reread.blocks.slice(0, 4).map((b) => b.moveRevision ?? null)).toEqual([
      'from',
      null,
      null,
      'to',
    ])
  })

  it('switched off, the same paragraph is a deletion and an insertion', async () => {
    const { xml, summary } = await compare(original, revised, { moves: false })
    expect(xml).not.toContain('w:moveFrom')
    expect(xml).not.toContain('w:moveTo')
    expect(summary).toMatchObject({ moves: 0, insertedParagraphs: 1, deletedParagraphs: 1 })
    expect(xml).toContain(`<w:delText xml:space="preserve">${A}</w:delText>`)
  })
})

describe('comparison settings: tables', () => {
  const table = (...cells: string[]) =>
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>' +
    cells.map((text) => `<w:tc><w:tcPr/>${p(text)}</w:tc>`).join('') +
    '</w:tr></w:tbl>'
  const original = { bodyXml: p('Before the table.') + table('Price', 'ten dollars') }
  const revised = { bodyXml: p('Before the table.') + table('Price', 'twelve dollars') }

  it('a table that kept its shape is compared cell by cell', async () => {
    const { xml, summary } = await compare(original, revised)
    expect(xml).not.toMatch(/<w:(?:ins|del)\b[^>]*>\s*<w:tbl/)
    expect(xml.match(/<w:tbl>/g)).toHaveLength(1)
    expect(xml).toMatch(/<w:del [^>]*><w:r><w:delText[^>]*>ten<\/w:delText><\/w:r><\/w:del>/)
    expect(xml).toMatch(/<w:ins [^>]*><w:r><w:t[^>]*>twelve<\/w:t><\/w:r><\/w:ins>/)
    expect(summary).toMatchObject({ insertions: 1, deletions: 1 })
    expect(resolved(xml, 'accept')).toContain('twelve dollars')
    expect(resolved(xml, 'reject')).toContain('ten dollars')
  })

  it('switched off, the table is the revised one with nothing marked', async () => {
    const { xml, summary } = await compare(original, revised, { tables: false })
    expect(xml).not.toContain('<w:ins ')
    expect(xml).not.toContain('<w:del ')
    expect(xml).toContain('twelve dollars')
    expect(summary).toMatchObject({ insertions: 0, deletions: 0 })
  })
})

describe('comparison settings: headers and footers', () => {
  const body = p('The body is the same.')
  const original = { bodyXml: body, headerXml: p('Draft contract'), footerXml: p('Page footer') }
  const revised = { bodyXml: body, headerXml: p('Final contract'), footerXml: p('Page footer') }

  it('a changed header is marked in the header part, an unchanged footer is left alone', async () => {
    const { part, summary } = await compare(original, revised)
    const header = await part('word/header1.xml')
    expect(header).toMatch(
      new RegExp(
        `<w:del w:id="\\d+" w:author="${AUTHOR}" w:date="${DATE}"><w:r><w:delText[^>]*>Draft</w:delText>`,
      ),
    )
    expect(header).toMatch(/<w:ins [^>]*><w:r><w:t[^>]*>Final<\/w:t>/)
    expect(resolved(header, 'accept')).toBe('Final contract')
    expect(resolved(header, 'reject')).toBe('Draft contract')
    expect(await part('word/footer1.xml')).not.toMatch(/<w:(?:ins|del) /)
    expect(summary).toMatchObject({ insertions: 1, deletions: 1, changedParagraphs: 0 })
  })

  it('the redline reads back with the header changes as tracked runs', async () => {
    const { bytes } = await compare(original, revised)
    const reread = await parseDocx(bytes)
    const runs = reread.headerParas?.flatMap((para) => para.runs ?? []) ?? []
    expect(runs.find((r) => r.text === 'Draft')?.del?.author).toBe(AUTHOR)
    expect(runs.find((r) => r.text === 'Final')?.ins?.author).toBe(AUTHOR)
  })

  it('a paragraph only one side has is marked whole, on its runs and its paragraph mark', async () => {
    const { part } = await compare(
      { bodyXml: body, headerXml: p('Kept line') + p('Line that goes') },
      { bodyXml: body, headerXml: p('Kept line') + p('Line that comes') + p('Another new line') },
    )
    const header = await part('word/header1.xml')
    expect(header).toMatch(
      /<w:pPr><w:rPr><w:ins [^>]*\/><\/w:rPr><\/w:pPr><w:ins [^>]*><w:r><w:t[^>]*>Another new line/,
    )
    expect(resolved(header, 'accept')).toBe('Kept lineLine that comesAnother new line')
    expect(resolved(header, 'reject')).toBe('Kept lineLine that goes')
  })

  it('a paragraph with a field that differs is left as the revised document has it and counted', async () => {
    const field = (text: string) =>
      `<w:p><w:r><w:t xml:space="preserve">${text} </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>`
    const { part, summary } = await compare(
      { bodyXml: body, footerXml: field('Page') },
      { bodyXml: body, footerXml: field('Sheet') },
    )
    expect(await part('word/footer1.xml')).not.toMatch(/<w:(?:ins|del) /)
    expect(summary.skipped).toBe(1)
  })

  it('switched off, the header part is the revised one byte for byte', async () => {
    const { part, summary } = await compare(original, revised, { headersFooters: false })
    const revisedZip = await JSZip.loadAsync(await buildDocx(revised))
    expect(await part('word/header1.xml')).toBe(
      await revisedZip.file('word/header1.xml')!.async('string'),
    )
    expect(summary).toMatchObject({ insertions: 0, deletions: 0 })
  })
})

const NOTE_REF = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
const footnotes = (...texts: string[]) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W}>` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  texts
    .map(
      (text, i) =>
        `<w:footnote w:id="${i + 1}"><w:p>${NOTE_REF}<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:footnote>`,
    )
    .join('') +
  '</w:footnotes>'
const noted = (count: number, ...texts: string[]): BuildDocxOptions => ({
  bodyXml:
    '<w:p><w:r><w:t>Body with notes.</w:t></w:r>' +
    Array.from(
      { length: count },
      (_, i) => `<w:r><w:footnoteReference w:id="${i + 1}"/></w:r>`,
    ).join('') +
    '</w:p>',
  extraRels:
    '<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
  extraParts: [
    {
      path: 'word/footnotes.xml',
      xml: footnotes(...texts),
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
    },
  ],
})

describe('comparison settings: footnotes and endnotes', () => {
  const original = noted(1, 'See the first edition.')
  const revised = noted(1, 'See the second edition.')

  it('a changed footnote is marked in the notes part and keeps its reference mark', async () => {
    const { part, bytes } = await compare(original, revised)
    const notes = await part('word/footnotes.xml')
    expect(notes).toMatch(/<w:del [^>]*><w:r><w:delText[^>]*>first<\/w:delText>/)
    expect(notes).toMatch(/<w:ins [^>]*><w:r><w:t[^>]*>second<\/w:t>/)
    expect(notes.match(/<w:footnoteRef\/>/g)).toHaveLength(1)
    expect(notes).toContain('<w:footnote w:type="separator" w:id="-1">')
    expect(resolved(notes, 'accept')).toBe(' See the second edition.')
    expect(resolved(notes, 'reject')).toBe(' See the first edition.')
    const note = (await parseDocx(bytes)).footnotes[0]!
    const runs = note.richParas?.flat() ?? []
    expect(runs.find((r) => r.text === 'first')?.del?.author).toBe(AUTHOR)
    expect(runs.find((r) => r.text === 'second')?.ins?.author).toBe(AUTHOR)
  })

  it('a footnote only the revised document has is marked inserted', async () => {
    const { part } = await compare(
      noted(1, 'See the first edition.'),
      noted(2, 'See the first edition.', 'A note added later.'),
    )
    const notes = await part('word/footnotes.xml')
    expect(notes).toMatch(
      /<w:footnote w:id="2"><w:p><w:pPr><w:rPr><w:ins [^>]*\/><\/w:rPr><\/w:pPr><w:ins /,
    )
    expect(notes).not.toContain('<w:del ')
    expect(resolved(notes, 'accept')).toContain('A note added later.')
    expect(resolved(notes, 'reject')).not.toContain('A note added later.')
  })

  it('switched off, the notes part is the revised one byte for byte', async () => {
    const { part } = await compare(original, revised, { notes: false })
    expect(await part('word/footnotes.xml')).toBe(footnotes('See the second edition.'))
  })
})

describe('comparison settings: comments', () => {
  const comments =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W}>` +
    '<w:comment w:id="0" w:author="Reviewer" w:date="2026-11-01T00:00:00Z"><w:p><w:r><w:t>Check this figure.</w:t></w:r></w:p></w:comment>' +
    '</w:comments>'
  const original = { bodyXml: p('The fee is ten dollars.') }
  const revised: BuildDocxOptions = {
    bodyXml:
      '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">The fee is twelve dollars.</w:t></w:r>' +
      '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>',
    extraRels:
      '<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
    extraParts: [
      {
        path: 'word/comments.xml',
        xml: comments,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
      },
    ],
  }

  it('the result carries the comments of the revised document', async () => {
    const { bytes, xml } = await compare(original, revised)
    expect(xml).toContain('<w:commentReference w:id="0"/>')
    const reread = await parseDocx(bytes)
    expect(reread.comments.map((c) => c.text)).toEqual(['Check this figure.'])
    expect(xml).toMatch(/<w:ins [^>]*><w:r><w:t[^>]*>twelve<\/w:t>/)
  })

  it('switched off, the result has the text changes and no comment', async () => {
    const { bytes, xml } = await compare(original, revised, { comments: false })
    expect(xml).not.toContain('w:commentReference')
    expect(xml).not.toContain('w:commentRangeStart')
    expect((await parseDocx(bytes)).comments).toEqual([])
    expect(xml).toMatch(/<w:ins [^>]*><w:r><w:t[^>]*>twelve<\/w:t>/)
  })
})

describe('comparison settings: case and white space', () => {
  it('a paragraph that differs mostly in case still pairs with its twin once case is ignored', async () => {
    const { xml, summary } = await compare(
      { bodyXml: p('Kept as it is.') + p('THE SELLER SHALL DELIVER THE GOODS ON MONDAY.') },
      { bodyXml: p('Kept as it is.') + p('the seller shall deliver the goods on friday.') },
      { caseChanges: false },
    )
    expect(summary).toMatchObject({
      changedParagraphs: 1,
      insertedParagraphs: 0,
      deletedParagraphs: 0,
    })
    expect(xml).toMatch(/<w:delText[^>]*>MONDAY\.<\/w:delText>/)
    expect(xml).toMatch(/<w:ins [^>]*><w:r><w:t[^>]*>friday\.<\/w:t>/)
  })

  it('a table replaced whole counts as an insertion and a deletion', async () => {
    const row = (text: string) => `<w:tr><w:tc><w:tcPr/>${p(text)}</w:tc></w:tr>`
    const table = (rows: string) =>
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl>`
    const { summary } = await compare(
      { bodyXml: p('Same.') + table(row('one')) },
      { bodyXml: p('Same.') + table(row('one') + row('two')) },
    )
    expect(summary).toMatchObject({ insertions: 1, deletions: 1 })
  })

  const original = { bodyXml: p('The Seller  shall deliver.') }
  const revised = { bodyXml: p('The seller shall deliver.') }

  it('both differences are changes unless switched off', async () => {
    expect((await compare(original, revised)).summary.changedParagraphs).toBe(1)
    expect(
      (await compare(original, revised, { caseChanges: false })).summary.changedParagraphs,
    ).toBe(1)
    const off = await compare(original, revised, { caseChanges: false, whiteSpace: false })
    expect(off.summary).toMatchObject({ changedParagraphs: 0, insertions: 0, deletions: 0 })
    expect(off.xml).not.toMatch(/<w:(?:ins|del) /)
  })
})

describe('comparison settings: text boxes', () => {
  const box = (text: string) =>
    `<w:p><w:r><w:pict><v:shape id="b1" type="#_x0000_t202" style="position:absolute;width:200pt;height:60pt"><v:textbox><w:txbxContent>${p('Box heading.')}${p(text)}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`
  const original = { bodyXml: p('Body stays.') + box('The fee is ten dollars.') }
  const revised = { bodyXml: p('Body stays.') + box('The fee is twelve dollars.') }

  it('the text inside a text box is compared in place', async () => {
    const { xml, summary } = await compare(original, revised)
    const inner = /<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/.exec(xml)?.[1] ?? ''
    expect(inner).toMatch(/<w:del [^>]*><w:r><w:delText[^>]*>ten<\/w:delText>/)
    expect(inner).toMatch(/<w:ins [^>]*><w:r><w:t[^>]*>twelve<\/w:t>/)
    expect(xml.match(/<v:shape /g)).toHaveLength(1)
    expect(summary).toMatchObject({ insertions: 1, deletions: 1, changedParagraphs: 1 })
    expect(resolved(inner, 'accept')).toBe('Box heading.The fee is twelve dollars.')
    expect(resolved(inner, 'reject')).toBe('Box heading.The fee is ten dollars.')
  })

  it('switched off, a text box stays as the revised document has it', async () => {
    const { xml, summary } = await compare(original, revised, { textBoxes: false })
    expect(xml).not.toMatch(/<w:(?:ins|del) /)
    expect(xml).toContain('The fee is twelve dollars.')
    expect(summary).toMatchObject({ insertions: 0, deletions: 0, changedParagraphs: 0 })
  })
})

describe('comparison settings: fields', () => {
  const field = (code: string, shown: string) =>
    `<w:p><w:r><w:t xml:space="preserve">Printed on </w:t></w:r><w:fldSimple w:instr=" ${code} "><w:r><w:t>${shown}</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> by the clerk.</w:t></w:r></w:p>`

  it('a field whose code changed is a change though it shows the same text', async () => {
    const { xml, summary } = await compare(
      { bodyXml: p('Same.') + field('DATE \\@ "yyyy"', '2026') },
      { bodyXml: p('Same.') + field('SAVEDATE \\@ "yyyy"', '2026') },
    )
    expect(xml).toMatch(/<w:del [^>]*><w:r><w:delText[^>]*>2026<\/w:delText>/)
    expect(xml).toMatch(/<w:ins [^>]*>[\s\S]*?SAVEDATE[\s\S]*?<\/w:ins>/)
    expect(summary).toMatchObject({ insertions: 1, deletions: 1 })
  })

  it('switched off, neither the code nor the shown text of a field is a change', async () => {
    const code = await compare(
      { bodyXml: p('Same.') + field('DATE \\@ "yyyy"', '2026') },
      { bodyXml: p('Same.') + field('SAVEDATE \\@ "yyyy"', '2026') },
      { fields: false },
    )
    expect(code.xml).not.toMatch(/<w:(?:ins|del) /)
    const shown = await compare(
      { bodyXml: p('Same.') + field('DATE \\@ "yyyy"', '2025') },
      { bodyXml: p('Same.') + field('DATE \\@ "yyyy"', '2026') },
      { fields: false },
    )
    expect(shown.xml).not.toMatch(/<w:(?:ins|del) /)
    expect(shown.xml).toContain('2026')
    expect(shown.summary).toMatchObject({ insertions: 0, deletions: 0, changedParagraphs: 0 })
  })

  it('switched off, text outside the field is still compared', async () => {
    const { xml } = await compare(
      { bodyXml: field('DATE \\@ "yyyy"', '2025') },
      {
        bodyXml: field('DATE \\@ "yyyy"', '2026').replace('by the clerk', 'by the registrar'),
      },
      { fields: false },
    )
    expect(xml).toMatch(/<w:delText[^>]*>clerk\.<\/w:delText>/)
    expect(xml).not.toMatch(/<w:delText[^>]*>2025/)
    expect(xml).toContain('2026')
  })

  it('the shown text of a field that changed is a change while fields are compared', async () => {
    const { xml } = await compare(
      { bodyXml: field('DATE \\@ "yyyy"', '2025') },
      { bodyXml: field('DATE \\@ "yyyy"', '2026') },
    )
    expect(xml).toMatch(/<w:delText[^>]*>2025<\/w:delText>/)
  })
})
