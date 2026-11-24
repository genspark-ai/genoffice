import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compareDocuments, parseDocx, type CompareOptions } from '../src/index'
import { applyTrackedChangesXml } from '../src/revision-xml'
import { buildDocx, TINY_PNG_BASE64, type BuildDocxOptions } from './helpers/build-docx'

const LABEL = 'Rae Visor'
const DATE = '2026-11-22T00:00:00Z'
const ABE = 'w:author="Abe Origin" w:date="2026-11-01T08:00:00Z"'
const OTHER_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPj/HwADBwIAMCbHYQAAAABJRU5ErkJggg=='

const r = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
const gone = (text: string) => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`
const p = (content: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content.startsWith('<') ? content : r(content)}</w:p>`
const tc = (content: string, tcPr = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${tcPr}</w:tcPr>${content.startsWith('<') ? content : p(content)}</w:tc>`
const tr = (cells: string[], trPr = '') =>
  `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells.map((c) => (c.startsWith('<w:tc>') ? c : tc(c))).join('')}</w:tr>`
const tbl = (rows: string[], cols = 2) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(cols)}</w:tblGrid>${rows.join('')}</w:tbl>`
const picture =
  '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
const math = (text: string) =>
  `<m:oMath>${text
    .split(' ')
    .map((t) => `<m:r><m:t>${t}</m:t></m:r>`)
    .join('')}</m:oMath>`

const INTRO = p('This agreement is made between the supplier and the customer.')
const SIGNED = p('Signed by both parties.')

async function withPng(bytes: Uint8Array, png: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  zip.file('word/media/image1.png', png, { base64: true })
  return zip.generateAsync({ type: 'uint8array' })
}

async function compare(
  original: BuildDocxOptions & { png?: string },
  revised: BuildDocxOptions & { png?: string },
  options: CompareOptions = {},
) {
  const bytesOf = async (side: BuildDocxOptions & { png?: string }) =>
    side.png ? withPng(await buildDocx(side), side.png) : buildDocx(side)
  const result = await compareDocuments(
    await parseDocx(await bytesOf(original)),
    await parseDocx(await bytesOf(revised)),
    { author: LABEL, date: DATE, ...options },
  )
  const zip = await JSZip.loadAsync(result.bytes)
  const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
  return {
    ...result,
    zip,
    body: /<w:body>([\s\S]*)<w:sectPr/.exec(xml)?.[1] ?? '',
    header: (await zip.file('word/header1.xml')?.async('string')) ?? '',
  }
}

function resolved(xml: string, mode: 'accept' | 'reject'): string {
  const out = applyTrackedChangesXml(xml, mode)
  expect(out.remaining).toBe(0)
  return out.xml
}
/** the text of each paragraph that has any, an equation by its text and a picture as [pic] */
const texts = (xml: string): string[] =>
  [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)]
    .map((m) =>
      [...m[0].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<m:t>([^<]*)<\/m:t>|(<a:blip )/g)]
        .map((t) => t[1] ?? t[2] ?? '[pic]')
        .join(''),
    )
    .filter(Boolean)
const count = (xml: string, re: RegExp): number => (xml.match(re) ?? []).length

describe('comparison: paragraph formatting outside the top-level body paragraphs', () => {
  const table = (pPr: string) =>
    tbl([tr(['Item', tc(p('Price', pPr))]), tr(['Bolt', 'ten dollars'])])

  it('a cell paragraph whose alignment alone changed carries a w:pPrChange with the original w:pPr', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + table('<w:jc w:val="left"/>') + SIGNED },
      { bodyXml: INTRO + table('<w:keepNext/><w:jc w:val="right"/>') + SIGNED },
    )
    expect(body).toMatch(
      new RegExp(
        `<w:tc><w:tcPr>[^]*?</w:tcPr><w:p><w:pPr><w:keepNext/><w:jc w:val="right"/><w:pPrChange w:id="\\d+" w:author="${LABEL}" w:date="${DATE}"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t[^>]*>Price</w:t></w:r></w:p></w:tc>`,
      ),
    )
    expect(summary).toMatchObject({ formatChanges: 1, insertions: 0, deletions: 0 })
    expect(resolved(body, 'accept')).toContain('<w:pPr><w:keepNext/><w:jc w:val="right"/></w:pPr>')
    expect(resolved(body, 'reject')).toContain('<w:pPr><w:jc w:val="left"/></w:pPr>')
    expect(resolved(body, 'reject')).not.toContain('w:keepNext')
  })

  it('a cell paragraph that changed in both text and formatting carries both changes', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + tbl([tr(['Item', 'Unit price'])]) + SIGNED },
      {
        bodyXml: INTRO + tbl([tr(['Item', tc(p('Unit cost', '<w:jc w:val="center"/>'))])]) + SIGNED,
      },
    )
    expect(summary).toMatchObject({ formatChanges: 1, insertions: 1, deletions: 1 })
    expect(body).toContain('<w:pPrChange ')
    expect(texts(resolved(body, 'reject'))).toContain('Unit price')
    expect(texts(resolved(body, 'accept'))).toContain('Unit cost')
  })

  it('with Formatting off, or with the same formatting spelled differently, a cell paragraph is left alone', async () => {
    const off = await compare(
      { bodyXml: INTRO + table('<w:jc w:val="left"/>') + SIGNED },
      { bodyXml: INTRO + table('<w:jc w:val="right"/>') + SIGNED },
      { settings: { formatting: false } },
    )
    expect(off.body).not.toContain('w:pPrChange')
    expect(off.summary.formatChanges).toBe(0)
    const same = await compare(
      { bodyXml: INTRO + table('<w:keepNext w:val="1"/><w:jc w:val="left"/>') + SIGNED },
      { bodyXml: INTRO + table('<w:keepNext/><w:jc w:val="left"/><w:rPr><w:b/></w:rPr>') + SIGNED },
    )
    expect(same.body).not.toContain('w:pPrChange')
    expect(same.summary).toMatchObject({ formatChanges: 0, changedParagraphs: 0 })
  })

  it('a header paragraph and a text box paragraph are compared the same way', async () => {
    const box = (pPr: string) =>
      `<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:100pt;height:40pt"><v:textbox><w:txbxContent>${p('Boxed words.', pPr)}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`
    const { header, body, summary } = await compare(
      { bodyXml: INTRO + box('') + SIGNED, headerXml: p('Draft contract') },
      {
        bodyXml: INTRO + box('<w:jc w:val="right"/>') + SIGNED,
        headerXml: p('Draft contract', '<w:jc w:val="center"/>'),
      },
    )
    expect(header).toMatch(
      /<w:pPr><w:jc w:val="center"\/><w:pPrChange [^>]*><w:pPr><\/w:pPr><\/w:pPrChange><\/w:pPr>/,
    )
    expect(count(body, /<w:pPrChange /g)).toBeGreaterThanOrEqual(1)
    expect(summary.formatChanges).toBeGreaterThanOrEqual(2)
  })
})

describe('comparison: table rows', () => {
  const HEAD = tr(['Item', 'Price'])
  const BOLT = tr(['Bolt', 'ten dollars'])
  const NUT = tr(['Nut', 'two dollars'])

  it('a row only the revised table has is a tracked row insertion', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + tbl([HEAD, BOLT]) + SIGNED },
      { bodyXml: INTRO + tbl([HEAD, BOLT, NUT]) + SIGNED },
    )
    expect(body).toMatch(
      new RegExp(
        `<w:tr><w:trPr><w:ins w:id="\\d+" w:author="${LABEL}" w:date="${DATE}"/></w:trPr>`,
      ),
    )
    expect(count(body, /<w:tr>/g)).toBe(3)
    expect(count(body, /<w:del /g)).toBe(0)
    expect(summary).toMatchObject({ insertions: 1, deletions: 0, changedParagraphs: 1 })
    expect(texts(resolved(body, 'accept'))).toContain('Nuttwo dollars'.slice(0, 3))
    expect(count(resolved(body, 'reject'), /<w:tr>/g)).toBe(2)
    expect(count(resolved(body, 'accept'), /<w:tr>/g)).toBe(3)
  })

  it('a row only the original table has is a tracked row deletion, and a changed cell beside it is diffed', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + tbl([HEAD, NUT, BOLT]) + SIGNED },
      { bodyXml: INTRO + tbl([HEAD, tr(['Bolt', 'twelve dollars'])]) + SIGNED },
    )
    expect(body).toMatch(/<w:tr><w:trPr><w:del w:id="\d+" w:author="Rae Visor"[^>]*\/><\/w:trPr>/)
    expect(summary).toMatchObject({ insertions: 1, deletions: 2 })
    expect(count(resolved(body, 'accept'), /<w:tr>/g)).toBe(2)
    expect(texts(resolved(body, 'accept'))).toContain('twelve dollars')
    expect(count(resolved(body, 'reject'), /<w:tr>/g)).toBe(3)
    expect(texts(resolved(body, 'reject'))).toEqual(
      expect.arrayContaining(['Nut', 'two dollars', 'ten dollars']),
    )
  })

  it('Compare reads a row either table tracks as deleted as gone', async () => {
    const struck = (who: string) =>
      tr(
        [
          tc(p(`<w:del w:id="5" ${who}>${gone('Nut')}</w:del>`)),
          tc(p(`<w:del w:id="6" ${who}>${gone('two dollars')}</w:del>`)),
        ],
        `<w:del w:id="4" ${who}/>`,
      )
    const kept = await compare(
      { bodyXml: INTRO + tbl([HEAD, struck(ABE), BOLT]) + SIGNED },
      { bodyXml: INTRO + tbl([HEAD, NUT, BOLT]) + SIGNED },
    )
    expect(kept.body).toMatch(/<w:trPr><w:ins w:id="\d+" w:author="Rae Visor"/)
    expect(kept.body).not.toContain('Abe Origin')
    expect(count(resolved(kept.body, 'reject'), /<w:tr>/g)).toBe(2)
    const dropped = await compare(
      { bodyXml: INTRO + tbl([HEAD, NUT, BOLT]) + SIGNED },
      { bodyXml: INTRO + tbl([HEAD, struck(ABE), BOLT]) + SIGNED },
    )
    expect(dropped.body).toMatch(/<w:trPr><w:del w:id="\d+" w:author="Rae Visor"/)
    expect(dropped.body).not.toContain('Abe Origin')
    expect(count(dropped.body, /<w:tr>/g)).toBe(3)
    expect(count(resolved(dropped.body, 'accept'), /<w:tr>/g)).toBe(2)
    expect(texts(resolved(dropped.body, 'reject'))).toEqual(
      expect.arrayContaining(['Nut', 'two dollars']),
    )
  })

  const cellsPerRow = (xml: string) =>
    [...xml.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map((m) => count(m[0], /<w:tc>/g))

  it('a column only the revised table has is its cells tracked as inserted', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + tbl([HEAD, BOLT]) + SIGNED },
      {
        bodyXml:
          INTRO +
          tbl([tr(['Item', 'Stock', 'Price']), tr(['Bolt', 'many', 'ten dollars'])], 3) +
          SIGNED,
      },
    )
    expect(body).not.toMatch(/<w:(?:del|ins) [^>]*><w:tbl>/)
    expect(count(body, /<w:cellIns w:id="\d+" w:author="Rae Visor"/g)).toBe(2)
    expect(count(body, /<w:cellDel /g)).toBe(0)
    expect(count(body, /<w:tbl>/g)).toBe(1)
    expect(summary).toMatchObject({ insertions: 1, deletions: 0 })
    // the new cells sit in the second column
    const first = /<w:tr>[\s\S]*?<\/w:tr>/.exec(body)![0]
    expect(first.split('<w:tc>')[2]).toContain('<w:cellIns ')
    expect(cellsPerRow(resolved(body, 'accept'))).toEqual([3, 3])
    expect(texts(resolved(body, 'accept'))).toEqual(
      expect.arrayContaining(['Item', 'Stock', 'Price', 'many']),
    )
    expect(cellsPerRow(resolved(body, 'reject'))).toEqual([2, 2])
    expect(texts(resolved(body, 'reject'))).not.toContain('Stock')
    expect(texts(resolved(body, 'reject'))).toEqual(
      expect.arrayContaining(['Item', 'Price', 'Bolt', 'ten dollars']),
    )
  })

  it('a column only the original table has is its cells tracked as deleted, a changed cell beside it diffed', async () => {
    const { body, summary } = await compare(
      {
        bodyXml:
          INTRO +
          tbl([tr(['Item', 'Stock', 'Price']), tr(['Bolt', 'many', 'ten dollars'])], 3) +
          SIGNED,
      },
      { bodyXml: INTRO + tbl([HEAD, tr(['Bolt', 'twelve dollars'])]) + SIGNED },
    )
    expect(count(body, /<w:cellDel w:id="\d+" w:author="Rae Visor"/g)).toBe(2)
    expect(count(body, /<w:gridCol /g)).toBe(3)
    expect(summary).toMatchObject({ insertions: 1, deletions: 2 })
    expect(cellsPerRow(body)).toEqual([3, 3])
    expect(cellsPerRow(resolved(body, 'accept'))).toEqual([2, 2])
    expect(texts(resolved(body, 'accept'))).toEqual(
      expect.arrayContaining(['Item', 'Price', 'Bolt', 'twelve dollars']),
    )
    expect(texts(resolved(body, 'accept'))).not.toContain('many')
    expect(cellsPerRow(resolved(body, 'reject'))).toEqual([3, 3])
    expect(texts(resolved(body, 'reject'))).toEqual(
      expect.arrayContaining(['Stock', 'many', 'ten dollars']),
    )
  })

  it('a column and a row added together: every row has a cell for every column', async () => {
    const SCREW = tr(['Screw', 'three dollars'])
    const { body } = await compare(
      { bodyXml: INTRO + tbl([HEAD, NUT, BOLT, SCREW]) + SIGNED },
      {
        bodyXml:
          INTRO +
          tbl(
            [
              tr(['Item', 'Price', 'Stock']),
              tr(['Bolt', 'ten dollars', 'many']),
              tr(['Screw', 'three dollars', 'some']),
              tr(['Washer', 'one dollar', 'few']),
            ],
            3,
          ) +
          SIGNED,
      },
    )
    expect(cellsPerRow(body)).toEqual([3, 3, 3, 3, 3])
    expect(count(body, /<w:trPr><w:del /g)).toBe(1)
    expect(count(body, /<w:trPr><w:ins /g)).toBe(1)
    expect(cellsPerRow(resolved(body, 'accept'))).toEqual([3, 3, 3, 3])
    expect(cellsPerRow(resolved(body, 'reject'))).toEqual([2, 2, 2, 2])
    expect(texts(resolved(body, 'reject'))).toEqual(
      expect.arrayContaining(['Nut', 'two dollars', 'Bolt', 'ten dollars']),
    )
    expect(texts(resolved(body, 'reject'))).not.toContain('some')
  })

  it('a table with merged cells and a column added is still replaced whole', async () => {
    const wide = tr([tc(p('Parts'), '<w:gridSpan w:val="2"/>')])
    const { body } = await compare(
      { bodyXml: INTRO + tbl([wide, HEAD, BOLT]) + SIGNED },
      {
        bodyXml:
          INTRO +
          tbl(
            [
              tr([tc(p('Parts'), '<w:gridSpan w:val="3"/>')]),
              tr(['Item', 'Price', 'Stock']),
              tr(['Bolt', 'ten dollars', 'many']),
            ],
            3,
          ) +
          SIGNED,
      },
    )
    expect(body).toMatch(/<w:del [^>]*><w:tbl>/)
    expect(body).toMatch(/<w:ins [^>]*><w:tbl>/)
    expect(body).not.toContain('<w:cellIns')
  })

  it('Combine keeps a row the original tracks as inserted, under its author, in a table the revised one reshaped', async () => {
    const tracked = tr(
      [
        tc(p(`<w:ins w:id="5" ${ABE}>${r('Nut')}</w:ins>`)),
        tc(p(`<w:ins w:id="6" ${ABE}>${r('two dollars')}</w:ins>`)),
      ],
      `<w:ins w:id="4" ${ABE}/>`,
    )
    const { body, summary } = await compare(
      { bodyXml: INTRO + tbl([HEAD, tracked, BOLT]) + SIGNED },
      { bodyXml: INTRO + tbl([HEAD, tr(['Bolt', 'twelve dollars'])]) + SIGNED },
      { keepRevisions: true },
    )
    expect(body).toMatch(/<w:tr><w:trPr><w:ins w:id="9\d+" w:author="Abe Origin"[^>]*\/><\/w:trPr>/)
    expect(body).toMatch(/<w:ins w:id="9\d+" w:author="Abe Origin"[^>]*><w:r><w:t[^>]*>Nut<\/w:t>/)
    expect(summary).toMatchObject({ carried: 1, insertions: 1, deletions: 1 })
    expect(count(resolved(body, 'reject'), /<w:tr>/g)).toBe(2)
    expect(texts(resolved(body, 'reject'))).toContain('ten dollars')
    expect(count(resolved(body, 'accept'), /<w:tr>/g)).toBe(3)
  })

  it('Combine keeps a tracked row deletion and a tracked cell insertion of the original on the rows both have', async () => {
    const struck = tr(
      [
        tc(p(`<w:del w:id="5" ${ABE}>${gone('Nut')}</w:del>`)),
        tc(p(`<w:del w:id="6" ${ABE}>${gone('two dollars')}</w:del>`)),
      ],
      `<w:del w:id="4" ${ABE}/>`,
    )
    const cellIns = tr([tc(p('Bolt'), `<w:cellIns w:id="7" ${ABE}/>`), 'ten dollars'])
    const { body, summary } = await compare(
      { bodyXml: INTRO + tbl([HEAD, struck, cellIns]) + SIGNED },
      { bodyXml: INTRO + tbl([HEAD, NUT, BOLT]) + p('Signed by both parties. Twice.') },
      { keepRevisions: true },
    )
    expect(body).toMatch(/<w:trPr><w:del w:id="9\d+" w:author="Abe Origin"[^>]*\/><\/w:trPr>/)
    expect(body).toMatch(/<w:cellIns w:id="9\d+" w:author="Abe Origin"[^>]*\/><\/w:tcPr>/)
    expect(body).toMatch(/<w:del w:id="9\d+" w:author="Abe Origin"[^>]*><w:r><w:delText[^>]*>Nut</)
    expect(summary.carried).toBe(4)
    expect(count(resolved(body, 'accept'), /<w:tr>/g)).toBe(2)
  })
})

describe('comparison: property changes of the original that Combine carries', () => {
  const SHADE = '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'
  const oldTcPr = `<w:tcPrChange w:id="3" ${ABE}><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr></w:tcPrChange>`
  const oldTrPr = `<w:trPrChange w:id="4" ${ABE}><w:trPr/></w:trPrChange>`
  const table = (records: boolean, shade = SHADE) =>
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/>${
      records
        ? `<w:tblPrChange w:id="5" ${ABE}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange>`
        : ''
    }</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>${
      records
        ? '<w:tblGridChange w:id="6"><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="3000"/></w:tblGrid></w:tblGridChange>'
        : ''
    }</w:tblGrid>${tr(['Item', 'Price'])}${tr(
      [tc(p('Bolt'), shade + (records ? oldTcPr : '')), 'ten dollars'],
      '<w:cantSplit/>' + (records ? oldTrPr : ''),
    )}</w:tbl>`

  it('w:tblPrChange, w:tblGridChange, w:trPrChange and w:tcPrChange come along under their author and reject to what they recorded', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + table(true) + SIGNED },
      { bodyXml: INTRO + table(false) + p('Signed by both parties. Twice.') },
      { keepRevisions: true },
    )
    expect(body).toMatch(/<w:tblPrChange w:id="9\d+" w:author="Abe Origin"/)
    expect(body).toMatch(/<w:tblGridChange w:id="9\d+">/)
    expect(body).toMatch(/<w:cantSplit\/><w:trPrChange w:id="9\d+" w:author="Abe Origin"/)
    expect(body).toMatch(/<w:shd [^>]*\/><w:tcPrChange w:id="9\d+" w:author="Abe Origin"/)
    expect(summary.carried).toBe(4)
    const rejected = resolved(body, 'reject')
    expect(rejected).not.toContain('<w:jc w:val="center"/>')
    expect(rejected).not.toContain('<w:cantSplit/>')
    expect(rejected).not.toContain('<w:shd ')
    expect(rejected).toContain('<w:gridCol w:w="1000"/><w:gridCol w:w="3000"/>')
    const accepted = resolved(body, 'accept')
    expect(accepted).toContain('<w:jc w:val="center"/>')
    expect(accepted).toContain('<w:cantSplit/>')
    expect(accepted).toContain(SHADE)
    expect(accepted).not.toMatch(/Change\b/)
  })

  it('a record is not carried onto properties the revised document changed, nor without Combine', async () => {
    const other = '<w:shd w:val="clear" w:color="auto" w:fill="00FF00"/>'
    const combined = await compare(
      { bodyXml: INTRO + table(true) + SIGNED },
      { bodyXml: INTRO + table(false, other) + p('Signed by both parties. Twice.') },
      { keepRevisions: true },
    )
    expect(combined.body).not.toContain('<w:tcPrChange')
    expect(combined.body).toContain('<w:trPrChange')
    const compared = await compare(
      { bodyXml: INTRO + table(true) + SIGNED },
      { bodyXml: INTRO + table(false) + p('Signed by both parties. Twice.') },
    )
    expect(compared.body).not.toMatch(/<w:t(?:bl|r|c)PrChange|<w:tblGridChange/)
  })
})

describe('comparison: existing changes Combine carries', () => {
  const bold = `<w:r><w:rPr><w:b/><w:rPrChange w:id="3" ${ABE}><w:rPr/></w:rPrChange></w:rPr><w:t xml:space="preserve">turns bold</w:t></w:r>`

  it("the original's tracked formatting change stays under its author when the revised text is formatted as it was before or after", async () => {
    for (const revised of [
      r('This line turns bold here.'),
      r('This line ') + r('turns bold', '<w:b/>') + r(' here.'),
    ]) {
      const { body, summary } = await compare(
        { bodyXml: INTRO + p(r('This line ') + bold + r(' here.')) + SIGNED },
        { bodyXml: INTRO + p(revised) + SIGNED },
        { keepRevisions: true },
      )
      expect(body).toMatch(
        /<w:rPr><w:b\/><w:rPrChange w:id="9\d+" w:author="Abe Origin" w:date="2026-11-01T08:00:00Z"><w:rPr\/><\/w:rPrChange><\/w:rPr><w:t[^>]*>turns bold</,
      )
      expect(body).not.toContain(LABEL)
      expect(summary).toMatchObject({ carried: 1, formatChanges: 0 })
    }
  })

  it('Compare reads the same original as accepted: its formatting change is not kept', async () => {
    const { body } = await compare(
      { bodyXml: INTRO + p(r('This line ') + bold + r(' here.')) + SIGNED },
      { bodyXml: INTRO + p('This line turns bold here.') + SIGNED },
    )
    expect(body).not.toContain('Abe Origin')
    expect(body).toContain(`w:author="${LABEL}"`)
  })

  const from = (name: string) =>
    `<w:p><w:pPr><w:rPr><w:moveFrom w:id="7" ${ABE}/></w:rPr></w:pPr><w:moveFromRangeStart w:id="8" ${ABE} w:name="${name}"/><w:moveFrom w:id="9" ${ABE}>${r('A clause that moves to the end.')}</w:moveFrom><w:moveFromRangeEnd w:id="8"/></w:p>`
  const to = (name: string) =>
    `<w:p><w:pPr><w:rPr><w:moveTo w:id="10" ${ABE}/></w:rPr></w:pPr><w:moveToRangeStart w:id="11" ${ABE} w:name="${name}"/><w:moveTo w:id="12" ${ABE}>${r('A clause that moves to the end.')}</w:moveTo><w:moveToRangeEnd w:id="11"/></w:p>`
  const TERM = p('The term of this agreement is twelve months.')

  it('a move the original tracks stays a move when both its ends come along', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + from('move1') + TERM + SIGNED + to('move1') },
      {
        bodyXml:
          INTRO +
          p('A clause that moves to the end.') +
          p('The term of this agreement is six months.') +
          SIGNED,
      },
      { keepRevisions: true },
    )
    expect(count(body, /<w:moveFrom [^>]*w:author="Abe Origin"/g)).toBe(2)
    expect(count(body, /<w:moveTo [^>]*w:author="Abe Origin"/g)).toBe(2)
    const name = /<w:moveFromRangeStart [^>]*w:name="([^"]+)"/.exec(body)?.[1]
    expect(name).toBeTruthy()
    expect(body).toContain(
      `<w:moveToRangeStart w:id="${/<w:moveToRangeStart w:id="(\d+)"/.exec(body)?.[1]}"`,
    )
    expect(body).toMatch(new RegExp(`<w:moveToRangeStart [^>]*w:name="${name}"`))
    expect(summary.moves).toBe(0)
    expect(texts(resolved(body, 'accept'))[3]).toBe('A clause that moves to the end.')
    expect(texts(resolved(body, 'reject'))[1]).toBe('A clause that moves to the end.')
  })

  it('a move with one end only comes along as a plain insertion', async () => {
    const { body } = await compare(
      { bodyXml: INTRO + TERM + SIGNED + to('move1') },
      { bodyXml: INTRO + p('The term of this agreement is six months.') + SIGNED },
      { keepRevisions: true },
    )
    expect(body).not.toContain('w:moveTo')
    expect(body).toMatch(
      /<w:ins w:id="\d+" w:author="Abe Origin"[^>]*><w:r><w:t[^>]*>A clause that moves/,
    )
    expect(body).toMatch(/<w:rPr><w:ins w:id="\d+" w:author="Abe Origin"[^>]*\/><\/w:rPr>/)
  })
})

describe('comparison: pictures and equations', () => {
  const figure = p(r('Logo: ') + picture + r(' as shown.'))

  it('a picture replaced by another is the old picture deleted and the new one inserted', async () => {
    const { body, summary, zip, bytes } = await compare(
      { bodyXml: INTRO + figure + SIGNED, withImage: true },
      { bodyXml: INTRO + figure + SIGNED, withImage: true, png: OTHER_PNG },
    )
    expect(summary).toMatchObject({ insertions: 1, deletions: 1, changedParagraphs: 1 })
    const deleted = /<w:del [^>]*><w:r><w:drawing>[\s\S]*?r:embed="(rId\d+)"[\s\S]*?<\/w:del>/.exec(
      body,
    )
    const inserted =
      /<w:ins [^>]*><w:r><w:drawing>[\s\S]*?r:embed="(rId\d+)"[\s\S]*?<\/w:ins>/.exec(body)
    expect(deleted?.[1]).toBeTruthy()
    expect(inserted?.[1]).toBe('rId10')
    expect(deleted?.[1]).not.toBe('rId10')
    const rels = (await zip.file('word/_rels/document.xml.rels')?.async('string')) ?? ''
    const target = new RegExp(`Id="${deleted![1]}"[^>]*Target="([^"]+)"`).exec(rels)?.[1]
    expect(await zip.file(`word/${target}`)?.async('base64')).toBe(TINY_PNG_BASE64)
    expect(await zip.file('word/media/image1.png')?.async('base64')).toBe(OTHER_PNG)
    // the redline reads back with the old picture pending deletion
    const runs = (await parseDocx(bytes)).blocks.flatMap((b) => b.runs ?? [])
    expect(runs.filter((run) => run.image).map((run) => !!run.del)).toEqual([true, false])
  })

  it('the same picture on both sides is kept while the text around it is diffed', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + figure + SIGNED, withImage: true },
      {
        bodyXml: INTRO + p(r('Logo: ') + picture + r(' as shown here.')) + SIGNED,
        withImage: true,
      },
    )
    expect(count(body, /<w:drawing>/g)).toBe(1)
    expect(summary).toMatchObject({ insertions: 1, deletions: 1 })
    expect(body).not.toMatch(/<w:(?:ins|del) [^>]*><w:r><w:drawing>/)
  })

  it('an equation that changed is the old equation deleted and the new one inserted, each whole and once', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + p(r('Area is ') + math('x + y = 2') + r(' here.')) + SIGNED },
      { bodyXml: INTRO + p(r('Area is ') + math('x + y = 3') + r(' here.')) + SIGNED },
    )
    expect(count(body, /<m:oMath>/g)).toBe(2)
    expect(body).toMatch(
      /<w:del [^>]*><m:oMath>(?:<m:r><m:t>[^<]*<\/m:t><\/m:r>)*<\/m:oMath><\/w:del>/,
    )
    expect(body).toMatch(
      /<w:ins [^>]*><m:oMath>(?:<m:r><m:t>[^<]*<\/m:t><\/m:r>)*<\/m:oMath><\/w:ins>/,
    )
    expect(summary).toMatchObject({ insertions: 1, deletions: 1 })
    expect(texts(resolved(body, 'accept'))).toContain('Area is x+y=3 here.')
    expect(texts(resolved(body, 'reject'))).toContain('Area is x+y=2 here.')
  })

  it('the same equation is kept while the text around it is diffed', async () => {
    const { body } = await compare(
      { bodyXml: INTRO + p(r('Area is ') + math('x + y = 2') + r(' here.')) + SIGNED },
      { bodyXml: INTRO + p(r('Area is ') + math('x + y = 2') + r(' there.')) + SIGNED },
    )
    expect(count(body, /<m:oMath>/g)).toBe(1)
    expect(body).not.toMatch(/<w:(?:ins|del) [^>]*><m:oMath>/)
  })
})

describe('comparison: text that moved inside a paragraph or between two', () => {
  const A = 'The supplier shall deliver the goods to the address of the customer.'
  const B = 'Delivery takes place on working days only.'
  const C = 'The customer inspects the goods on arrival.'

  const moveNames = (xml: string) =>
    [...xml.matchAll(/<w:move(?:From|To)RangeStart [^>]*w:name="([^"]+)"/g)].map((m) => m[1])

  it('in a header, a sentence that changed place and a paragraph that changed place are moves', async () => {
    const D = 'Payment is due within thirty days.'
    const { header, summary } = await compare(
      { bodyXml: INTRO + SIGNED, headerXml: p(`${A} ${B} ${C}`) + p(D) + p('Page heading') },
      { bodyXml: INTRO + SIGNED, headerXml: p(`${B} ${A} ${C}`) + p('Page heading') + p(D) },
    )
    expect(count(header, /<w:moveFrom w:id/g)).toBe(3)
    expect(count(header, /<w:moveTo w:id/g)).toBe(3)
    expect(count(header, /<w:(?:ins|del) /g)).toBe(0)
    expect(header).not.toContain('w:delText')
    // the paragraph that moved carries the move on its mark too
    expect(header).toMatch(/<w:rPr><w:moveFrom w:id="\d+"[^>]*\/><\/w:rPr>/)
    expect(header).toMatch(/<w:rPr><w:moveTo w:id="\d+"[^>]*\/><\/w:rPr>/)
    const names = moveNames(header)
    expect(names).toHaveLength(4)
    expect(new Set(names).size).toBe(2)
    expect(summary).toMatchObject({ moves: 2, insertions: 0, deletions: 0 })
    expect(texts(resolved(header, 'accept'))).toEqual([`${B} ${A} ${C}`, 'Page heading', D])
    expect(texts(resolved(header, 'reject'))).toEqual([`${A} ${B} ${C}`, D, 'Page heading'])
  })

  it('in table cells, a sentence that went to another cell is a move; not with Moves off, not under Combine', async () => {
    const original = {
      bodyXml: INTRO + tbl([tr([`${A} ${B}`, C]), tr(['Bolt', 'ten dollars'])]) + SIGNED,
    }
    const revised = {
      bodyXml: INTRO + tbl([tr([A, `${C} ${B}`]), tr(['Bolt', 'ten dollars'])]) + SIGNED,
    }
    const { body, summary } = await compare(original, revised)
    expect(count(body, /<w:moveFrom /g)).toBe(1)
    expect(count(body, /<w:moveTo /g)).toBe(1)
    expect(count(body, /<w:(?:ins|del) /g)).toBe(0)
    const names = moveNames(body)
    expect(names).toHaveLength(2)
    expect(names[0]).toBe(names[1])
    expect(summary).toMatchObject({ moves: 1, insertions: 0, deletions: 0 })
    expect(texts(resolved(body, 'accept'))).toEqual(expect.arrayContaining([A, `${C} ${B}`]))
    expect(texts(resolved(body, 'reject'))).toEqual(expect.arrayContaining([`${A} ${B}`, C]))
    const reopened = await parseDocx((await compare(original, revised)).bytes)
    expect(reopened.internal.documentXml).toContain('<w:moveTo ')

    const off = await compare(original, revised, { settings: { moves: false } })
    expect(off.body).not.toContain('w:moveFrom')
    expect(off.summary).toMatchObject({ moves: 0, insertions: 1, deletions: 1 })
    const combined = await compare(original, revised, { keepRevisions: true })
    expect(combined.body).not.toContain('w:moveFrom')
  })

  it('a sentence that changed place inside its paragraph is a move', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + p(`${A} ${B} ${C}`) + SIGNED },
      { bodyXml: INTRO + p(`${B} ${A} ${C}`) + SIGNED },
    )
    expect(count(body, /<w:moveFrom /g)).toBe(1)
    expect(count(body, /<w:moveTo /g)).toBe(1)
    expect(count(body, /<w:(?:ins|del) /g)).toBe(0)
    const names = [...body.matchAll(/<w:move(?:From|To)RangeStart [^>]*w:name="([^"]+)"/g)].map(
      (m) => m[1],
    )
    expect(names).toHaveLength(2)
    expect(names[0]).toBe(names[1])
    expect(summary).toMatchObject({ moves: 1, insertions: 0, deletions: 0 })
    expect(texts(resolved(body, 'accept'))[1]).toBe(`${B} ${A} ${C}`)
    expect(texts(resolved(body, 'reject'))[1]).toBe(`${A} ${B} ${C}`)
  })

  it('a sentence that went to another paragraph is a move', async () => {
    const { body, summary } = await compare(
      { bodyXml: INTRO + p(`${A} ${B}`) + p(C) + SIGNED },
      { bodyXml: INTRO + p(A) + p(`${C} ${B}`) + SIGNED },
    )
    expect(count(body, /<w:moveFrom /g)).toBe(1)
    expect(count(body, /<w:moveTo /g)).toBe(1)
    expect(summary.moves).toBe(1)
    expect(texts(resolved(body, 'reject'))).toEqual(expect.arrayContaining([`${A} ${B}`, C]))
    expect(texts(resolved(body, 'accept'))).toEqual(expect.arrayContaining([A, `${C} ${B}`]))
  })

  it('with Moves off, a short phrase, or a stretch with two twins stays a deletion and an insertion', async () => {
    const off = await compare(
      { bodyXml: INTRO + p(`${A} ${B} ${C}`) + SIGNED },
      { bodyXml: INTRO + p(`${B} ${A} ${C}`) + SIGNED },
      { settings: { moves: false } },
    )
    expect(off.body).not.toContain('w:moveFrom')
    expect(off.summary).toMatchObject({ moves: 0, insertions: 1, deletions: 1 })
    const short = await compare(
      {
        bodyXml:
          INTRO + p(`Red apples and green pears are sold here every day of the week.`) + SIGNED,
      },
      {
        bodyXml:
          INTRO + p(`Green pears and red apples are sold here every day of the week.`) + SIGNED,
      },
    )
    expect(short.body).not.toContain('w:moveFrom')
    expect(short.summary.moves).toBe(0)
  })
})
