import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  listStylesXml,
  parseDocx,
  readDocxStylesXml,
  saveDocx,
  styleCopyUpserts,
  upsertStyleXml,
  writeDocxStyles,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'
import { xmlProblem } from './helpers/strict-xml'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const style = (id: string, inner: string, attrs = 'w:type="paragraph" w:customStyle="1"') =>
  `<w:style ${attrs} w:styleId="${id}"><w:name w:val="${id}"/>${inner}</w:style>`
const stylesXml = (...styles: string[]) => `<w:styles ${W}>${styles.join('')}</w:styles>`

const NORMAL = style('Normal', '', 'w:type="paragraph" w:default="1"')
const QUOTE = style(
  'MyQuote',
  '<w:basedOn w:val="Normal"/><w:next w:val="MyQuote"/><w:link w:val="MyQuoteChar"/><w:rPr><w:i/></w:rPr>',
)
const QUOTE_CHAR = style(
  'MyQuoteChar',
  '<w:link w:val="MyQuote"/><w:rPr><w:i/></w:rPr>',
  'w:type="character" w:customStyle="1"',
)
const CHILD = style('MyChild', '<w:basedOn w:val="MyQuote"/><w:next w:val="MyQuote"/>')

describe('deleting a style from styles.xml', () => {
  it('takes the style and its linked character twin out and re-parents what was based on it', () => {
    const xml = upsertStyleXml(stylesXml(NORMAL, QUOTE, QUOTE_CHAR, CHILD), {
      styleId: 'MyQuote',
      remove: true,
    })
    const left = listStylesXml(xml)
    expect(left.map((s) => s.styleId)).toEqual(['Normal', 'MyChild'])
    expect(left[1]).toMatchObject({ basedOn: 'Normal' })
    expect(left[1].next).toBeUndefined()
  })

  it('leaves the part alone when the style is not there', () => {
    const xml = stylesXml(NORMAL)
    expect(upsertStyleXml(xml, { styleId: 'Gone', remove: true })).toBe(xml)
  })

  it('replace builds the style from scratch instead of patching the old definition', () => {
    const xml = upsertStyleXml(stylesXml(NORMAL, QUOTE), {
      styleId: 'MyQuote',
      replace: true,
      name: 'Fresh',
      rPr: { bold: true },
    })
    const fresh = listStylesXml(xml).find((s) => s.styleId === 'MyQuote')!
    expect(fresh.name).toBe('Fresh')
    expect(fresh.xml).not.toContain('<w:i/>')
    expect(fresh.xml).toContain('<w:b/>')
  })

  it('a save without the style drops it and every reference to it in the body', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:pStyle w:val="MyQuote"/></w:pPr><w:r><w:t>quoted</w:t></w:r></w:p>' +
        '<w:p><w:r><w:rPr><w:rStyle w:val="MyQuoteChar"/></w:rPr><w:t>run</w:t></w:r></w:p>',
      extraStylesXml: QUOTE + QUOTE_CHAR,
    })
    const parsed = await parseDocx(bytes)
    const blocks = parsed.blocks.map((_, i) => ({ kind: 'original' as const, docxIndex: i }))
    const saved = await saveDocx(parsed, blocks, {
      styleUpserts: [{ styleId: 'MyQuote', remove: true }],
    })
    const zip = await JSZip.loadAsync(saved)
    const styles = await zip.file('word/styles.xml')!.async('string')
    const body = await zip.file('word/document.xml')!.async('string')
    expect(styles).not.toContain('MyQuote')
    expect(body).not.toContain('MyQuote')
    expect(body).toContain('quoted')
    const reopened = await parseDocx(saved)
    expect(reopened.styles.has('MyQuote')).toBe(false)
    expect(reopened.blocks[0].styleId ?? null).toBeNull()
  })
})

describe('copying styles between files', () => {
  it('brings the parents and the linked twin the target lacks, parents first', () => {
    const source = stylesXml(NORMAL, QUOTE, QUOTE_CHAR, CHILD)
    const target = stylesXml(NORMAL)
    const ups = styleCopyUpserts(source, 'MyChild', target)
    expect(ups.map((u) => u.styleId)).toEqual(['MyQuoteChar', 'MyQuote', 'MyChild'])
    let out = target
    for (const up of ups) out = upsertStyleXml(out, up)
    const copied = listStylesXml(out)
    expect(copied.map((s) => s.styleId)).toEqual(['Normal', 'MyQuoteChar', 'MyQuote', 'MyChild'])
    expect(copied[3]).toMatchObject({ basedOn: 'MyQuote', next: 'MyQuote' })
  })

  it('overwrites a style of the same name in place and renames a clashing id', () => {
    const source = stylesXml(
      NORMAL,
      style('Fancy', '<w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr>'),
    )
    const sameName = stylesXml(
      NORMAL,
      `<w:style w:type="paragraph" w:customStyle="1" w:styleId="X1"><w:name w:val="fancy"/></w:style>`,
    )
    expect(styleCopyUpserts(source, 'Fancy', sameName).map((u) => u.styleId)).toEqual(['X1'])
    const sameId = stylesXml(
      NORMAL,
      `<w:style w:type="paragraph" w:customStyle="1" w:styleId="Fancy"><w:name w:val="Other"/></w:style>`,
    )
    expect(styleCopyUpserts(source, 'Fancy', sameId).map((u) => u.styleId)).toEqual(['Fancy1'])
  })

  it('leaves numbering links, revision ids and the default flag behind', () => {
    const source = stylesXml(
      style(
        'Listy',
        '<w:rsid w:val="00AB12CD"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr><w:ind w:left="720"/></w:pPr>',
        'w:type="paragraph" w:default="1" w:customStyle="1"',
      ),
    )
    const [up] = styleCopyUpserts(source, 'Listy', stylesXml(NORMAL))
    expect(up.sourceXml).not.toMatch(/numPr|rsid|w:default/)
    expect(up.sourceXml).toContain('<w:ind w:left="720"/>')
  })

  it('writes the edits into another file and reads them back', async () => {
    const other = await buildDocx({
      bodyXml: '<w:p><w:pPr><w:pStyle w:val="MyQuote"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      extraStylesXml: QUOTE + QUOTE_CHAR,
    })
    const source = stylesXml(
      NORMAL,
      style('Fancy', '<w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr>'),
    )
    const ups = styleCopyUpserts(source, 'Fancy', await readDocxStylesXml(other))
    const written = await writeDocxStyles(other, [...ups, { styleId: 'MyQuote', remove: true }])
    const after = listStylesXml(await readDocxStylesXml(written))
    expect(after.some((s) => s.styleId === 'Fancy')).toBe(true)
    expect(after.some((s) => s.styleId.startsWith('MyQuote'))).toBe(false)
    const body = await (await JSZip.loadAsync(written)).file('word/document.xml')!.async('string')
    expect(body).not.toContain('MyQuote')
    expect((await parseDocx(written)).styles.get('Fancy')?.display).toMatchObject({ bold: true })
  })

  it('declares the prefix of a copied text effect the other file does not know', async () => {
    const other = await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' })
    const source =
      `<w:styles ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
      NORMAL +
      style('Fancy', '<w:rPr><w:b/><w14:ligatures w14:val="standardContextual"/></w:rPr>') +
      '</w:styles>'
    const ups = styleCopyUpserts(source, 'Fancy', await readDocxStylesXml(other))
    const styles = await readDocxStylesXml(await writeDocxStyles(other, ups))
    expect(styles).toContain('<w14:ligatures w14:val="standardContextual"/>')
    expect(xmlProblem(styles)).toBeNull()
  })
})
