import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  STYLE_SET_IDS,
  customStyleSetUpserts,
  listStylesXml,
  parseDocx,
  previewFontSettings,
  readDocxStylesXml,
  saveDocx,
  styleRestoreUpserts,
  styleSetUpserts,
  styleSetXmlOf,
  upsertStyleXml,
} from '../src'
import { buildDocx } from './helpers/build-docx'
import { xmlProblem } from './helpers/strict-xml'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const styles = (inner: string) => `<w:styles ${W}>${inner}</w:styles>`
const apply = (xml: string, set: (typeof STYLE_SET_IDS)[number]) =>
  styleSetUpserts(set, xml).reduce(upsertStyleXml, xml)
const byName = (xml: string, name: string) =>
  listStylesXml(xml).find((s) => s.name.toLowerCase() === name)

describe('style sets', () => {
  it('lists the sets in gallery order, each with the styles a card previews', () => {
    expect(STYLE_SET_IDS[0]).toBe('distinctive')
    expect(STYLE_SET_IDS).toHaveLength(12)
    for (const set of STYLE_SET_IDS) {
      const out = apply(styles(''), set)
      for (const name of ['normal', 'title', 'heading 1']) expect(byName(out, name)).toBeDefined()
    }
  })

  it('overwrites a style under the id the document gave it and leaves the rest alone', () => {
    const doc = styles(
      '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:rPr><w:sz w:val="40"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:rPr><w:b/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Mine"><w:name w:val="Mine"/><w:basedOn w:val="a"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>',
    )
    const out = apply(doc, 'modern')
    const list = listStylesXml(out)
    expect(list.filter((s) => s.name === 'Normal')).toHaveLength(1)
    const normal = byName(out, 'normal')!
    expect(normal.styleId).toBe('a')
    expect(normal.isDefault).toBe(true)
    expect(normal.xml).toContain('<w:sz w:val="20"/>')
    const heading = byName(out, 'heading 1')!
    expect(heading.styleId).toBe('1')
    expect(heading.basedOn).toBe('a')
    expect(heading.xml).toContain('<w:caps/>')
    expect(byName(out, 'title')!.basedOn).toBe('a')
    expect(byName(out, 'mine')!.xml).toContain('<w:u w:val="single"/>')
  })

  it('gives the look to a default paragraph style that is not called Normal, keeping its name', () => {
    const doc = styles(
      '<w:style w:type="paragraph" w:default="1" w:styleId="Standard"><w:name w:val="Standard"/></w:style>',
    )
    const out = apply(doc, 'manuscript')
    const list = listStylesXml(out)
    expect(list.filter((s) => s.isDefault)).toHaveLength(1)
    const standard = list.find((s) => s.styleId === 'Standard')!
    expect(standard.name).toBe('Standard')
    expect(standard.xml).toContain('w:line="480"')
    expect(byName(out, 'normal')).toBeUndefined()
    expect(byName(out, 'heading 1')!.basedOn).toBe('Standard')
  })

  it('a second set replaces the first', () => {
    const once = apply(styles(''), 'fancy')
    const twice = apply(once, 'simple')
    expect(listStylesXml(twice).length).toBe(listStylesXml(apply(styles(''), 'simple')).length + 1)
    expect(byName(twice, 'title')!.xml).not.toContain('<w:shd')
  })
})

describe('text effects of a set', () => {
  it("Thatch keeps Word's outlined title and hollow heading, and the page reads the outline", async () => {
    const xml = apply(styles(''), 'thatch')
    const title = byName(xml, 'title')!.xml
    expect(title).toContain('<w:color w:val="FFFFFF" w:themeColor="background1"/>')
    expect(title).toMatch(/<w14:textOutline w14:w="13335"[\s\S]*<w14:schemeClr w14:val="accent1">/)
    expect(byName(xml, 'heading 2')!.xml).toContain('<w14:textFill><w14:noFill/></w14:textFill>')

    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }),
    )
    const stylesXml = await readDocxStylesXml(parsed.internal.originalBytes)
    const upserts = styleSetUpserts('thatch', stylesXml)
    const live = await previewFontSettings(parsed, upserts)
    const outline = [...live.styles.values()].find((s) => s.name === 'Title')?.display?.textOutline
    expect(outline).toMatchObject({ widthPt: 1.05 })
    expect(outline?.color).toMatch(/^[0-9A-F]{6}$/)

    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      styleUpserts: upserts,
    })
    const out = await (await JSZip.loadAsync(saved)).file('word/styles.xml')!.async('string')
    expect(xmlProblem(out)).toBeNull()
    expect(out).toMatch(
      /<w:styles\b[^>]*xmlns:w14="http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordml"/,
    )
    const reopened = await parseDocx(saved)
    expect(
      [...reopened.styles.values()].find((s) => s.name === 'Title')?.display?.textOutline,
    ).toEqual(outline)
  })
})

describe('a document saved as a style set', () => {
  const doc = styles(
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="26"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="300"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="30"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:link w:val="Heading1Char"/><w:rPr><w:color w:val="C00000"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char"><w:name w:val="Heading 1 Char"/><w:link w:val="Heading1"/><w:rPr><w:color w:val="C00000"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Mine"><w:name w:val="Mine"/><w:basedOn w:val="Normal"/></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>',
  )

  it('keeps the built-in paragraph and character styles, the defaults folded into the parentless ones', () => {
    const set = styleSetXmlOf(doc)
    const kept = listStylesXml(`<w:styles>${set}</w:styles>`)
    expect(kept.map((s) => s.styleId)).toEqual(['Normal', 'Heading1', 'Heading1Char'])
    const normal = kept[0]!.xml
    expect(normal).toContain('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>')
    expect(normal).toContain('<w:sz w:val="30"/>')
    expect(normal).not.toContain('<w:sz w:val="26"/>')
    expect(normal).toContain('<w:spacing w:after="300"/>')
    expect(kept[1]!.xml).not.toContain('Georgia')
  })

  it('gives its look to another document, and a restore puts that document back', () => {
    const other = styles(
      '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>' +
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Own"><w:name w:val="Own"/></w:style>',
    )
    const applied = customStyleSetUpserts(styleSetXmlOf(doc), other).reduce(upsertStyleXml, other)
    expect(byName(applied, 'normal')!.styleId).toBe('a')
    expect(byName(applied, 'normal')!.xml).toContain('Georgia')
    expect(byName(applied, 'heading 1')!.xml).toContain('C00000')
    expect(byName(applied, 'own')).toBeDefined()
    expect(byName(applied, 'mine')).toBeUndefined()

    const restored = styleRestoreUpserts(other).reduce(upsertStyleXml, applied)
    expect(byName(restored, 'normal')!.xml).toBe(byName(other, 'normal')!.xml)
  })
})
