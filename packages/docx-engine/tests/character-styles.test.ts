import { describe, expect, it } from 'vitest'
import { generateParagraphXml, parseDocx, type GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

const EMPHASIS_STYLES =
  '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/>' +
  '<w:rPr><w:b/><w:color w:val="C00000"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:basedOn w:val="Strong"/>' +
  '<w:rPr><w:i/><w:u w:val="single"/></w:rPr></w:style>'

describe('character styles (w:rStyle)', () => {
  it('parses character styles into the styles map with resolved basedOn display', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      extraStylesXml: EMPHASIS_STYLES,
    })
    const doc = await parseDocx(bytes)
    const strong = doc.styles.get('Strong')!
    expect(strong.type).toBe('character')
    expect(strong.display).toMatchObject({ bold: true, color: 'C00000' })
    const emphasis = doc.styles.get('Emphasis')!
    // inherits bold/color from Strong, adds its own italic/underline
    expect(emphasis.display).toMatchObject({
      bold: true,
      color: 'C00000',
      italic: true,
      underline: true,
    })
    expect(doc.styles.get('Hyperlink')!.display).toMatchObject({ underline: true, color: '0563C1' })
  })

  it('captures w:rStyle on runs, except the implied Hyperlink style', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>styled</w:t></w:r>' +
        '<w:r><w:t>plain</w:t></w:r>' +
        '<w:hyperlink r:id="rId20"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>link</w:t></w:r></w:hyperlink></w:p>',
      extraStylesXml: EMPHASIS_STYLES,
      extraRels:
        '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>',
    })
    const doc = await parseDocx(bytes)
    const runs = doc.blocks[0].runs!
    expect(runs[0].styleId).toBe('Emphasis')
    expect(runs[1].styleId).toBeUndefined()
    expect(runs[2].link?.href).toBe('https://example.com/')
    expect(runs[2].styleId).toBeUndefined()
  })

  it('does not merge adjacent runs with different character styles', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>a</w:t></w:r>' +
        '<w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>b</w:t></w:r></w:p>',
      extraStylesXml: EMPHASIS_STYLES,
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs!.map((r) => r.styleId)).toEqual(['Strong', 'Emphasis'])
  })

  it('regenerates w:rStyle first in rPr schema order', () => {
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ text: 'x', styleId: 'Emphasis', bold: true, font: '宋体' }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rStyle w:val="Emphasis"/>')
    expect(xml.indexOf('<w:rStyle')).toBeLessThan(xml.indexOf('<w:rFonts'))
  })

  it('round-trips styleId through generate -> parse', async () => {
    const para = generateParagraphXml(
      { type: 'paragraph', runs: [{ text: '强调', styleId: 'Emphasis' }] },
      GEN_CTX,
    )
    const bytes = await buildDocx({ bodyXml: para, extraStylesXml: EMPHASIS_STYLES })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs![0].styleId).toBe('Emphasis')
  })

  it('resolves paragraph style basedOn chains and survives cycles', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/>' +
        '<w:rPr><w:rFonts w:ascii="Georgia"/><w:sz w:val="20"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Base"/>' +
        '<w:rPr><w:i/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="CycA"><w:name w:val="CycA"/><w:basedOn w:val="CycB"/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="CycB"><w:name w:val="CycB"/><w:basedOn w:val="CycA"/></w:style>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.styles.get('Quote')!.display).toMatchObject({
      italic: true,
      font: 'Georgia',
      sizeHalfPoints: 20,
    })
    expect(doc.styles.get('CycA')).toBeDefined()
    expect(doc.styles.get('CycB')).toBeDefined()
  })
})

describe('linkedStyle (w:link) and docDefaults backfill', () => {
  it('when a character-style shell has no rPr, run-level display attributes come from the linked paragraph style', async () => {
    const styles =
      '<w:style w:type="paragraph" w:styleId="Heading9"><w:name w:val="heading 9"/>' +
      '<w:link w:val="Heading9Char"/>' +
      '<w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="24"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Heading9Char"><w:name w:val="标题 9 字符"/>' +
      '<w:link w:val="Heading9"/></w:style>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>', extraStylesXml: styles }),
    )
    expect(doc.styles.get('Heading9Char')!.display).toMatchObject({
      bold: true,
      color: '2F5496',
      sizeHalfPoints: 24,
    })
    // A style's own properties are not overridden by the link
    const styles2 =
      '<w:style w:type="paragraph" w:styleId="Quote2"><w:name w:val="Quote2"/>' +
      '<w:link w:val="Quote2Char"/><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Quote2Char"><w:name w:val="Quote2 Char"/>' +
      '<w:link w:val="Quote2"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>'
    const doc2 = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>', extraStylesXml: styles2 }),
    )
    expect(doc2.styles.get('Quote2Char')!.display).toMatchObject({ italic: true, color: 'FF0000' })
  })

  it('docDefaults parses bold/italic/color from rPrDefault', async () => {
    const zipBytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' })
    // buildDocx's styles.xml has no docDefaults: inject it manually
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(zipBytes)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    zip.file(
      'word/styles.xml',
      stylesXml.replace(
        /(<w:styles[^>]*>)/,
        '$1<w:docDefaults><w:rPrDefault><w:rPr>' +
          '<w:b/><w:i w:val="0"/><w:color w:val="333333"/><w:sz w:val="21"/>' +
          '</w:rPr></w:rPrDefault></w:docDefaults>',
      ),
    )
    const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    expect(doc.docDefaults).toMatchObject({ bold: true, color: '333333', sizeHalfPoints: 21 })
    expect(doc.docDefaults?.italic).toBeUndefined()
  })
})

describe('styleUpserts style write-back', () => {
  it('new styles are appended, modified styles are replaced in place', async () => {
    const { saveDocx } = await import('../src/index')
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }),
    )
    const blocks = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, {
      styleUpserts: [
        {
          styleId: 'MyQuote',
          type: 'paragraph',
          name: '我的引用',
          basedOn: 'Normal',
          rPr: { italic: true, color: '595959', sizeHalfPoints: 20 },
          pPr: { align: 'center', spaceBeforeTwips: 120, spaceAfterTwips: 120 },
        },
      ],
    })
    const reparsed = await parseDocx(saved)
    const st = reparsed.styles.get('MyQuote')!
    expect(st.name).toBe('我的引用')
    expect(st.display).toMatchObject({ italic: true, color: '595959', sizeHalfPoints: 20 })
    // Modify: upsert the same styleId again — replaces instead of duplicating
    const parsed2 = await parseDocx(saved)
    const blocks2 = parsed2.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    const saved2 = await saveDocx(parsed2, blocks2, {
      styleUpserts: [
        { styleId: 'MyQuote', type: 'paragraph', name: '我的引用', rPr: { bold: true } },
      ],
    })
    const zip = await (await import('jszip')).default.loadAsync(saved2)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml.match(/w:styleId="MyQuote"/g)).toHaveLength(1)
    const reparsed2 = await parseDocx(saved2)
    expect(reparsed2.styles.get('MyQuote')!.display).toMatchObject({ bold: true })
    expect(reparsed2.styles.get('MyQuote')!.display?.italic).toBeUndefined()
  })
})

describe('toggle-off (w:val="0") overrides inherited formatting', () => {
  const TOGGLE_STYLES =
    '<w:style w:type="paragraph" w:styleId="BoldBase"><w:name w:val="Bold Base"/>' +
    '<w:rPr><w:b/><w:i/><w:strike/><w:u w:val="single"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="PlainChild"><w:name w:val="Plain Child"/><w:basedOn w:val="BoldBase"/>' +
    '<w:rPr><w:b w:val="0"/><w:i w:val="false"/><w:strike w:val="0"/><w:u w:val="none"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="GrandChild"><w:name w:val="Grand Child"/><w:basedOn w:val="PlainChild"/></w:style>'

  it('records explicit off in the style display and through basedOn chains', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraStylesXml: TOGGLE_STYLES,
      }),
    )
    expect(doc.styles.get('BoldBase')!.display).toMatchObject({
      bold: true,
      italic: true,
      strike: true,
      underline: true,
    })
    const child = doc.styles.get('PlainChild')!.display!
    expect(child.bold).toBe(false)
    expect(child.italic).toBe(false)
    expect(child.strike).toBe(false)
    expect(child.underline).toBe(false)
    // the grandchild inherits the explicit off, not the grandparent's on
    const grand = doc.styles.get('GrandChild')!.display!
    expect(grand.bold).toBe(false)
    expect(grand.underline).toBe(false)
  })

  it('records explicit off on runs as false, absent as undefined', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:rPr><w:b w:val="0"/><w:u w:val="none"/></w:rPr><w:t>off</w:t></w:r>' +
          '<w:r><w:rPr><w:b/></w:rPr><w:t>on</w:t></w:r>' +
          '<w:r><w:t>unset</w:t></w:r></w:p>',
      }),
    )
    const runs = doc.blocks[0].runs!
    expect(runs[0].bold).toBe(false)
    expect(runs[0].underline).toBe(false)
    expect(runs[1].bold).toBe(true)
    expect(runs[2].bold).toBeUndefined()
    expect(runs[2].underline).toBeUndefined()
  })
})
