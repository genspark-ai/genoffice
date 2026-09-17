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

  it('captures w:rStyle on runs, the Hyperlink style included', async () => {
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
    expect(runs[2].styleId).toBe('Hyperlink')
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

describe('hyperlink runs keep their own rStyle', () => {
  const LINK_RELS =
    '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>'
  // localized Word builds give the Hyperlink style an opaque id ("ae", "a3")
  const LOCAL_HYPERLINK =
    '<w:style w:type="character" w:styleId="ae"><w:name w:val="Hyperlink"/>' +
    '<w:rPr><w:u w:val="single"/></w:rPr></w:style>'

  it('a localized style id on a link run survives a text edit', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:hyperlink r:id="rId20"><w:r><w:rPr><w:rStyle w:val="ae"/><w:sz w:val="16"/></w:rPr>' +
          '<w:t>https://example.com/</w:t></w:r></w:hyperlink></w:p>',
        extraRels: LINK_RELS,
        extraStylesXml: LOCAL_HYPERLINK,
      }),
    )
    const run = doc.blocks[0].runs![0]
    expect(run.styleId).toBe('ae')
    expect(run.link?.rId).toBe('rId20')
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, text: 'https://example.com/x' }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rStyle w:val="ae"/>')
    expect(xml).not.toContain('Hyperlink')
  })

  it('an unstyled run inside an existing hyperlink stays unstyled', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:hyperlink r:id="rId20"><w:r><w:rPr><w:i/></w:rPr>' +
          '<w:t>mail@example.com</w:t></w:r></w:hyperlink></w:p>',
        extraRels: LINK_RELS,
      }),
    )
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, text: 'mail@example.com [1]' }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:hyperlink r:id="rId20">')
    expect(xml).toContain('<w:rPr><w:i/></w:rPr>')
    expect(xml).not.toContain('<w:rStyle')
  })

  it('a newly linked run gets the implied Hyperlink style', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>plain</w:t></w:r></w:p>' }),
    )
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, link: { href: 'https://new.example/' } }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rStyle w:val="Hyperlink"/>')
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

  it('empty East Asian slot + explicit w:lang w:eastAsia backfills the locale default face', async () => {
    const JSZip = (await import('jszip')).default
    const withDefaults = async (rPr: string) => {
      const zip = await JSZip.loadAsync(
        await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }),
      )
      const stylesXml = await zip.file('word/styles.xml')!.async('string')
      zip.file(
        'word/styles.xml',
        stylesXml.replace(
          /(<w:styles[^>]*>)/,
          `$1<w:docDefaults><w:rPrDefault><w:rPr>${rPr}</w:rPr></w:rPrDefault></w:docDefaults>`,
        ),
      )
      return parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    }
    // theme EA typeface empty → w:eastAsiaTheme resolves to nothing → lang decides
    const rFonts = '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>'
    const ko = await withDefaults(`${rFonts}<w:lang w:val="en-US" w:eastAsia="ko-KR"/>`)
    expect(ko.docDefaults?.eastAsiaFont).toBe('Malgun Gothic')
    const ja = await withDefaults(`${rFonts}<w:lang w:eastAsia="ja-JP"/>`)
    expect(ja.docDefaults?.eastAsiaFont).toBe('MS Mincho')
    const zh = await withDefaults(`${rFonts}<w:lang w:eastAsia="zh-CN"/>`)
    expect(zh.docDefaults?.eastAsiaFont).toBe('SimSun')
    // en-US EA lang or no lang: slot stays empty
    const en = await withDefaults(`${rFonts}<w:lang w:eastAsia="en-US"/>`)
    expect(en.docDefaults?.eastAsiaFont).toBeUndefined()
    const none = await withDefaults(rFonts)
    expect(none.docDefaults?.eastAsiaFont).toBeUndefined()
    // explicit literal face wins over the lang backfill
    const literal = await withDefaults('<w:rFonts w:eastAsia="Gulim"/><w:lang w:eastAsia="ko-KR"/>')
    expect(literal.docDefaults?.eastAsiaFont).toBe('Gulim')
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
    // a re-upsert patches: bold added, the earlier italic/color kept
    expect(reparsed2.styles.get('MyQuote')!.display).toMatchObject({
      bold: true,
      italic: true,
      color: '595959',
    })
  })

  it('patching an existing style keeps the children and attributes it does not name', async () => {
    const { saveDocx, mergeStyleXml } = await import('../src/index')
    const existing =
      '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body Text"/>' +
      '<w:basedOn w:val="Normal"/><w:link w:val="BodyChar"/><w:uiPriority w:val="9"/>' +
      '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120" w:line="276" w:lineRule="auto"/><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
      '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsiaTheme="minorEastAsia"/><w:b/><w:sz w:val="24"/></w:rPr></w:style>'
    const out = mergeStyleXml(existing, {
      styleId: 'Body',
      pPr: { spaceAfterTwips: 0, firstLineTwips: 480, align: 'justify' },
      rPr: { bold: false, italic: true, sizeHalfPoints: 22, eastAsiaFont: 'SimSun' },
    })
    expect(out).toContain('<w:basedOn w:val="Normal"/>')
    expect(out).toContain('<w:link w:val="BodyChar"/>')
    expect(out).toContain('<w:uiPriority w:val="9"/>')
    expect(out).toContain('<w:keepNext/>')
    expect(out).toContain('<w:spacing w:before="240" w:after="0" w:line="276" w:lineRule="auto"/>')
    expect(out).toContain('<w:ind w:left="720" w:firstLine="480"/>')
    expect(out).not.toContain('w:hanging')
    expect(out).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="SimSun"/>')
    expect(out).toContain('<w:b w:val="0"/>')
    expect(out).toContain('<w:i/>')
    expect(out).toContain('<w:sz w:val="22"/><w:szCs w:val="22"/>')
    // schema order: pPr children spacing < ind < jc; rPr b before i before sz
    expect(out.indexOf('<w:spacing')).toBeLessThan(out.indexOf('<w:ind'))
    expect(out.indexOf('<w:ind')).toBeLessThan(out.indexOf('<w:jc'))
    expect(out.indexOf('<w:b ')).toBeLessThan(out.indexOf('<w:i/>'))
    expect(out.indexOf('<w:pPr>')).toBeLessThan(out.indexOf('<w:rPr>'))
    expect(out.indexOf('<w:uiPriority')).toBeLessThan(out.indexOf('<w:pPr>'))

    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }),
    )
    const blocks = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, {
      styleUpserts: [
        { styleId: 'Heading1', pPr: { spaceBeforeTwips: 480 }, rPr: { color: 'FF0000' } },
        {
          styleId: 'Callout',
          type: 'paragraph',
          basedOn: 'Normal',
          next: 'Normal',
          rPr: { italic: true },
        },
      ],
    })
    const reparsed = await parseDocx(saved)
    expect(reparsed.styles.get('Heading1')!.headingLevel).toBe(1)
    expect(reparsed.styles.get('Heading1')!.display?.color).toBe('FF0000')
    expect(reparsed.styles.get('Callout')).toMatchObject({ name: 'Callout', basedOn: 'Normal' })
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml).toContain('<w:next w:val="Normal"/>')
    expect(stylesXml.match(/w:styleId="Heading1"/g)).toHaveLength(1)
  })
})

describe('toggle-off (w:val="0") overrides inherited formatting', () => {
  const TOGGLE_STYLES =
    '<w:style w:type="paragraph" w:styleId="BoldBase"><w:name w:val="Bold Base"/>' +
    '<w:rPr><w:b/><w:i/><w:strike/><w:u w:val="single"/><w:smallCaps/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="PlainChild"><w:name w:val="Plain Child"/><w:basedOn w:val="BoldBase"/>' +
    '<w:rPr><w:b w:val="0"/><w:i w:val="false"/><w:strike w:val="0"/><w:u w:val="none"/><w:smallCaps w:val="0"/></w:rPr></w:style>' +
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
      caps: 'small',
    })
    const child = doc.styles.get('PlainChild')!.display!
    expect(child.bold).toBe(false)
    expect(child.italic).toBe(false)
    expect(child.strike).toBe(false)
    expect(child.underline).toBe(false)
    expect(child.caps).toBe('none')
    // the grandchild inherits the explicit off, not the grandparent's on
    const grand = doc.styles.get('GrandChild')!.display!
    expect(grand.bold).toBe(false)
    expect(grand.underline).toBe(false)
    expect(grand.caps).toBe('none')
  })

  it('records explicit off on runs as false, absent as undefined', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:rPr><w:b w:val="0"/><w:u w:val="none"/><w:caps w:val="0"/></w:rPr><w:t>off</w:t></w:r>' +
          '<w:r><w:rPr><w:b/></w:rPr><w:t>on</w:t></w:r>' +
          '<w:r><w:t>unset</w:t></w:r></w:p>',
      }),
    )
    const runs = doc.blocks[0].runs!
    expect(runs[0].bold).toBe(false)
    expect(runs[0].underline).toBe(false)
    expect(runs[0].caps).toBe('none')
    expect(runs[1].bold).toBe(true)
    expect(runs[2].bold).toBeUndefined()
    expect(runs[2].underline).toBeUndefined()
    expect(runs[2].caps).toBeUndefined()
  })
})

describe('linked styles (w:link)', () => {
  const BODY_TEXT_PAIR =
    '<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:link w:val="BodyTextChar"/></w:style>' +
    '<w:style w:type="character" w:styleId="BodyTextChar"><w:name w:val="Body Text Char"/><w:link w:val="BodyText"/></w:style>'

  it('fills a character shell from its reciprocal paragraph twin', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraStylesXml:
          '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:link w:val="QuoteChar"/><w:rPr><w:i/></w:rPr></w:style>' +
          '<w:style w:type="character" w:styleId="QuoteChar"><w:name w:val="Quote Char"/><w:link w:val="Quote"/></w:style>',
      }),
    )
    expect(doc.styles.get('QuoteChar')!.display?.italic).toBe(true)
  })

  it("ignores a one-way w:link into another style's character twin", async () => {
    // a caption style pointing at Body Text Char must not italicize Body Text
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>plain</w:t></w:r></w:p>',
        extraStylesXml:
          BODY_TEXT_PAIR +
          '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:link w:val="BodyTextChar"/><w:rPr><w:i/></w:rPr></w:style>',
      }),
    )
    expect(doc.styles.get('BodyText')!.display?.italic).toBeUndefined()
    expect(doc.styles.get('BodyTextChar')!.display?.italic).toBeUndefined()
    expect(doc.styles.get('Caption')!.display?.italic).toBe(true)
  })
})
