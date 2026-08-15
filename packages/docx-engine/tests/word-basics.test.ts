import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  DEFAULT_SECTION,
  generateParagraphXml,
  parseDocx,
  readSections,
  readSectionSettings,
  saveDocx,
  sectionSettingsFromXml,
  type GenerateContext,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const CTX: GenerateContext = {
  headingStyleIds: new Map(),
  allocateHyperlinkRel: () => 'rId999',
}

function originals(indexes: number[]): SaveBlock[] {
  return indexes.map((docxIndex) => ({ kind: 'original', docxIndex }))
}

describe('paragraph shading and borders', () => {
  it('round-trips shading fill and border sides through generate + parse', async () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        format: { shadingFill: 'FFF2CC', borders: 'tb' },
        runs: [{ text: '底纹段落' }],
      },
      CTX,
    )
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>')
    expect(xml).toContain('<w:pBdr><w:top')
    expect(xml).toContain('<w:bottom')
    expect(xml).not.toContain('<w:left')

    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].format?.shadingFill).toBe('FFF2CC')
    expect(doc.blocks[0].format?.borders).toBe('tb')
  })

  it('keeps declared pBdr color and width (w:sz eighth-points -> pt)', async () => {
    const xml =
      '<w:p><w:pPr><w:pBdr>' +
      '<w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/>' +
      '<w:bottom w:val="single" w:sz="12" w:space="1" w:color="0B5394"/>' +
      '</w:pBdr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].format?.borders).toBe('tb')
    expect(doc.blocks[0].format?.borderLines).toEqual({
      t: { szPt: 0.5 },
      b: { color: '0B5394', szPt: 1.5 },
    })
  })

  it('a later duplicated w:pBdr overrides a side declared look', async () => {
    const xml =
      '<w:p><w:pPr>' +
      '<w:pBdr><w:bottom w:val="single" w:sz="4" w:color="FF0000"/></w:pBdr>' +
      '<w:pBdr><w:bottom w:val="single" w:sz="24" w:color="0B5394"/></w:pBdr>' +
      '</w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].format?.borders).toBe('b')
    expect(doc.blocks[0].format?.borderLines).toEqual({ b: { color: '0B5394', szPt: 3 } })
  })
})

describe('page border and columns (sectPr)', () => {
  it('adds and reads a page border box', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }))
    const saved = await saveDocx(doc, originals([0]), {
      section: { ...readSectionSettings(doc), pageBorder: true },
    })
    const reparsed = await parseDocx(saved)
    expect(readSectionSettings(reparsed).pageBorder).toBe(true)
    expect(reparsed.internal.documentXml).toContain('<w:pgBorders w:offsetFrom="page">')

    // and removing it again
    const cleared = await saveDocx(reparsed, originals([0]), {
      section: { ...readSectionSettings(reparsed), pageBorder: false },
    })
    expect(readSectionSettings(await parseDocx(cleared)).pageBorder).toBe(false)
  })

  it('sets and clears column count', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }))
    const saved = await saveDocx(doc, originals([0]), {
      section: { ...readSectionSettings(doc), columns: 2 },
    })
    const reparsed = await parseDocx(saved)
    expect(readSectionSettings(reparsed).columns).toBe(2)

    const back = await saveDocx(reparsed, originals([0]), {
      section: { ...readSectionSettings(reparsed), columns: 1 },
    })
    expect(readSectionSettings(await parseDocx(back)).columns).toBe(1)
  })

  it('DEFAULT_SECTION carries the new fields', () => {
    expect(DEFAULT_SECTION.pageBorder).toBe(false)
    expect(DEFAULT_SECTION.columns).toBe(1)
  })

  // Government documents write explicit "no border" as pgBorders with all sides
  // val="none"; the presence of the element alone must not draw a border
  const NONE_PG_BORDERS =
    '<w:pgBorders>' +
    '<w:top w:val="none" w:sz="0" w:space="0"/>' +
    '<w:left w:val="none" w:sz="0" w:space="0"/>' +
    '<w:bottom w:val="none" w:sz="0" w:space="0"/>' +
    '<w:right w:val="none" w:sz="0" w:space="0"/>' +
    '</w:pgBorders>'

  it('pgBorders with all sides val="none" yields pageBorder=false', () => {
    const sectPr = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>${NONE_PG_BORDERS}</w:sectPr>`
    expect(sectionSettingsFromXml(sectPr).pageBorder).toBe(false)
  })

  it('pgBorders side values nil / missing are not visible; any drawn side is', () => {
    const wrap = (sides: string) => `<w:sectPr><w:pgBorders>${sides}</w:pgBorders></w:sectPr>`
    expect(sectionSettingsFromXml(wrap('<w:top w:val="nil"/>')).pageBorder).toBe(false)
    expect(sectionSettingsFromXml(wrap('')).pageBorder).toBe(false)
    expect(
      sectionSettingsFromXml(wrap('<w:top w:val="none"/><w:left w:val="single" w:sz="4"/>'))
        .pageBorder,
    ).toBe(true)
    expect(
      sectionSettingsFromXml(
        wrap(
          '<w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
            '<w:left w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
            '<w:bottom w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
            '<w:right w:val="single" w:sz="4" w:space="24" w:color="auto"/>',
        ),
      ).pageBorder,
    ).toBe(true)
  })

  it('sections whose pgBorders declare all sides none parse without a page border', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      sectPrExtra: NONE_PG_BORDERS,
    })
    const parsed = await parseDocx(bytes)
    expect(readSectionSettings(parsed).pageBorder).toBe(false)
    const sections = readSections(parsed)
    expect(sections).toHaveLength(1)
    expect(sections[0].settings.pageBorder).toBe(false)
  })
})

describe('header / footer', () => {
  it('creates default header and footer parts with references and content types', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }))
    expect(doc.headerText).toBeNull()

    const saved = await saveDocx(doc, originals([0]), {
      header: { text: '机密文件' },
      footer: { text: '第', pageNumber: true },
    })

    const zip = await JSZip.loadAsync(saved)
    const headerXml = await zip.file('word/header1.xml')!.async('string')
    const footerXml = await zip.file('word/footer1.xml')!.async('string')
    expect(headerXml).toContain('机密文件')
    expect(footerXml).toContain(' PAGE ')
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).toMatch(
      /<w:sectPr[^>]*><w:headerReference w:type="default" r:id="rId\d+"\/><w:footerReference/,
    )
    const ct = await zip.file('[Content_Types].xml')!.async('string')
    expect(ct).toContain('/word/header1.xml')
    expect(ct).toContain('header+xml')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('Target="header1.xml"')

    // reparse exposes the text for display
    const reparsed = await parseDocx(saved)
    expect(reparsed.headerText).toBe('机密文件')
    expect(reparsed.footerHasPageNumber).toBe(true)
  })

  it('overwrites an existing default header instead of adding a second one', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }))
    const first = await saveDocx(doc, originals([0]), { header: { text: '版本一' } })
    const second = await saveDocx(await parseDocx(first), originals([0]), {
      header: { text: '版本二' },
    })

    const zip = await JSZip.loadAsync(second)
    expect(zip.file('word/header2.xml')).toBeNull()
    expect(await zip.file('word/header1.xml')!.async('string')).toContain('版本二')
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml.match(/<w:headerReference/g)).toHaveLength(1)
  })
})

describe('protection password hash', () => {
  it('round-trips after setting a password: the correct password verifies, a wrong one does not', async () => {
    const { hashProtectionPassword, verifyProtectionPassword, parseDocx, saveDocx } =
      await import('../src/index')
    const { buildDocx } = await import('./helpers/build-docx')
    const creds = await hashProtectionPassword('s3cret', 1000) // lower iterations for the test
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }),
    )
    const blocks = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, {
      protection: { edit: 'readOnly', enforced: true, ...creds },
    })
    const reparsed = await parseDocx(saved)
    expect(reparsed.protection).toMatchObject({
      edit: 'readOnly',
      enforced: true,
      spinCount: 1000,
      algorithmSid: 14,
    })
    expect(reparsed.protection?.hash).toBe(creds.hash)
    expect(await verifyProtectionPassword('s3cret', reparsed.protection!)).toBe(true)
    expect(await verifyProtectionPassword('wrong', reparsed.protection!)).toBe(false)
    // settings.xml carries the full field set
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const settings = await zip.file('word/settings.xml')!.async('string')
    expect(settings).toContain('w:cryptAlgorithmSid="14"')
    expect(settings).toContain('w:cryptSpinCount="1000"')
  })
})
