import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  NEW_DOCUMENT_APPLICATION,
  NEW_DOCUMENT_APP_VERSION,
  NEW_DOCUMENT_TEMPLATE,
  buildBlankDocx,
  paperSizeForLocale,
  paperSizeForRegion,
  parseDocx,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import enginePackage from '../package.json'

describe('blank document paper size follows the OS region like Word', () => {
  it('maps Letter regions and defaults everything else to A4', () => {
    for (const r of [
      'US',
      'CA',
      'MX',
      'PH',
      'LR',
      'MM',
      'PR',
      'CL',
      'CO',
      'VE',
      'GT',
      'CR',
      'PA',
      'DO',
    ])
      expect(paperSizeForRegion(r), r).toBe('Letter')
    expect(paperSizeForRegion('us')).toBe('Letter')
    for (const r of ['GB', 'DE', 'CN', 'JP', 'BR', 'IN', 'AU', '', undefined, null])
      expect(paperSizeForRegion(r), String(r)).toBe('A4')
  })

  it('reads the region out of a BCP 47 system locale, inferring it when absent', () => {
    expect(paperSizeForLocale('en-US')).toBe('Letter')
    expect(paperSizeForLocale('es-MX')).toBe('Letter')
    expect(paperSizeForLocale('en_US')).toBe('Letter')
    expect(paperSizeForLocale('en-GB')).toBe('A4')
    expect(paperSizeForLocale('zh-Hans-CN')).toBe('A4')
    expect(paperSizeForLocale('ja')).toBe('A4')
    expect(paperSizeForLocale('en')).toBe('Letter')
    expect(paperSizeForLocale('')).toBe('A4')
    expect(paperSizeForLocale('not a locale !!')).toBe('A4')
  })

  it('writes the matching w:pgSz and keeps 1-inch margins', async () => {
    const pgSz = async (paperSize?: 'A4' | 'Letter') => {
      const zip = await JSZip.loadAsync(await buildBlankDocx(paperSize ? { paperSize } : undefined))
      const xml = await zip.file('word/document.xml')!.async('string')
      expect(xml).toContain('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"')
      return xml.match(/<w:pgSz [^>]*\/>/)![0]
    }
    expect(await pgSz()).toBe('<w:pgSz w:w="11906" w:h="16838"/>')
    expect(await pgSz('A4')).toBe('<w:pgSz w:w="11906" w:h="16838"/>')
    expect(await pgSz('Letter')).toBe('<w:pgSz w:w="12240" w:h="15840"/>')
  })
})

describe('blank document template (new-document / AI generation base)', () => {
  it('parses to a single empty paragraph with standard heading styles available', async () => {
    const bytes = await buildBlankDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible).toHaveLength(1)
    expect(visible[0].type).toBe('paragraph')
    for (const styleId of ['Heading1', 'Heading2', 'Heading3', 'Heading6']) {
      expect(doc.styles.has(styleId), styleId).toBe(true)
    }
  })

  it('ships Word\u2019s quick style set with gallery flags and resolved theme colors', async () => {
    const doc = await parseDocx(await buildBlankDocx())
    const quick = [...doc.styles.values()].filter((s) => s.qFormat)
    const ids = quick.map((s) => s.styleId)
    for (const id of [
      'NoSpacing',
      'Title',
      'Subtitle',
      'SubtleEmphasis',
      'Emphasis',
      'IntenseEmphasis',
      'Strong',
      'Quote',
      'IntenseQuote',
      'SubtleReference',
      'IntenseReference',
      'BookTitle',
      'ListParagraph',
    ])
      expect(ids, id).toContain(id)
    expect(doc.styles.get('Heading2')).toMatchObject({ uiPriority: 9, semiHidden: undefined })
    expect(doc.styles.get('Heading3')).toMatchObject({
      uiPriority: 9,
      semiHidden: true,
      unhideWhenUsed: true,
    })
    expect(doc.styles.get('Title')?.display).toMatchObject({
      sizeHalfPoints: 56,
      spaceAfterTwips: 80,
      charSpacingTwips: -10,
    })
    expect(doc.styles.get('IntenseQuote')?.display).toMatchObject({
      italic: true,
      color: '2F5496',
      align: 'center',
    })
    expect(doc.styles.get('IntenseReference')?.display).toMatchObject({
      bold: true,
      caps: 'small',
      charSpacingTwips: 5,
    })
    expect(doc.styles.get('Strong')?.custom).toBeUndefined()
  })

  it('AI-generated content (headings, paragraphs, lists) saves and reparses correctly', async () => {
    const bytes = await buildBlankDocx()
    const doc = await parseDocx(bytes)

    const generated: SaveBlock[] = [
      { kind: 'generated', block: { type: 'heading', level: 1, runs: [{ text: '生成的标题' }] } },
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [
            { text: '\u6b63\u6587\u6bb5\u843d,', bold: false },
            { text: '\u91cd\u70b9', bold: true },
          ],
        },
      },
      {
        kind: 'generated',
        block: {
          type: 'listItem',
          list: { kind: 'bullet', numId: BLANK_BULLET_NUM_ID, ilvl: 0 },
          runs: [{ text: '无序项' }],
        },
      },
      {
        kind: 'generated',
        block: {
          type: 'listItem',
          list: { kind: 'ordered', numId: BLANK_ORDERED_NUM_ID, ilvl: 0 },
          runs: [{ text: '有序项' }],
        },
      },
    ]
    const saved = await saveDocx(doc, generated)
    const reparsed = await parseDocx(saved)
    const visible = reparsed.blocks.filter((b) => !b.hidden)

    expect(visible.map((b) => b.type)).toEqual(['heading', 'paragraph', 'listItem', 'listItem'])
    expect(visible[0].level).toBe(1)
    expect(visible[0].runs?.map((r) => r.text).join('')).toBe('生成的标题')
    expect(visible[1].runs?.some((r) => r.bold && r.text === '重点')).toBe(true)
    expect(visible[2].list?.kind).toBe('bullet')
    expect(visible[3].list?.kind).toBe('ordered')
  })

  // Word always writes docProps/core.xml and docProps/app.xml; a newly created
  // document carries minimal Word-like ones so later property sets have a home
  const CREATED = new Date('2026-01-02T03:04:05Z')

  it('writes minimal Word-like core and app parts', async () => {
    const zip = await JSZip.loadAsync(await buildBlankDocx({ createdAt: CREATED }))
    const core = await zip.file('docProps/core.xml')!.async('string')
    expect(core).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-02T03:04:05Z</dcterms:created>',
    )
    expect(core).toContain(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-02T03:04:05Z</dcterms:modified>',
    )
    expect(core).toContain('<cp:revision>1</cp:revision>')
    // no people data on a fresh document
    expect(core).not.toContain('<dc:creator')
    expect(core).not.toContain('lastModifiedBy')

    const app = await zip.file('docProps/app.xml')!.async('string')
    expect(app).toContain(`<Template>${NEW_DOCUMENT_TEMPLATE}</Template>`)
    expect(app).toContain(`<Application>${NEW_DOCUMENT_APPLICATION}</Application>`)
    expect(app).toContain(`<AppVersion>${NEW_DOCUMENT_APP_VERSION}</AppVersion>`)
    expect(app).not.toContain('Company')
    expect(app).not.toContain('Manager')
  })

  it('registers the parts in [Content_Types].xml and the root relationships', async () => {
    const zip = await JSZip.loadAsync(await buildBlankDocx({ createdAt: CREATED }))
    const types = await zip.file('[Content_Types].xml')!.async('string')
    expect(types).toContain(
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    )
    expect(types).toContain(
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    )
    const rels = await zip.file('_rels/.rels')!.async('string')
    expect(rels).toContain(
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    )
    expect(rels).toContain(
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
    )
  })

  it('parses back as document properties, keeping the app version pinned to the package', async () => {
    expect(NEW_DOCUMENT_APP_VERSION).toBe(enginePackage.version)
    const doc = await parseDocx(await buildBlankDocx({ createdAt: CREATED }))
    expect(doc.properties.core).toEqual({
      created: '2026-01-02T03:04:05Z',
      modified: '2026-01-02T03:04:05Z',
      revision: '1',
    })
    expect(doc.properties.app).toEqual({
      template: NEW_DOCUMENT_TEMPLATE,
      application: NEW_DOCUMENT_APPLICATION,
      appVersion: NEW_DOCUMENT_APP_VERSION,
    })
    expect(doc.properties.custom).toEqual({})
    expect(doc.properties.customTypes).toEqual({})
  })

  it('a real save stamps modified/revision and keeps the parts', async () => {
    const doc = await parseDocx(await buildBlankDocx({ createdAt: CREATED }))
    const saved = await saveDocx(
      doc,
      [{ kind: 'generated', block: { type: 'paragraph', runs: [{ text: 'hi' }] } }],
      { savedAt: '2026-07-28T08:30:00Z' },
    )
    const zip = await JSZip.loadAsync(saved)
    const core = await zip.file('docProps/core.xml')!.async('string')
    expect(core).toContain(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-28T08:30:00Z</dcterms:modified>',
    )
    expect(core).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-02T03:04:05Z</dcterms:created>',
    )
    expect(core).toContain('<cp:revision>2</cp:revision>')
    expect(zip.file('docProps/app.xml')).not.toBeNull()
  })

  it('lets a caller override the application stamp', async () => {
    const zip = await JSZip.loadAsync(
      await buildBlankDocx({
        createdAt: CREATED,
        application: 'GenOffice Docs',
        appVersion: '9.9.9',
      }),
    )
    const app = await zip.file('docProps/app.xml')!.async('string')
    expect(app).toContain('<Application>GenOffice Docs</Application>')
    expect(app).toContain('<AppVersion>9.9.9</AppVersion>')
  })
})
