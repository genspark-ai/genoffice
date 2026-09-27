import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  paperSizeForLocale,
  paperSizeForRegion,
  parseDocx,
  saveDocx,
  type SaveBlock,
} from '../src/index'

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
})
