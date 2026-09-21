import { describe, expect, it } from 'vitest'
import { mergeStyleXml } from '../src/style-upsert'

describe('style font size validation', () => {
  it('clamps non-finite and out-of-range half-point sizes to the Word range', () => {
    const sz = (v: number) =>
      mergeStyleXml(null, { styleId: 'Normal', rPr: { sizeHalfPoints: v } }).match(
        /w:sz w:val="(\d+)"/,
      )?.[1]
    for (const low of [NaN, -Infinity, -10, 0, 1]) expect(sz(low)).toBe('2')
    for (const high of [Infinity, 3277, 1e9]) expect(sz(high)).toBe('3276')
    expect(sz(3276)).toBe('3276')
    expect(sz(3200)).toBe('3200')
  })

  it('accepts normal sizes', () => {
    const xml = mergeStyleXml(null, { styleId: 'Normal', rPr: { sizeHalfPoints: 24 } })
    expect(xml).toContain('w:val="24"')
  })
})

describe('style font slot isolation', () => {
  it('editing Latin font keeps complex script and East Asian theme references', () => {
    const xml = mergeStyleXml(
      '<w:style w:type="paragraph" w:styleId="Normal"><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:b/></w:rPr></w:style>',
      { styleId: 'Normal', rPr: { font: 'Times New Roman' } },
    )
    expect(xml).toContain('w:cstheme="minorBidi"')
    expect(xml).toContain('w:eastAsiaTheme="minorEastAsia"')
    expect(xml).toContain('w:ascii="Times New Roman"')
    expect(xml).toContain('<w:b/>')
    expect(xml).not.toContain('w:asciiTheme')
  })
})

import { buildDocx } from './helpers/build-docx'
import { parseDocx, saveDocx } from '../src/index'
import JSZip from 'jszip'
it('saves document default fonts independently without changing paragraph defaults', async () => {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>\u4e2d\u6587 English 123</w:t></w:r></w:p>',
  })
  const zip = await JSZip.loadAsync(bytes)
  zip.file(
    'word/styles.xml',
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="SimSun" w:cs="Amiri"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>',
  )
  const input = await zip.generateAsync({ type: 'uint8array' })
  const parsed = await parseDocx(input)
  const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
    defaultFonts: { font: 'Times New Roman' },
  })
  const reopened = await parseDocx(saved)
  expect(reopened.docDefaults).toMatchObject({
    asciiFont: 'Times New Roman',
    eastAsiaFont: 'SimSun',
    sizeHalfPoints: 24,
    spaceAfterTwips: 160,
  })
  expect(await (await JSZip.loadAsync(saved)).file('word/styles.xml')!.async('string')).toContain(
    'w:cs="Amiri"',
  )
})
it('records an explicit East Asian choice equal to the previous Latin fallback', async () => {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>\u4e2d\u6587 English</w:t></w:r></w:p>',
    }),
  )
  const run = parsed.blocks[0].runs![0]
  const saved = await saveDocx(parsed, [
    {
      kind: 'generated',
      block: { type: 'paragraph', runs: [{ ...run, font: 'Arial', eastAsiaFont: 'Arial' }] },
    },
  ])
  const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
  expect(xml).toContain('w:eastAsia="Arial"')
})

import { previewFontSettings } from '../src/font-settings'
it('previews and persists per-slot inheritance for both style types', async () => {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>\u4e2d\u6587 English 123</w:t></w:r></w:p>',
      stylesXml:
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:rFonts w:eastAsia="SimSun" w:ascii="Arial"/><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style><w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:i/></w:rPr></w:style></w:styles>',
    }),
  )
  const upserts = [
    { styleId: 'Base', rPr: { font: 'Times New Roman' } },
    { styleId: 'Emphasis', rPr: { eastAsiaFont: 'KaiTi' } },
  ]
  const preview = await previewFontSettings(parsed, upserts, { eastAsiaFont: 'SimHei' })
  expect(preview.styles.get('Child')?.display).toMatchObject({
    eastAsiaFont: 'SimSun',
    fontAscii: 'Times New Roman',
    bold: true,
  })
  expect(preview.styles.get('Emphasis')?.display).toMatchObject({
    eastAsiaFont: 'KaiTi',
    italic: true,
  })
  const reopened = await parseDocx(
    await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      styleUpserts: upserts,
      defaultFonts: { eastAsiaFont: 'SimHei' },
    }),
  )
  expect(preview.styles).toEqual(reopened.styles)
  expect(preview.docDefaults).toEqual(reopened.docDefaults)
})
it('an East Asian-only edit does not materialize inherited Latin or complex-script slots', async () => {
  const parsed = await parseDocx(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>\u4e2d\u6587 English 123</w:t></w:r></w:p>' }),
  )
  const saved = await saveDocx(parsed, [
    {
      kind: 'generated',
      block: {
        type: 'paragraph',
        runs: [{ ...parsed.blocks[0].runs![0], font: 'KaiTi', eastAsiaFont: 'KaiTi' }],
      },
    },
  ])
  const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
  expect(xml).toContain('w:eastAsia="KaiTi"')
  expect(xml).not.toMatch(/w:(ascii|hAnsi|cs)=/)
})
