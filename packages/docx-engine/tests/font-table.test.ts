/**
 * word/fontTable.xml substitution hints + empty-EA-slot resolution priority
 * (theme script table / themeFontLang vs the docDefaults w:lang backfill).
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, parseFontTable } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const FONT_TABLE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:font w:name="원신한 Light"><w:panose1 w:val="020B0303000000000000"/>' +
  '<w:charset w:val="81"/><w:family w:val="modern"/><w:pitch w:val="variable"/></w:font>' +
  '<w:font w:name="汉仪旗黑-50简"><w:altName w:val="黑体"/>' +
  '<w:panose1 w:val="00000000000000000000"/><w:family w:val="auto"/><w:pitch w:val="default"/></w:font>' +
  '<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>' +
  '</w:fonts>'

const FONT_TABLE_PART = {
  path: 'word/fontTable.xml',
  xml: FONT_TABLE_XML,
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
}

describe('parseFontTable', () => {
  it('reads name / altName / panose / family / pitch per entry', () => {
    const entries = parseFontTable(FONT_TABLE_XML)
    expect(entries).toEqual([
      { name: '원신한 Light', panose: '020B0303000000000000', family: 'modern', pitch: 'variable' },
      {
        name: '汉仪旗黑-50简',
        altName: '黑体',
        panose: '00000000000000000000',
        family: 'auto',
        pitch: 'default',
      },
      { name: 'Calibri', panose: '020F0502020204030204' },
    ])
  })

  it('degrades to empty on malformed or rootless xml', () => {
    expect(parseFontTable('<w:font w:name="X"/>')).toEqual([])
    expect(parseFontTable('not xml <')).toEqual([])
  })
})

describe('ParsedDoc.fontTable', () => {
  it('exposes word/fontTable.xml entries', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraParts: [FONT_TABLE_PART],
      }),
    )
    expect(doc.fontTable?.map((e) => e.name)).toEqual(['원신한 Light', '汉仪旗黑-50简', 'Calibri'])
    expect(doc.fontTable?.[1].altName).toBe('黑体')
  })

  it('is absent when the doc has no fontTable part', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }))
    expect(doc.fontTable).toBeUndefined()
  })
})

const themePart = (minorFontExtra: string) => ({
  path: 'word/theme/theme1.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
    '<a:themeElements><a:fontScheme name="T">' +
    '<a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    `<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/>${minorFontExtra}</a:minorFont>` +
    '</a:fontScheme></a:themeElements></a:theme>',
  contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
})

const settingsPart = (eaLang: string) => ({
  path: 'word/settings.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:themeFontLang w:val="en-US" w:eastAsia="${eaLang}"/></w:settings>`,
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
})

const HANG_SCRIPT = '<a:font script="Hang" typeface="맑은 고딕"/>'

/** docDefaults with an empty-EA theme ref and a stale w:lang */
async function withDocDefaults(
  rPr: string,
  extraParts: Parameters<typeof buildDocx>[0]['extraParts'],
) {
  const zip = await JSZip.loadAsync(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>', extraParts }),
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

describe('empty EA slot resolution priority', () => {
  const RFONTS = '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>'

  it('theme script table entry wins for the themeFontLang script', async () => {
    const para =
      '<w:p><w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/></w:rPr><w:t>본문</w:t></w:r></w:p>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: para,
        extraParts: [themePart(HANG_SCRIPT), settingsPart('ko-KR')],
      }),
    )
    expect(doc.blocks[0].runs![0].font).toBe('맑은 고딕')
  })

  it('themeFontLang outranks a stale docDefaults w:lang backfill', async () => {
    const doc = await withDocDefaults(`${RFONTS}<w:lang w:val="en-US" w:eastAsia="zh-CN"/>`, [
      themePart(HANG_SCRIPT),
      settingsPart('ko-KR'),
    ])
    expect(doc.docDefaults?.eastAsiaFont).toBe('맑은 고딕')
    expect(doc.docDefaults?.eaSlotEmpty).toBe(true)
  })

  it('without a script table the probed locale default still outranks w:lang', async () => {
    const doc = await withDocDefaults(`${RFONTS}<w:lang w:eastAsia="zh-CN"/>`, [
      themePart(''),
      settingsPart('ko-KR'),
    ])
    expect(doc.docDefaults?.eastAsiaFont).toBe('Malgun Gothic')
  })

  it('without themeFontLang the w:lang backfill keeps deciding', async () => {
    const doc = await withDocDefaults(`${RFONTS}<w:lang w:eastAsia="zh-CN"/>`, [
      themePart(HANG_SCRIPT),
    ])
    expect(doc.docDefaults?.eastAsiaFont).toBe('SimSun')
    expect(doc.docDefaults?.eaFromLang).toBe(true)
  })

  it('a literal declared EA face is not flagged as a lang backfill', async () => {
    const doc = await withDocDefaults('<w:rFonts w:eastAsia="Noto Sans CJK KR"/>', [])
    expect(doc.docDefaults?.eastAsiaFont).toBe('Noto Sans CJK KR')
    expect(doc.docDefaults?.eaFromLang).toBeUndefined()
  })

  it('themeFontLang alone must not invent an EA face when w:lang would not backfill', async () => {
    const enLang = await withDocDefaults(`${RFONTS}<w:lang w:val="en-US" w:eastAsia="en-US"/>`, [
      themePart(HANG_SCRIPT),
      settingsPart('ja-JP'),
    ])
    expect(enLang.docDefaults?.eastAsiaFont).toBeUndefined()
    const noLang = await withDocDefaults(RFONTS, [themePart(HANG_SCRIPT), settingsPart('ja-JP')])
    expect(noLang.docDefaults?.eastAsiaFont).toBeUndefined()
  })
})
