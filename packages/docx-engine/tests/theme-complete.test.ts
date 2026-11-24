import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, previewTheme, saveDocx, type SaveBlock } from '../src/index'
import {
  THEME_CONTENT_TYPE,
  THEME_REL_TYPE,
  buildThemeXml,
  completeThemeElements,
  rewriteThemeXml,
} from '../src/theme'
import { buildDocx } from './helpers/build-docx'

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const CLR =
  '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  ['dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink']
    .map((tag) => `<a:${tag}><a:srgbClr val="112233"/></a:${tag}>`)
    .join('') +
  '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
const FONTS =
  '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Georgia"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
const theme = (elements: string) =>
  `<?xml version="1.0" encoding="UTF-8"?><a:theme ${A} name="Office"><a:themeElements>${elements}</a:themeElements><a:objectDefaults/></a:theme>`

/** what CT_BaseStyles asks for: the three schemes in order, the format scheme with its four lists */
const COMPLETE =
  /<a:themeElements><a:clrScheme [\s\S]*<a:dk1>[\s\S]*<\/a:clrScheme><a:fontScheme [\s\S]*<a:majorFont>[\s\S]*<\/a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst>(?:<a:(?:solid|grad)Fill[\s\S]*){3}<\/a:fillStyleLst><a:lnStyleLst>(?:<a:ln [\s\S]*){3}<\/a:lnStyleLst><a:effectStyleLst>(?:<a:effectStyle>[\s\S]*){3}<\/a:effectStyleLst><a:bgFillStyleLst>[\s\S]*<\/a:bgFillStyleLst><\/a:fmtScheme><\/a:themeElements>/

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

const withTheme = (xml: string) =>
  buildDocx({
    bodyXml: '<w:p><w:r><w:t>text</w:t></w:r></w:p>',
    extraRels: `<Relationship Id="rIdTheme9" Type="${THEME_REL_TYPE}" Target="theme/theme1.xml"/>`,
    extraParts: [{ path: 'word/theme/theme1.xml', xml, contentType: THEME_CONTENT_TYPE }],
  })

const themeOf = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('word/theme/theme1.xml')!.async('string')

describe('a theme part Word opens', () => {
  it('fills in the format scheme a theme lacks and keeps what it has', () => {
    const bare = theme(CLR + FONTS)
    const done = completeThemeElements(bare)
    expect(done).toMatch(COMPLETE)
    expect(done).toContain(CLR)
    expect(done).toContain(FONTS)
    expect(done).toContain('</a:themeElements><a:objectDefaults/>')
    expect(completeThemeElements(done)).toBe(done)
  })

  it('fills in a missing or empty color or font scheme in schema order', () => {
    for (const elements of [
      FONTS,
      CLR,
      '',
      '<a:clrScheme name="x"/><a:fontScheme name="y"/><a:fmtScheme name="z"/>',
      '<a:clrScheme name="x"><a:accent1><a:srgbClr val="112233"/></a:accent1></a:clrScheme>' +
        '<a:fontScheme name="y"></a:fontScheme><a:fmtScheme name="z"><a:lnStyleLst/></a:fmtScheme>',
    ]) {
      const done = completeThemeElements(theme(elements))
      expect(done).toMatch(COMPLETE)
      for (const tag of ['a:clrScheme', 'a:fontScheme', 'a:fmtScheme'])
        expect(done.split(`<${tag}`)).toHaveLength(2)
    }
    expect(completeThemeElements(theme(FONTS))).toContain(FONTS)
    expect(
      completeThemeElements(`<a:theme ${A} name="Office"><a:themeElements/></a:theme>`),
    ).toMatch(COMPLETE)
  })

  it('leaves a complete theme and a theme under another prefix as they are', () => {
    const full = buildThemeXml({ major: 'Cambria', minor: 'Georgia' }, {})
    expect(full).toMatch(COMPLETE)
    expect(completeThemeElements(full)).toBe(full)
    const other = full.replace(/<(\/?)a:/g, '<$1d:').replace('xmlns:a=', 'xmlns:d=')
    expect(completeThemeElements(other)).toBe(other)
  })

  it('a rewrite applies the pick to the completed theme', () => {
    const out = rewriteThemeXml(
      theme(FONTS),
      { major: 'Trebuchet MS', minor: 'Trebuchet MS' },
      {
        accent1: '90C226',
      },
    )
    expect(out).toMatch(COMPLETE)
    expect(out).toContain('<a:accent1><a:srgbClr val="90C226"/></a:accent1>')
    expect(out).toContain('<a:latin typeface="Trebuchet MS"/>')
  })

  it('a save that changes the theme writes a complete one; a save that does not leaves the part alone', async () => {
    const bare = theme(CLR + FONTS)
    const doc = await parseDocx(await withTheme(bare))
    const picked = await saveDocx(doc, originalOrder(doc), { themeColors: { accent1: '90C226' } })
    const saved = await themeOf(picked)
    expect(saved).toMatch(COMPLETE)
    expect(saved).toContain('<a:accent1><a:srgbClr val="90C226"/></a:accent1>')
    expect((await parseDocx(picked)).themeColors?.accent1).toBe('90C226')

    const untouched = await saveDocx(doc, originalOrder(doc), {})
    expect(await themeOf(untouched)).toBe(bare)
  })

  it('the preview of a pick reads the theme the save will write', () => {
    const parsed = { themeFonts: null, themeColors: null }
    const live = previewTheme(theme(FONTS), parsed, { colors: { accent1: '90C226' } })
    expect(live.themeColors?.accent1).toBe('90C226')
    expect(live.themeColors?.dk2).toBe('44546A')
    expect(live.themeFonts?.minor).toBe('Georgia')
  })
})
