import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  OFFICE_THEMES,
  OFFICE_THEME_COLORS,
  OFFICE_THEME_FONTS,
  editedThemeXml,
  parseDocx,
  readThemeColors,
  readThemeFonts,
  saveDocx,
  themedDocxBytes,
  type Block,
  type ParsedDocFull,
  type ThemeEdit,
} from '../src/index'
import { OFFICE_THEME_XML, themedDocx } from './helpers/themed-docx'
import { xmlProblem } from './helpers/strict-xml'

const SLOTS = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const
const HEX = /^[0-9A-F]{6}$/
const count = (xml: string, tag: string) => xml.split(`<${tag}`).length - 1
const children = (xml: string, list: string) => {
  const inner = new RegExp(`<a:${list}>([\\s\\S]*?)</a:${list}>`).exec(xml)?.[1] ?? ''
  let depth = 0
  let n = 0
  for (const m of inner.matchAll(/<(\/?)a:[A-Za-z]+[^>]*?(\/?)>/g)) {
    if (m[1]) depth--
    else {
      if (depth === 0) n++
      if (!m[2]) depth++
    }
  }
  return n
}

describe("Word's Office themes as data", () => {
  it('lists the galleries in the order Word shows them', () => {
    expect(OFFICE_THEMES.map((t) => t.name)).toEqual([
      'Office',
      'Office 2013 - 2022 Theme',
      'Atlas',
      'Badge',
      'Berlin',
      'Celestial',
      'Crop',
      'Depth',
      'Droplet',
      'Facet',
      'Feathered',
      'Gallery',
      'Headlines',
      'Integral',
      'Ion',
      'Ion Boardroom',
      'Madison',
      'Main Event',
      'Mesh',
      'Organic',
      'Parallax',
      'Parcel',
      'Retrospect',
      'Savon',
      'Slice',
      'Vapor Trail',
      'View',
      'Wisp',
      'Wood Type',
    ])
    expect(OFFICE_THEME_COLORS).toHaveLength(24)
    expect(OFFICE_THEME_FONTS).toHaveLength(26)
    expect(OFFICE_THEME_COLORS.slice(0, 4).map((c) => c.name)).toEqual([
      'Office',
      'Office 2013 - 2022',
      'Office 2007 - 2010',
      'Grayscale',
    ])
    expect(OFFICE_THEME_FONTS.slice(0, 4).map((f) => f.name)).toEqual([
      'Office',
      'Office 2013 - 2022',
      'Office 2007 - 2010',
      'Calibri',
    ])
    for (const list of [OFFICE_THEMES, OFFICE_THEME_COLORS, OFFICE_THEME_FONTS]) {
      expect(new Set(list.map((e) => e.id)).size).toBe(list.length)
      for (const { id } of list) expect(id).toMatch(/^[a-z][A-Za-z0-9]*$/)
    }
    expect(OFFICE_THEME_COLORS[0]!.colors).toEqual(OFFICE_THEMES[0]!.colors)
    expect(OFFICE_THEME_FONTS[0]!.fonts).toEqual(OFFICE_THEMES[0]!.fonts)
  })

  for (const { name, colors } of [...OFFICE_THEMES, ...OFFICE_THEME_COLORS]) {
    it(`${name}: a full colour scheme`, () => {
      expect(colors.name).toBe(name)
      for (const slot of SLOTS) expect(colors[slot]).toMatch(HEX)
    })
  }

  for (const { name, fonts } of [...OFFICE_THEMES, ...OFFICE_THEME_FONTS]) {
    it(`${name}: a full font scheme`, () => {
      expect(fonts.major).toBeTruthy()
      expect(fonts.minor).toBeTruthy()
      // a scheme without per-script faces says so, and a pick clears the document's
      for (const scripts of [fonts.majorScripts, fonts.minorScripts]) {
        expect(scripts).toBeDefined()
        for (const face of Object.values(scripts!)) expect(face).toBeTruthy()
      }
    })
  }

  for (const theme of OFFICE_THEMES) {
    it(`${theme.name}: a format scheme Word accepts, and a theme part that reads back as the data`, () => {
      const { format } = theme
      expect(format.startsWith(`<a:fmtScheme name="`)).toBe(true)
      expect(format.endsWith('</a:fmtScheme>')).toBe(true)
      expect(children(format, 'fillStyleLst')).toBe(3)
      expect(children(format, 'lnStyleLst')).toBe(3)
      expect(children(format, 'effectStyleLst')).toBe(3)
      expect(children(format, 'bgFillStyleLst')).toBe(3)
      // a picture fill would point at a part the document does not have
      expect(format).not.toMatch(/blip|r:embed|r:id/)
      const edit = {
        colors: theme.colors,
        fonts: theme.fonts,
        format: { name: theme.name, xml: format },
      }
      for (const xml of [editedThemeXml(OFFICE_THEME_XML, edit), editedThemeXml(undefined, edit)]) {
        expect(xmlProblem(xml)).toBeNull()
        expect(count(xml, 'a:fmtScheme')).toBe(1)
        expect(xml).toContain(format)
        expect(xml).toMatch(new RegExp(`<a:theme [^>]*name="${theme.name}"`))
        expect(readThemeColors(xml)).toEqual(theme.colors)
        const read = readThemeFonts(xml)!
        expect([read.major, read.minor]).toEqual([theme.fonts.major, theme.fonts.minor])
        expect(read.majorScripts ?? {}).toEqual(theme.fonts.majorScripts)
        expect(read.minorScripts ?? {}).toEqual(theme.fonts.minorScripts)
        expect(xml).toContain(`<a:fontScheme name="${theme.name}">`)
      }
    })
  }

  it('a Colors or Fonts pick keeps the format scheme the file came with', () => {
    const own = /<a:fmtScheme[\s\S]*<\/a:fmtScheme>/.exec(OFFICE_THEME_XML)![0]
    const facet = OFFICE_THEMES.find((t) => t.id === 'facet')!
    expect(editedThemeXml(OFFICE_THEME_XML, { colors: facet.colors })).toContain(own)
    expect(editedThemeXml(OFFICE_THEME_XML, { fonts: facet.fonts })).toContain(own)
  })

  it('a colour pick keeps the system pair for black text on a white page and states any other', () => {
    const grey = editedThemeXml(OFFICE_THEME_XML, { colors: { dk1: '333333', lt1: 'FFFFFF' } })
    expect(grey).toContain('<a:dk1><a:srgbClr val="333333"/></a:dk1>')
    expect(grey).toContain('<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>')
    const back = editedThemeXml(grey, { colors: { dk1: '000000' } })
    expect(back).toContain('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>')
  })
})

describe('a theme pick shows in tables, shapes and charts before it is saved', () => {
  const facet = OFFICE_THEMES.find((t) => t.id === 'facet')!
  const edit: ThemeEdit = {
    fonts: facet.fonts,
    colors: facet.colors,
    format: { name: facet.name, xml: facet.format },
  }
  const untouched = (parsed: ParsedDocFull) =>
    parsed.blocks.flatMap((b) =>
      b.hidden || b.docxIndex === null
        ? []
        : [{ kind: 'original' as const, docxIndex: b.docxIndex }],
    )
  const table = (blocks: Block[]) => blocks.find((b) => b.type === 'table')!
  const chart = (blocks: Block[]) => blocks.find((b) => b.chartDisplay)!.chartDisplay!
  const box = (blocks: Block[]) => blocks.find((b) => b.textboxes?.length)!.textboxes![0]!

  for (const themed of [true, false]) {
    it(`the live parse equals save and reopen, ${themed ? 'with' : 'without'} a theme part`, async () => {
      const bytes = await themedDocx(themed)
      const parsed = await parseDocx(bytes)
      const live = await parseDocx(await themedDocxBytes(bytes, edit))
      const saved = await parseDocx(
        await saveDocx(parsed, untouched(parsed), {
          themeFonts: facet.fonts,
          themeColors: facet.colors,
          themeFormat: edit.format!,
        }),
      )
      expect(live.blocks).toEqual(saved.blocks)
      expect(live.styles).toEqual(saved.styles)
      expect(live.themeColors).toEqual(saved.themeColors)
      expect(live.themeFonts).toEqual(saved.themeFonts)

      const colors = (blocks: Block[]) => {
        const display = chart(blocks)
        return [display.series[0]!.color, display.palette?.[1]]
      }
      expect(colors(parsed.blocks)).toEqual(['4472C4', 'ED7D31'])
      expect(colors(live.blocks)).toEqual(['90C226', '54A021'])
      expect(box(parsed.blocks)).toMatchObject({ fill: 'ED7D31', borderColor: '4472C4' })
      expect(box(live.blocks)).toMatchObject({ fill: '54A021', borderColor: '90C226' })
      const look = (doc: ParsedDocFull) => {
        const style = doc.styles.get('Banded')!.tableDisplay!
        const own = table(doc.blocks).table!.rows[2]![1]!.fill
        return [style.firstRow?.fill, style.band1Fill, style.borders?.insideH?.color, own]
      }
      // on the file's own theme Word's cached values stand as they are
      expect(look(parsed)).toEqual(['4472C4', 'D9E2F3', '8EAADB', '70AD47'])
      const [head, band, rule, own] = look(live)
      expect([head, own]).toEqual(['90C226', '918655'])
      expect(band).not.toBe('D9E2F3')
      expect(rule).not.toBe('8EAADB')
    })
  }

  it('the saved theme part carries the format scheme and the document keeps its references', async () => {
    const parsed = await parseDocx(await themedDocx())
    const zip = await JSZip.loadAsync(
      await saveDocx(parsed, untouched(parsed), {
        themeFonts: facet.fonts,
        themeColors: facet.colors,
        themeFormat: edit.format!,
      }),
    )
    const theme = await zip.file('word/theme/theme1.xml')!.async('string')
    expect(xmlProblem(theme)).toBeNull()
    expect(theme).toContain(facet.format)
    expect(theme).toContain('<a:hlink><a:srgbClr val="99CA3C"/></a:hlink>')
    expect(theme).toContain('<a:font script="Hans"')
    const body = await zip.file('word/document.xml')!.async('string')
    expect(body).toContain('<a:schemeClr val="accent2"/>')
    expect(body).toContain('w:themeColor="accent2"')
  })

  it('writes a script face with a dollar sign in its name as it is', () => {
    const theme = editedThemeXml(OFFICE_THEME_XML, {
      fonts: { major: 'Aptos Display', minor: 'Aptos', majorScripts: { Hans: "A$&B$$C$'D" } },
    })
    expect(xmlProblem(theme)).toBeNull()
    expect(theme).toContain(`<a:font script="Hans" typeface="A$&amp;B$$C$'D"/></a:majorFont>`)
  })
})
