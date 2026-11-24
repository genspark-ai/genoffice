import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  OFFICE_THEMES,
  buildThmx,
  readDocxThemeXml,
  readThemeFileXml,
  themeOfXml,
  themeXmlOf,
  themedDocxBytes,
} from '../src/index'
import { OFFICE_THEME_XML, themedDocx } from './helpers/themed-docx'
import { xmlProblem } from './helpers/strict-xml'

const facet = OFFICE_THEMES.find((t) => t.id === 'facet')!
const facetPick = {
  fonts: facet.fonts,
  colors: facet.colors,
  format: { name: facet.name, xml: facet.format },
}

describe('a theme as a file', () => {
  it('writes a theme file whose main part is a theme manager pointing at the theme', async () => {
    const zip = await JSZip.loadAsync(await buildThmx(themeXmlOf(facetPick)))
    expect(Object.keys(zip.files).sort()).toEqual([
      '[Content_Types].xml',
      '_rels/',
      '_rels/.rels',
      'theme/',
      'theme/theme/',
      'theme/theme/_rels/',
      'theme/theme/_rels/themeManager.xml.rels',
      'theme/theme/theme1.xml',
      'theme/theme/themeManager.xml',
    ])
    const types = await zip.file('[Content_Types].xml')!.async('string')
    expect(types).toContain(
      '<Override PartName="/theme/theme/themeManager.xml" ContentType="application/vnd.openxmlformats-officedocument.themeManager+xml"/>',
    )
    expect(types).toContain(
      '<Override PartName="/theme/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    )
    expect(await zip.file('_rels/.rels')!.async('string')).toMatch(
      /relationships\/officeDocument" Target="theme\/theme\/themeManager\.xml"/,
    )
    for (const part of Object.keys(zip.files).filter((name) => /\.(xml|rels)$/.test(name)))
      expect(xmlProblem(await zip.file(part)!.async('string')), part).toBeNull()
  })

  it('reads back everything a theme pick carries', async () => {
    const xml = await readThemeFileXml(await buildThmx(themeXmlOf(facetPick)))
    const theme = themeOfXml(xml!)!
    expect(theme.name).toBe('Facet')
    expect(theme.colors).toEqual(facet.colors)
    expect(theme.fonts).toMatchObject(facet.fonts)
    expect(theme.format).toEqual(facetPick.format)
    // a second trip changes nothing
    expect(themeXmlOf(theme)).toBe(xml)
  })

  it('names the theme after the file it is saved as', async () => {
    const xml = await readThemeFileXml(await buildThmx(themeXmlOf(facetPick), 'Mine & yours'))
    expect(xml).toContain(
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Mine &amp; yours">',
    )
    const theme = themeOfXml(xml!)!
    expect(theme.format.xml).toBe(facet.format)
    expect(theme.colors.name).toBe('Facet')
  })

  it('finds the theme part of a document and of a template through the main part', async () => {
    const docx = await themedDocx()
    expect(await readThemeFileXml(docx)).toBe(OFFICE_THEME_XML)
    const zip = await JSZip.loadAsync(docx)
    const types = await zip.file('[Content_Types].xml')!.async('string')
    zip.file('[Content_Types].xml', types.replace('document.main+xml', 'template.main+xml'))
    const dotx = await zip.generateAsync({ type: 'uint8array' })
    expect(themeOfXml((await readThemeFileXml(dotx))!)?.colors.accent1).toBe('4472C4')
  })

  it('reads the document theme with a pending pick written in, and makes one for a document without', async () => {
    const docx = await themedDocx()
    expect(await readDocxThemeXml(docx)).toBe(OFFICE_THEME_XML)
    const picked = themeOfXml(await readDocxThemeXml(docx, facetPick))!
    expect(picked.colors).toEqual(facet.colors)
    expect(picked.name).toBe('Facet')
    const saved = await themedDocxBytes(docx, facetPick)
    expect(await readThemeFileXml(saved)).toBe(await readDocxThemeXml(docx, facetPick))
    const bare = themeOfXml(await readDocxThemeXml(await themedDocx(false)))!
    expect(bare.fonts.minor).toBe('Calibri')
    expect(bare.colors.accent1).toBe('4472C4')
  })

  it('gives null for a package without a theme part and for a part that is no theme', async () => {
    expect(await readThemeFileXml(await themedDocx(false))).toBeNull()
    expect(themeOfXml('<w:document/>')).toBeNull()
    await expect(readThemeFileXml(new TextEncoder().encode('not a zip'))).rejects.toThrow()
  })

  it('reads a theme of a foreign package: absolute targets, and picture fills become plain fills', async () => {
    const withPicture = OFFICE_THEME_XML.replace(
      /<a:fillStyleLst>/,
      '<a:fillStyleLst><a:blipFill rotWithShape="1"><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch/></a:blipFill>',
    )
    const zip = await JSZip.loadAsync(await buildThmx(withPicture))
    const rels = await zip.file('theme/theme/_rels/themeManager.xml.rels')!.async('string')
    zip.file(
      'theme/theme/_rels/themeManager.xml.rels',
      rels.replace('Target="theme1.xml"', 'Target="/theme/theme/theme1.xml"'),
    )
    const xml = await readThemeFileXml(await zip.generateAsync({ type: 'uint8array' }))
    expect(xml).toContain('a:blipFill')
    const theme = themeOfXml(xml!)!
    expect(theme.format.xml).not.toContain('blip')
    expect(theme.format.xml).not.toContain('r:embed')
    expect(xmlProblem(themeXmlOf(theme))).toBeNull()
  })

  it('keeps a name with a dollar sign as the name', () => {
    const name = "Save $100 $& $1 $'"
    const xml = themeXmlOf({
      fonts: { ...facet.fonts, name },
      colors: { ...facet.colors, name },
      format: { name, xml: facet.format },
    })
    expect(xmlProblem(xml)).toBeNull()
    expect(xml.split(`name="Save $100 $&amp; $1 $'"`)).toHaveLength(4)
  })
})
