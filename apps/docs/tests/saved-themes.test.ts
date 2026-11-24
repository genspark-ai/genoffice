import { describe, expect, it } from 'vitest'
import { OFFICE_THEMES, themeXmlOf } from '@genoffice/docx-engine'
import { savedThemesOf, themePickOf } from '../src/renderer/saved-themes'

const facet = OFFICE_THEMES.find((t) => t.id === 'facet')!
const xml = themeXmlOf({
  fonts: facet.fonts,
  colors: facet.colors,
  format: { name: facet.name, xml: facet.format },
})

describe('the theme files of the themes folder', () => {
  it('lists a file under its name without the extension, with the theme it holds', () => {
    const [saved, ...rest] = savedThemesOf([
      { name: 'Quarterly report.thmx', xml },
      { name: 'Broken.thmx', xml: '<a:themeManager/>' },
    ])
    expect(rest).toEqual([])
    expect(saved).toMatchObject({ file: 'Quarterly report.thmx', name: 'Quarterly report' })
    const pick = themePickOf(saved!.theme)
    expect(pick.colors).toEqual(facet.colors)
    expect(pick.fonts).toMatchObject({ major: 'Trebuchet MS', minor: 'Trebuchet MS' })
    expect(pick.format).toEqual({ name: 'Facet', xml: facet.format })
    expect(Object.keys(pick).sort()).toEqual(['colors', 'fonts', 'format'])
  })
})
