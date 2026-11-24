import { describe, expect, it } from 'vitest'
import { BUILT_IN_TABLE_STYLES, builtInTableStyle, builtInTableStyleInfos } from '../src/index'

describe("Word's built-in table styles", () => {
  it('lists the 105 gallery styles in Word order, grouped like its gallery', () => {
    expect(BUILT_IN_TABLE_STYLES).toHaveLength(105)
    expect(BUILT_IN_TABLE_STYLES.slice(0, 3).map((style) => style.name)).toEqual([
      'Table Grid',
      'Grid Table Light',
      'Plain Table 1',
    ])
    const groups = new Map<string, number>()
    for (const style of BUILT_IN_TABLE_STYLES)
      groups.set(style.group, (groups.get(style.group) ?? 0) + 1)
    expect(Object.fromEntries(groups)).toEqual({ plain: 7, grid: 49, list: 49 })
    for (const style of BUILT_IN_TABLE_STYLES) {
      expect(style.xml).toContain(`w:styleId="${style.styleId}"`)
      expect(style.xml).not.toContain('w:rsid')
    }
  })

  it('reads them for previews: header fill, banding and borders', async () => {
    const infos = await builtInTableStyleInfos()
    expect(infos.size).toBe(105)
    const grid4 = infos.get('GridTable4-Accent1')!
    expect(grid4).toMatchObject({ name: 'Grid Table 4 Accent 1', type: 'table' })
    // themed colors resolve through the theme (the default one here), not the saved hex
    expect(grid4.tableDisplay?.firstRow).toMatchObject({ fill: '156082', bold: true })
    expect(grid4.tableDisplay?.band1Fill).toMatch(/^[0-9A-F]{6}$/)
    expect(infos.get('TableGrid')?.tableDisplay?.borders?.insideH?.style).toBe('single')
    expect(builtInTableStyle('ListTable3')?.name).toBe('List Table 3')
  })
})
