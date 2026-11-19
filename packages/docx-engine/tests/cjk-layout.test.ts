import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, sectionSettingsFromXml } from '../src/index'
import { buildDocx } from './helpers/build-docx'

describe('emphasis marks (w:em)', () => {
  it('parses dot/circle emphasis onto runs and skips none', async () => {
    const p =
      '<w:p><w:r><w:rPr><w:em w:val="dot"/></w:rPr><w:t>重点</w:t></w:r>' +
      '<w:r><w:rPr><w:em w:val="circle"/></w:rPr><w:t>圈点</w:t></w:r>' +
      '<w:r><w:rPr><w:em w:val="none"/></w:rPr><w:t>普通</w:t></w:r></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: p }))
    const runs = parsed.blocks[0].runs!
    expect(runs[0].em).toBe('dot')
    expect(runs[1].em).toBe('circle')
    expect(runs[2].em).toBeUndefined()
  })

  it('keeps the w:em bytes via rawRPr for untouched-style saves', async () => {
    const p = '<w:p><w:r><w:rPr><w:em w:val="dot"/></w:rPr><w:t>重点</w:t></w:r></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: p }))
    expect(parsed.blocks[0].runs![0].rawRPr).toContain('<w:em w:val="dot"/>')
  })
})

describe('section textDirection (w:textDirection)', () => {
  it('reads vertical tbRl from sectPr and omits horizontal lrTb', () => {
    expect(
      sectionSettingsFromXml('<w:sectPr><w:textDirection w:val="tbRl"/></w:sectPr>').textDirection,
    ).toBe('tbRl')
    expect(
      sectionSettingsFromXml('<w:sectPr><w:textDirection w:val="lrTb"/></w:sectPr>').textDirection,
    ).toBeUndefined()
    expect(sectionSettingsFromXml('<w:sectPr/>').textDirection).toBeUndefined()
  })

  it('readSections surfaces per-section columns and direction from a real doc', async () => {
    const body =
      '<w:p><w:r><w:t>one</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:cols w:num="2" w:space="425"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>two</w:t></w:r></w:p>'
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: body, sectPrExtra: '<w:textDirection w:val="tbRl"/>' }),
    )
    const sections = readSections(parsed)
    expect(sections).toHaveLength(2)
    expect(sections[0].settings.columns).toBe(2)
    expect(sections[0].startType).toBe('continuous')
    expect(sections[1].settings.textDirection).toBe('tbRl')
  })
})
