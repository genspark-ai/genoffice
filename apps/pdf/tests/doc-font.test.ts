import { describe, expect, it } from 'vitest'
import { mapDocFont } from '../src/renderer/doc-font'

describe('mapDocFont', () => {
  it('strips the subset tag and leads with the real family', () => {
    const f = mapDocFont('ABCDEF+NotoSerifCJKsc-Bold')
    expect(f.css.startsWith('"Noto Serif CJKsc"')).toBe(true)
    expect(f.weight).toBe(700)
  })

  it('classifies CJK serif names onto local serif faces', () => {
    for (const name of ['NotoSerifCJKsc-Regular', 'SourceHanSerifCN-Heavy', 'SimSun', 'STSong']) {
      expect(mapDocFont(name).css).toContain('Songti SC')
    }
  })

  it('classifies CJK sans names onto local sans faces', () => {
    for (const name of ['NotoSansCJKsc-Medium', 'MicrosoftYaHei', 'SimHei', 'PingFangSC-Regular']) {
      expect(mapDocFont(name).css).toContain('PingFang SC')
    }
  })

  it('maps kai and fangsong styles', () => {
    expect(mapDocFont('KaiTi').css).toContain('Kaiti SC')
    expect(mapDocFont('STFangsong').css).toContain('STFangsong')
  })

  it('classifies latin serif and sans names', () => {
    expect(mapDocFont('TimesNewRomanPSMT').css).toContain('Times New Roman')
    expect(mapDocFont('Georgia-Italic').css).toContain('Georgia')
    expect(mapDocFont('Helvetica').css).toContain('Helvetica')
    expect(mapDocFont('Arial-BoldMT').weight).toBe(700)
  })

  it('reads weight and italic from name suffixes', () => {
    expect(mapDocFont('SourceHanSansSC-Heavy').weight).toBe(900)
    expect(mapDocFont('NotoSans-SemiBold').weight).toBe(600)
    expect(mapDocFont('Lato-Light').weight).toBe(300)
    expect(mapDocFont('Georgia-BoldItalic')).toMatchObject({ weight: 700, italic: true })
    expect(mapDocFont('Helvetica-Oblique').italic).toBe(true)
    expect(mapDocFont('Roboto-Regular').weight).toBeUndefined()
  })

  it('style tokens do not leak into the family candidates', () => {
    const f = mapDocFont('ABCDEF+SourceHanSerifCN-Bold')
    expect(f.css).toContain('"Source Han Serif CN"')
    expect(f.css).not.toMatch(/"[^"]*Bold[^"]*"/)
  })

  it('unknown names still yield a usable stack', () => {
    const f = mapDocFont('F1')
    expect(f.css).toContain('sans-serif')
  })
})
