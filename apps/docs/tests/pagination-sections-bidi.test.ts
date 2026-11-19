import { describe, expect, it } from 'vitest'
import { sectionBidi } from '../src/renderer/pagination-sections'

type Section = Parameters<typeof sectionBidi>[0]
const sec = (sectPrXml: string): Section => ({ sectPrXml }) as Section

describe('sectionBidi', () => {
  it('detects self-closing and explicit-on marks', () => {
    expect(sectionBidi(sec('<w:sectPr><w:bidi/></w:sectPr>'))).toBe(true)
    expect(sectionBidi(sec('<w:sectPr><w:bidi w:val="1"/></w:sectPr>'))).toBe(true)
    expect(sectionBidi(sec('<w:sectPr><w:bidi w:val="on"/></w:sectPr>'))).toBe(true)
  })

  it('detects paired tags written by non-Word producers', () => {
    expect(sectionBidi(sec('<w:sectPr><w:bidi></w:bidi></w:sectPr>'))).toBe(true)
  })

  it('treats missing and explicit-off marks as left-to-right', () => {
    expect(sectionBidi(sec('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'))).toBe(false)
    expect(sectionBidi(sec('<w:sectPr><w:bidi w:val="0"/></w:sectPr>'))).toBe(false)
    expect(sectionBidi(sec('<w:sectPr><w:bidi w:val="off"></w:bidi></w:sectPr>'))).toBe(false)
  })
})
