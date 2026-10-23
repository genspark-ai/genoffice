import { describe, expect, it } from 'vitest'
import { revertParaMarkRPrChange, stripParaMarkRPrChange } from '../src'

describe('paragraph-mark w:rPrChange helpers', () => {
  const AT = 'w:id="3" w:author="X" w:date="2026-01-01T00:00:00Z"'

  it('strip removes the record and keeps the live formatting', () => {
    const raw = `<w:pPr><w:rPr><w:caps/><w:rPrChange ${AT}><w:rPr><w:strike/></w:rPr></w:rPrChange></w:rPr></w:pPr>`
    expect(stripParaMarkRPrChange(raw)).toBe('<w:pPr><w:rPr><w:caps/></w:rPr></w:pPr>')
  })

  it('strip drops the mark rPr when the record was its only content', () => {
    const raw = `<w:pPr><w:jc w:val="center"/><w:rPr><w:rPrChange ${AT}><w:rPr/></w:rPrChange></w:rPr></w:pPr>`
    expect(stripParaMarkRPrChange(raw)).toBe('<w:pPr><w:jc w:val="center"/></w:pPr>')
  })

  it('strip never touches the old snapshot inside a pPrChange record', () => {
    const raw =
      `<w:pPr><w:rPr><w:caps/><w:rPrChange ${AT}><w:rPr/></w:rPrChange></w:rPr>` +
      `<w:pPrChange w:id="9" w:author="Y"><w:pPr><w:rPr><w:rPrChange w:id="10" w:author="Y"><w:rPr/></w:rPrChange></w:rPr></w:pPr></w:pPrChange></w:pPr>`
    const out = stripParaMarkRPrChange(raw)
    expect(out).toContain('<w:caps/>')
    expect(out).toContain('<w:pPrChange')
    expect(out.match(/<w:rPrChange/g)).toHaveLength(1)
  })

  it('revert restores the recorded old properties and drops the record', () => {
    const raw = `<w:pPr><w:rPr><w:caps/><w:rPrChange ${AT}><w:rPr><w:strike/></w:rPr></w:rPrChange></w:rPr></w:pPr>`
    expect(revertParaMarkRPrChange(raw)).toBe('<w:pPr><w:rPr><w:strike/></w:rPr></w:pPr>')
  })

  it('revert removes the mark rPr when the recorded old properties are empty', () => {
    const raw = `<w:pPr><w:rPr><w:caps/><w:rPrChange ${AT}><w:rPr/></w:rPrChange></w:rPr></w:pPr>`
    expect(revertParaMarkRPrChange(raw)).toBe('<w:pPr></w:pPr>')
  })

  it('revert returns null when the paragraph has no w:rPrChange record', () => {
    const raw = '<w:pPr><w:rPr><w:caps/></w:rPr></w:pPr>'
    expect(revertParaMarkRPrChange(raw)).toBeNull()
  })
})
