import { describe, expect, it } from 'vitest'
import { normalizePastedMath } from '../src/renderer/editor/math'

describe('normalizePastedMath', () => {
  it('converts inline backslash-paren delimiters to dollars', () => {
    expect(normalizePastedMath(String.raw`\(E=mc^2\)`)).toBe('$E=mc^2$')
  })

  it('converts block backslash-bracket delimiters to double dollars', () => {
    expect(normalizePastedMath(String.raw`\[x^2\]`)).toBe('$$x^2$$')
  })

  it('leaves currency amounts untouched', () => {
    expect(normalizePastedMath('paid $5 and $10')).toBe('paid $5 and $10')
  })

  it('ignores empty and multiline inline delimiters', () => {
    expect(normalizePastedMath(String.raw`\(\)`)).toBe(String.raw`\(\)`)
    expect(normalizePastedMath('plain text')).toBe('plain text')
  })

  it('leaves Markdown-escaped brackets (citations) untouched', () => {
    expect(normalizePastedMath(String.raw`see \[1\] and \[2\]`)).toBe(
      String.raw`see \[1\] and \[2\]`,
    )
  })

  it('leaves LaTeX source verbatim when pasting into a code block', () => {
    const source = String.raw`\(x_1\) and \[\int_0^1 x\]`
    expect(normalizePastedMath(source, true)).toBe(source)
    expect(normalizePastedMath(source, false)).not.toBe(source)
  })

  it('does not convert inline delimiters nested inside a display block', () => {
    expect(normalizePastedMath(String.raw`\[a = \(b^2\)\]`)).toBe(String.raw`$$a = \(b^2\)$$`)
    expect(normalizePastedMath(String.raw`$$x^2$$ and \(y_1\)`)).toBe('$$x^2$$ and $y_1$')
  })

  it('leaves Markdown-escaped parens untouched', () => {
    expect(normalizePastedMath(String.raw`\(note\)`)).toBe(String.raw`\(note\)`)
  })

  it('leaves inline-looking block math mid-line untouched', () => {
    expect(normalizePastedMath(String.raw`see \[x^2\] here`)).toBe(String.raw`see \[x^2\] here`)
  })

  it('converts display math alone on its line', () => {
    expect(
      normalizePastedMath(String.raw`intro
\[x^2 + y^2\]
outro`),
    ).toBe('intro\n$$x^2 + y^2$$\noutro')
  })
})
