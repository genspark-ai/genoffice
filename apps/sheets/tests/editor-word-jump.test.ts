import { describe, expect, it } from 'vitest'
import { wordJumpOffset } from '../src/renderer/editor-word-jump'

// Univer dataStreams terminate in "\r\n"; "\r" is the paragraph token.
const stream = (text: string) => `${text}\r\n`

describe('wordJumpOffset', () => {
  it('jumps right to the start of each next word', () => {
    const s = stream('alpha beta gamma')
    expect(wordJumpOffset(s, 0, 1)).toBe(6) // |alpha -> |beta
    expect(wordJumpOffset(s, 6, 1)).toBe(11) // |beta -> |gamma
    expect(wordJumpOffset(s, 11, 1)).toBe(16) // |gamma -> text end
    expect(wordJumpOffset(s, 16, 1)).toBe(16) // clamped before the terminator
  })

  it('jumps right from inside a word and from whitespace', () => {
    const s = stream('alpha beta')
    expect(wordJumpOffset(s, 2, 1)).toBe(6) // al|pha -> |beta
    expect(wordJumpOffset(s, 5, 1)).toBe(6) // alpha| -> |beta
  })

  it('jumps left to the start of the current word, then the previous one', () => {
    const s = stream('alpha beta gamma')
    expect(wordJumpOffset(s, 13, -1)).toBe(11) // ga|mma -> |gamma
    expect(wordJumpOffset(s, 11, -1)).toBe(6) // |gamma -> |beta
    expect(wordJumpOffset(s, 6, -1)).toBe(0) // |beta -> |alpha
    expect(wordJumpOffset(s, 0, -1)).toBe(0)
  })

  it('treats punctuation as a separator, not a word', () => {
    const s = stream('one, two; three')
    expect(wordJumpOffset(s, 0, 1)).toBe(5) // |one, -> |two
    expect(wordJumpOffset(s, 5, 1)).toBe(10) // |two; -> |three
    expect(wordJumpOffset(s, 10, -1)).toBe(5)
  })

  it('handles accented words as single words', () => {
    const s = stream('déjà vu café')
    expect(wordJumpOffset(s, 0, 1)).toBe(5) // |déjà -> |vu
    expect(wordJumpOffset(s, 5, 1)).toBe(8) // |vu -> |café
    expect(wordJumpOffset(s, 8, -1)).toBe(5)
    expect(wordJumpOffset(s, 2, -1)).toBe(0) // dé|jà -> |déjà
  })

  it('segments CJK runs into words', () => {
    const s = stream('\u4eca\u5929\u5929\u6c14\u5f88\u597d') // CJK sentence, escaped; public-hygiene: fixture
    const first = wordJumpOffset(s, 0, 1)
    expect(first).toBeGreaterThan(0)
    expect(first).toBeLessThanOrEqual(6)
    expect(wordJumpOffset(s, first, -1)).toBe(0)
  })

  it('stops at a paragraph token and crosses it on the next press', () => {
    const s = 'one two\rthree four\r\n'
    expect(wordJumpOffset(s, 4, 1)).toBe(7) // |two -> end of paragraph (before \r)
    expect(wordJumpOffset(s, 7, 1)).toBe(8) // crosses the token to |three
    expect(wordJumpOffset(s, 8, -1)).toBe(4) // back across: |three -> |two
  })

  it('treats tabs as whitespace, not boundaries', () => {
    const s = stream('one\ttwo')
    expect(wordJumpOffset(s, 0, 1)).toBe(4) // |one -> |two across the tab
    expect(wordJumpOffset(s, 4, -1)).toBe(0)
  })

  it('never lands on the stream terminator pair', () => {
    const s = stream('word')
    expect(wordJumpOffset(s, 0, 1)).toBe(4)
    expect(wordJumpOffset(s, 4, 1)).toBe(4)
    expect(wordJumpOffset(s, 99, 1)).toBe(4)
  })

  it('handles an empty cell', () => {
    const s = stream('')
    expect(wordJumpOffset(s, 0, 1)).toBe(0)
    expect(wordJumpOffset(s, 0, -1)).toBe(0)
  })
})
