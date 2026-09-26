import { describe, expect, it } from 'vitest'
import { wordSegmentAt, wordUnitAt } from '../src/renderer/editor/word-range'

describe('wordSegmentAt', () => {
  it('finds the word under the offset and the one a caret just left', () => {
    expect(wordSegmentAt('alpha beta gamma', 7)).toEqual({ start: 6, end: 10 })
    expect(wordSegmentAt('alpha beta gamma', 10)).toEqual({ start: 6, end: 10 })
    expect(wordSegmentAt('alpha beta gamma', 0)).toEqual({ start: 0, end: 5 })
  })

  it('returns null on whitespace with no word right before it', () => {
    expect(wordSegmentAt('  x', 0)).toBeNull()
    expect(wordSegmentAt('', 0)).toBeNull()
  })

  it('treats leaf placeholders as boundaries', () => {
    expect(wordSegmentAt('one\0two', 5)).toEqual({ start: 4, end: 7 })
  })
})

describe('wordUnitAt', () => {
  it('carries the trailing spaces like a Word word selection', () => {
    expect(wordUnitAt('alpha  beta gamma', 2)).toEqual({ start: 0, end: 7 })
    expect(wordUnitAt('alpha beta gamma', 12)).toEqual({ start: 11, end: 16 })
  })

  it('does not swallow other whitespace or punctuation', () => {
    expect(wordUnitAt('alpha\tbeta', 1)).toEqual({ start: 0, end: 5 })
    expect(wordUnitAt('alpha, beta', 1)).toEqual({ start: 0, end: 5 })
  })

  it('segments CJK text into dictionary words', () => {
    const text = '\u6211\u4eec\u7231\u5317\u4eac'
    const seg = wordUnitAt(text, 3)
    expect(seg).not.toBeNull()
    expect(seg!.end - seg!.start).toBeLessThan(text.length)
  })
})
