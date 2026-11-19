/**
 * Double-click selects Word's unit: the word plus its trailing spaces, so a
 * double-click inside a fill-in blank of spaces never selects the spaces alone.
 */
import { describe, expect, it } from 'vitest'
import { wordUnitAt } from '../src/renderer/editor/dblclick-word'

const unit = (text: string, index: number) => {
  const r = wordUnitAt(text, index)
  return r && text.slice(r[0], r[1])
}

describe('wordUnitAt', () => {
  const blank =
    '\u7532\u65b9\uff1a                           \u4ee5\u4e0b\u7b80\u79f0\uff08\u201c\u7532\u65b9\u201d\uff09'

  it('takes the preceding punctuation with the whole space run', () => {
    expect(unit(blank, 3)).toBe('\uff1a' + ' '.repeat(27))
    expect(unit(blank, 20)).toBe('\uff1a' + ' '.repeat(27))
    expect(unit(blank, 29)).toBe('\uff1a' + ' '.repeat(27))
  })

  it('attaches trailing spaces to a Latin word and stops before the next word', () => {
    expect(unit('name: dddd            \u4ee5\u4e0b', 8)).toBe('dddd            ')
    expect(unit('name: dddd            \u4ee5\u4e0b', 15)).toBe('dddd            ')
    expect(unit('name: dddd            \u4ee5\u4e0b', 0)).toBe('name')
  })

  it('leaves leading spaces as their own unit and skips inline leaves', () => {
    expect(unit('   abc', 1)).toBe('   ')
    expect(unit('a\ufffc  b', 1)).toBeNull()
    expect(unit('a\ufffc  b', 2)).toBe('  ')
  })
})
