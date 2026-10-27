import { describe, expect, it } from 'vitest'
import { anchorAt, discardAnchored, replaceRange, writeAnchored } from '../src/voice-text-splice'

describe('replaceRange', () => {
  it('replaces a half-open span', () => {
    expect(replaceRange('abcdef', 1, 3, 'X')).toBe('aXdef')
  })

  it('clamps a span that ran past the end of a shortened draft', () => {
    expect(replaceRange('ab', 1, 99, 'X')).toBe('aX')
  })

  it('treats a reversed span as empty at its start', () => {
    expect(replaceRange('abcdef', 3, 1, 'X')).toBe('abcXdef')
  })
})

describe('anchorAt', () => {
  it('opens an empty span at the caret', () => {
    expect(anchorAt('hello', 3)).toEqual({ start: 3, end: 3 })
  })

  it('covers a selection so the first chunk replaces it', () => {
    expect(anchorAt('hello', 1, 4)).toEqual({ start: 1, end: 4 })
  })
})

describe('writeAnchored', () => {
  it('inserts at an empty span and advances it', () => {
    const written = writeAnchored('say: ', anchorAt('say: ', 5), 'hello')
    expect(written.value).toBe('say: hello')
    expect(written.anchor).toEqual({ start: 5, end: 10 })
  })

  it('replaces the previous interim transcript instead of appending to it', () => {
    const first = writeAnchored('say: ', anchorAt('say: ', 5), 'hel')
    const second = writeAnchored(first.value, first.anchor, 'hello there')
    expect(second.value).toBe('say: hello there')
  })

  it('replaces a selection', () => {
    expect(writeAnchored('one two', anchorAt('one two', 4, 7), 'three').value).toBe('one three')
  })

  it('leaves the text after the span untouched', () => {
    expect(writeAnchored('a [x] b', anchorAt('a [x] b', 2, 5), 'y').value).toBe('a y b')
  })
})

describe('discardAnchored', () => {
  it('takes the interim transcript back out', () => {
    const written = writeAnchored('say: ', anchorAt('say: ', 5), 'hello')
    expect(discardAnchored(written.value, written.anchor)).toBe('say: ')
  })
})
