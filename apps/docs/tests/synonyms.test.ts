import { describe, expect, it } from 'vitest'
import { parseSynonyms } from '../src/renderer/ai/synonyms'

describe('parseSynonyms', () => {
  it('reads one candidate per line, without list markers or quotes', () => {
    expect(parseSynonyms('1. quick\n- rapid\n* "swift"\n2) speedy.', 'fast')).toEqual([
      'quick',
      'rapid',
      'swift',
      'speedy',
    ])
  })

  it('reads a comma-separated reply, CJK separators included', () => {
    expect(parseSynonyms('quick, rapid; swift', 'fast')).toEqual(['quick', 'rapid', 'swift'])
    expect(
      parseSynonyms('\u8fc5\u901f\u3001\u5feb\u6377\uff0c\u98de\u5feb\u3002', '\u5feb'),
    ).toEqual(['\u8fc5\u901f', '\u5feb\u6377', '\u98de\u5feb'])
  })

  it('drops repeats, the word itself, sentences, and anything past the eighth', () => {
    expect(parseSynonyms('Fast\nquick\nQuick\n' + 'x'.repeat(41), 'fast')).toEqual(['quick'])
    const many = Array.from({ length: 12 }, (_, i) => `w${i}`).join('\n')
    expect(parseSynonyms(many, 'fast')).toHaveLength(8)
    expect(parseSynonyms('', 'fast')).toEqual([])
  })
})
