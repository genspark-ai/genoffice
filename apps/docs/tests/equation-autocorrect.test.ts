import { beforeEach, describe, expect, it } from 'vitest'
import {
  DEFAULT_MATH_ENTRIES,
  resetAutocorrectPrefs,
  setAutocorrectList,
  setAutocorrectPref,
  upsertEntry,
} from '../src/renderer/autocorrect-pref'
import { equationAutoCorrect } from '../src/renderer/components/equation-autocorrect'

const typeInto = (text: string) => {
  let value = ''
  for (const ch of text) value = equationAutoCorrect(value, value + ch)
  return value
}

beforeEach(() => resetAutocorrectPrefs())

describe('equation box: recognized functions', () => {
  it('a bare function name becomes its upright command at the space or bracket', () => {
    expect(typeInto('sin x + max(a)')).toBe('\\sin x + \\max(a)')
    expect(typeInto('sech x')).toBe('\\operatorname{sech} x')
  })
  it('leaves commands, longer words and text groups alone', () => {
    expect(typeInto('\\sin x')).toBe('\\sin x')
    expect(typeInto('asin x')).toBe('asin x')
    expect(typeInto('\\text{the max size} ')).toBe('\\text{the max size} ')
  })
  it('follows the list edited in the dialog', () => {
    setAutocorrectList('functions', ['erf'])
    expect(typeInto('erf(x) sin x')).toBe('\\operatorname{erf}(x) sin x')
  })
})

describe('equation box: Math AutoCorrect entries', () => {
  it('replaces symbols and names LaTeX does not read, at the space', () => {
    setAutocorrectList('math', upsertEntry(DEFAULT_MATH_ENTRIES, '\\RR', '\u211d'))
    expect(typeInto('a <= b -> \\RR ')).toBe('a \u2264 b \u2192 \u211d ')
  })
  it('leaves names LaTeX reads as typed', () => {
    expect(typeInto('\\alpha \\sqrt x ')).toBe('\\alpha \\sqrt x ')
  })
  it('is off when its switch or the master switch is off', () => {
    setAutocorrectPref('mathInside', false)
    expect(typeInto('a <= b sin x')).toBe('a <= b \\sin x')
    setAutocorrectPref('enabled', false)
    expect(typeInto('a <= b sin x')).toBe('a <= b sin x')
  })
  it('does not touch a pasted or edited value', () => {
    expect(equationAutoCorrect('a', 'a <= b ')).toBe('a <= b ')
    expect(equationAutoCorrect('sin x', 'sin ')).toBe('sin ')
  })
})
