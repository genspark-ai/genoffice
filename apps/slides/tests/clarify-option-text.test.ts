import { describe, expect, it } from 'vitest'
import { clarifyOptionText } from '../src/renderer/ai/slides-skill'

describe('clarifyOptionText', () => {
  it('keeps string options and reads object options instead of "[object Object]"', () => {
    expect(clarifyOptionText('Formal')).toBe('Formal')
    expect(clarifyOptionText({ label: 'Formal', value: 'formal' })).toBe('Formal')
    expect(clarifyOptionText({ text: 'Casual' })).toBe('Casual')
    expect(clarifyOptionText({ value: 'blue' })).toBe('blue')
    expect(clarifyOptionText(5)).toBe('5')
    expect(clarifyOptionText({ id: 1 })).toBe('')
    expect(clarifyOptionText(null)).toBe('')
    expect(clarifyOptionText({ key: 'a', desc: 'Formal' })).toBe('Formal')
    expect(clarifyOptionText('  ')).toBe('')
    expect(clarifyOptionText(' Casual ')).toBe('Casual')
  })
})
