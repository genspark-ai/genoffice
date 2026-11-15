import { describe, expect, it } from 'vitest'
import { clarifyOptionText } from '../src/renderer/ai/tools'

describe('clarifyOptionText', () => {
  it('keeps string options and reads object options instead of "[object Object]"', () => {
    expect(clarifyOptionText('Formal')).toBe('Formal')
    expect(clarifyOptionText({ label: 'Formal', value: 'formal' })).toBe('Formal')
    expect(clarifyOptionText({ title: 'Dark' })).toBe('Dark')
    expect(clarifyOptionText({ id: 1 })).toBe('')
    expect(clarifyOptionText(undefined)).toBe('')
  })
})
