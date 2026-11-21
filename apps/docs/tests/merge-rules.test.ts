import { conditionHolds } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { ifInstruction, skipIfInstruction } from '../src/renderer/mail-merge/rules'

const ADA = { City: 'London', Orders: '12' }
const ALAN = { City: 'Wilmslow', Orders: '3' }

describe('Mailings > Rules', () => {
  it('writes the field Word writes, with the data column in guillemets', () => {
    expect(skipIfInstruction('City', '=', 'London')).toBe('SKIPIF «City» = "London"')
    expect(ifInstruction('Orders', '>=', '10', 'Thank you', 'Welcome')).toBe(
      'IF «Orders» >= "10" "Thank you" "Welcome"',
    )
  })

  it('a straight quote in a text cannot end the literal early', () => {
    expect(ifInstruction('City', '=', 'x', 'say "hi"', '')).toBe('IF «City» = "x" "say ”hi”" ""')
  })

  it('the merge engine reads the condition back per record', () => {
    const skip = skipIfInstruction('City', '=', 'London')
    expect(conditionHolds(skip, ADA)).toBe(true)
    expect(conditionHolds(skip, ALAN)).toBe(false)
    const many = ifInstruction('Orders', '>=', '10', 'a', 'b')
    expect(conditionHolds(many, ADA)).toBe(true)
    expect(conditionHolds(many, ALAN)).toBe(false)
  })
})
