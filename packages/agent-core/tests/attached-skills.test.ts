import { describe, expect, it } from 'vitest'
import { withAttachedSkills } from '../src/attached-skills'

describe('withAttachedSkills', () => {
  it('leaves the instruction alone when nothing is attached', () => {
    expect(withAttachedSkills('Make a table', [])).toBe('Make a table')
  })

  it('puts every skill body in front of the request, each in its own tagged block', () => {
    const out = withAttachedSkills('Make a table', [
      { name: 'tables', body: 'Always add a header row.\n' },
      { name: 'tone', body: 'Write plainly.' },
    ])
    expect(out.indexOf('<skill name="tables">')).toBeLessThan(out.indexOf('<skill name="tone">'))
    expect(out).toContain('<skill name="tables">\nAlways add a header row.\n</skill>')
    expect(out).toContain('<skill name="tone">\nWrite plainly.\n</skill>')
    expect(out.endsWith('User request:\nMake a table')).toBe(true)
    expect(out).toContain('the following skills to this request')
  })

  it('cannot have its tag broken by a quote in the name', () => {
    const out = withAttachedSkills('x', [{ name: 'a"b', body: 'c' }])
    expect(out).toContain('<skill name="ab">')
  })
})
