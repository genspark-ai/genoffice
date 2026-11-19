import { describe, expect, it } from 'vitest'
import { flagBool, parseArgs } from '../src/args'

describe('boolean flags reject =values', () => {
  it('--force=false used to be read as --force (flagBool ignored the value)', () => {
    expect(() => parseArgs(['convert', '--force=false', 'a.docx'], new Set(['force']))).toThrow(
      /boolean flag/,
    )
  })

  it('the usage error names only spellings the parser actually supports', () => {
    expect(() => parseArgs(['--force=false'], new Set(['force']))).toThrow(
      '--force is a boolean flag; use --force, not --force=<value>',
    )
  })

  it('plain boolean flags and string =values keep working', () => {
    const a = parseArgs(['convert', '--force', 'a.docx'], new Set(['force']))
    expect(flagBool(a, 'force')).toBe(true)
    const b = parseArgs(['convert', '--out=x.docx', 'a.docx'], new Set(['force']))
    expect(b.flags.out).toBe('x.docx')
  })
})
