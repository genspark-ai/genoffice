import { describe, expect, it } from 'vitest'
import {
  checkPlaceholders,
  collectPlaceholders,
  isWholePlaceholder,
  placeholderInstruction,
  placeholderSource,
  readPlaceholderLabel,
  sanitizeLabel,
  MAX_LABEL_LENGTH,
  type PlaceholderSpec,
} from '../src/redact-core'

/**
 * The rules every editor now shares, pinned once.
 *
 * Three of the five originals spelled `checkPlaceholders` differently and
 * meant the same thing, which is the failure this file exists to prevent: a
 * near-identical copy is where a real difference goes to hide. These are the
 * properties the copies all agreed on, plus the one place they genuinely
 * differ — a comment carrier needs two more characters stripped.
 */

const SECRET = '13800138000'

describe('a label', () => {
  it('loses the characters that would make a marker ambiguous', () => {
    expect(sanitizeLabel('cli{ent} "phone"')).toBe('client phone')
    expect(sanitizeLabel('a  b   c')).toBe('a b c')
  })

  it('is capped, and the cap does not leave a dangling space', () => {
    const long = 'x'.repeat(MAX_LABEL_LENGTH + 10)
    expect(sanitizeLabel(long)).toHaveLength(MAX_LABEL_LENGTH)
    expect(sanitizeLabel(`${'x'.repeat(MAX_LABEL_LENGTH - 1)} tail`)).not.toMatch(/\s$/)
  })

  it('loses two more characters for a carrier that is an XML comment', () => {
    // the sheet's label ends up in a comment, where `*` and `/` would close it
    // early. A document's does not, and stripping them there would rename a
    // span the reader can see.
    expect(sanitizeLabel('a*b/c', { forComment: true })).toBe('abc')
    expect(sanitizeLabel('a*b/c')).toBe('a*b/c')
  })

  it('becomes the marker the model reads', () => {
    expect(placeholderSource('client phone')).toBe('{{client phone}}')
    // an empty label still produces a marker rather than empty braces
    expect(placeholderSource('   ')).toBe('{{private}}')
  })
})

describe('reading a marker back', () => {
  it('accepts a whole marker and refuses anything else', () => {
    expect(isWholePlaceholder(' {{a}} ')).toBe(true)
    expect(isWholePlaceholder('call {{a}}')).toBe(false)
    expect(isWholePlaceholder('{{}}')).toBe(false)
  })

  it('gives the label, or null', () => {
    expect(readPlaceholderLabel('{{client phone}}')).toBe('client phone')
    expect(readPlaceholderLabel('call {{a}}')).toBeNull()
  })

  it('finds every whole marker and no half of one', () => {
    expect(collectPlaceholders('a {{one}} b {{two}}')).toEqual(['{{one}}', '{{two}}'])
    expect(collectPlaceholders('{{a} and {b}}')).toEqual([])
  })
})

describe('a model answer against the markers it was given', () => {
  it('passes when every marker came through', () => {
    expect(checkPlaceholders('call {{a}} now', 'call {{a}} later')).toEqual([])
  })

  it('catches a marker that was dropped', () => {
    expect(checkPlaceholders('call {{a}}', 'call them')).toEqual([
      { found: '', expected: '{{a}}', reason: 'missing' },
    ])
  })

  it('catches one duplicated', () => {
    const issues = checkPlaceholders('{{a}}', '{{a}} {{a}}')
    expect(issues).toEqual([{ found: '{{a}}', expected: '{{a}}', reason: 'unknown' }])
  })

  it('catches a marker the model invented', () => {
    expect(checkPlaceholders('nothing', 'call {{a}}')).toEqual([
      { found: '{{a}}', expected: '', reason: 'unknown' },
    ])
  })

  it('catches one it split in half', () => {
    expect(checkPlaceholders('{{a}}', '{{a}').map((i) => i.reason)).toContain('split')
  })

  it('does not mistake a well-formed marker in the answer for a stray brace', () => {
    expect(checkPlaceholders('{{a}}', 'x {{a}} y')).toEqual([])
  })
})

describe('the prompt an editor sends', () => {
  const spec: PlaceholderSpec = {
    subject: 'document',
    opening:
      'Each stands in for something the reader has deliberately withheld from you; you cannot see what is inside, and that is the point.',
    boundary: 'a line break',
    middle: ['Write around them.'],
  }

  it('carries the shared rules whatever the editor supplies', () => {
    const out = placeholderInstruction(['a', 'b', 'a'], spec)
    for (const rule of [
      'Treat every placeholder as one indivisible object:',
      'same letters, same order, same spacing',
      'Never merge two into one',
      'Never rename, translate, re-case, expand or shorten it.',
      'Never drop one, and never add a placeholder',
    ]) {
      expect(out, rule).toContain(rule)
    }
  })

  it("uses the editor's own subject, boundary and middle", () => {
    const sheet = placeholderInstruction(['a'], {
      ...spec,
      subject: 'workbook',
      boundary: 'a cell boundary',
      middle: ['Cells are left out of every statistic.'],
    })
    expect(sheet).toContain('This workbook contains')
    expect(sheet).toContain('across a cell boundary')
    expect(sheet).toContain('The placeholders in this workbook:')
    expect(sheet).toContain('Cells are left out of every statistic.')
    expect(sheet).not.toContain('document')
  })

  it('lists each label once, through the sanitiser', () => {
    const out = placeholderInstruction(['a', 'b', 'a'], spec)
    expect(out).toContain('- {{a}}')
    expect(out).toContain('- {{b}}')
    expect(out.match(/- \{\{a\}\}/g)).toHaveLength(1)
  })

  it('carries the labels the reader chose, and nothing else', () => {
    // The span's *contents* never reach the prompt because the caller passes
    // labels, not text. A label the reader typed is theirs: a reader who calls
    // a span after the number in it has made a choice, and second-guessing it
    // here would only hide a mistake the reader can still see on screen.
    const out = placeholderInstruction(['client phone'], spec)
    expect(out).toContain('{{client phone}}')
    expect(out).not.toContain(SECRET)
  })
})
