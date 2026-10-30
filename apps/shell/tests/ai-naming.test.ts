import { describe, expect, it } from 'vitest'
import { buildNamingInstruction, excerptForNaming, sanitizeFileStem } from '../src/main/ai-naming'

describe('sanitizeFileStem', () => {
  it('keeps a plain name unchanged', () => {
    expect(sanitizeFileStem('\u5b63\u5ea6\u8425\u6536\u603b\u7ed3')).toBe(
      '\u5b63\u5ea6\u8425\u6536\u603b\u7ed3',
    )
    expect(sanitizeFileStem('Q3 Revenue Review')).toBe('Q3 Revenue Review')
  })

  it('drops a path, keeping the last segment', () => {
    expect(sanitizeFileStem('reports/2026/Q3 \u8425\u6536')).toBe('Q3 \u8425\u6536')
    expect(sanitizeFileStem('C:\\Users\\me\\notes')).toBe('notes')
  })

  it('drops an extension the model volunteered', () => {
    expect(sanitizeFileStem('Q3 Revenue.md')).toBe('Q3 Revenue')
    expect(sanitizeFileStem('data.json')).toBe('data')
    // a name that is only an extension keeps its stem
    expect(sanitizeFileStem('.gitignore')).toBe('gitignore')
  })

  it('strips quotes, bullets and trailing punctuation', () => {
    expect(sanitizeFileStem('"Q3 Revenue"')).toBe('Q3 Revenue')
    expect(sanitizeFileStem('**Q3 Revenue**')).toBe('Q3 Revenue')
    expect(sanitizeFileStem('- Q3 Revenue')).toBe('Q3 Revenue')
    expect(sanitizeFileStem('`Q3 Revenue`')).toBe('Q3 Revenue')
  })

  it('replaces every character Windows forbids', () => {
    // a raw / would split the path; the rest are illegal on Windows
    expect(sanitizeFileStem('a/b')).toBe('b')
    expect(sanitizeFileStem('a?b')).toBe('a b')
    expect(sanitizeFileStem('a*b')).toBe('a b')
    expect(sanitizeFileStem('a:b')).toBe('a b')
    expect(sanitizeFileStem('a<b>c')).toBe('a b c')
    expect(sanitizeFileStem('a"b"c')).toBe('a b c')
    expect(sanitizeFileStem('a|b')).toBe('a b')
  })

  it('never returns a name made of separators or empty', () => {
    expect(sanitizeFileStem('///')).toBe('')
    expect(sanitizeFileStem('')).toBe('')
    expect(sanitizeFileStem('   ')).toBe('')
    expect(sanitizeFileStem('...')).toBe('')
  })

  it('does not hide the file behind a leading dot', () => {
    expect(sanitizeFileStem('.env')).toBe('env')
    expect(sanitizeFileStem('...hidden')).toBe('hidden')
  })

  it('escapes a reserved Windows device name', () => {
    expect(sanitizeFileStem('CON')).toBe('CON-file')
    expect(sanitizeFileStem('con')).toBe('con-file')
    expect(sanitizeFileStem('LPT1')).toBe('LPT1-file')
    // a name that merely starts with one is fine
    expect(sanitizeFileStem('CONTENT')).toBe('CONTENT')
  })

  it('drops a trailing dot or space, which Windows also forbids', () => {
    expect(sanitizeFileStem('Meeting notes.')).toBe('Meeting notes')
    expect(sanitizeFileStem('Strategy 2026. ')).toBe('Strategy 2026')
    expect(sanitizeFileStem('Review.')).toBe('Review')
    // a stem that was only dots and spaces still collapses to the empty, "no
    // name" answer rather than to a bare dot
    expect(sanitizeFileStem('. . .')).toBe('')
  })

  it('caps the length without cutting mid-word when it can avoid it', () => {
    const long = Array.from({ length: 40 }, (_, i) => `word${i}`).join(' ')
    const out = sanitizeFileStem(long)
    expect(out.length).toBeLessThanOrEqual(60)
    // every surviving token is whole: no "word1" cut into "word"
    for (const token of out.split(' ')) expect(token).toMatch(/^word\d+$/)
  })

  it('collapses a long unbroken token rather than dropping it', () => {
    const out = sanitizeFileStem('x'.repeat(200))
    expect(out).toHaveLength(60)
  })

  it('strips control characters a model can emit', () => {
    expect(sanitizeFileStem('Q3\u0000Revenue')).toBe('Q3 Revenue')
    expect(sanitizeFileStem('line1\nline2')).toBe('line1 line2')
  })

  it('is idempotent — cleaning twice changes nothing', () => {
    for (const raw of ['Q3/Revenue.md', '"a*b"', 'CON', '...x', 'a\nb', 'x'.repeat(90)]) {
      const once = sanitizeFileStem(raw)
      expect(sanitizeFileStem(once)).toBe(once)
    }
  })
})

describe('buildNamingInstruction', () => {
  it('asks for a bare name and names the language', () => {
    const out = buildNamingInstruction('some text', '\u7b80\u4f53\u4e2d\u6587')
    expect(out).toContain('file name')
    expect(out).toContain('\u7b80\u4f53\u4e2d\u6587')
    expect(out).toContain('some text')
    expect(out).toContain('no extension')
  })
})

describe('excerptForNaming', () => {
  it('leaves a short document alone', () => {
    expect(excerptForNaming('short text')).toBe('short text')
  })

  it('caps a long document and marks the cut', () => {
    const out = excerptForNaming('y'.repeat(10_000))
    expect(out.length).toBeLessThan(4100)
    expect(out.endsWith('…')).toBe(true)
  })

  it('collapses runs of blank lines so the budget goes to content', () => {
    expect(excerptForNaming('a\n\n\n\n\nb')).toBe('a\n\nb')
  })
})
