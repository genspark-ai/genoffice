import { describe, expect, it } from 'vitest'
import { safeName } from '../src/mcp/files'

// CJK samples are written with \u escapes on purpose: literal Han characters are
// rejected outside locale files, and comments must stay ASCII/English.
const REPORT_CJK = '\u5831\u544a\u66f8 2.docx'

describe('safeName', () => {
  it('keeps letters of any script and still reduces a name to one safe segment', () => {
    // an ASCII-only class folded these to r_sum_.pdf, collapsing distinct
    // upload names onto the same path segment
    expect(safeName('résumé.pdf')).toBe('résumé.pdf')
    expect(safeName('rçsumé.pdf')).toBe('rçsumé.pdf')
    expect(safeName('résumé.pdf')).not.toBe(safeName('rçsumé.pdf'))
    expect(safeName(REPORT_CJK)).toBe(REPORT_CJK)
    expect(safeName('Q3 rapport (final).docx')).toBe('Q3 rapport (final).docx')
    // path safety is unchanged
    expect(safeName('../../etc/passwd')).toBe('passwd')
    expect(safeName('.bashrc')).toBe('bashrc')
    expect(safeName('we:ird*name?.txt')).toBe('we_ird_name_.txt')
  })

  it('keeps combining marks so an NFD name is not folded', () => {
    // macOS volumes and many mac clients emit NFD; dropping the marks folded
    // this onto re_sume_.pdf
    const nfd = 'résumé.pdf'.normalize('NFD')
    expect(nfd).not.toBe('résumé.pdf')
    expect(safeName(nfd)).toBe(nfd)
    // and it stays distinct from the name it used to collide with
    expect(safeName(nfd)).not.toBe('re_sume_.pdf')
  })
})
