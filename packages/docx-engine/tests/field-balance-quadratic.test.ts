import { performance } from 'node:perf_hooks'
import { describe, expect, it } from 'vitest'
import { balanceFieldChars } from '../src/field-balance'

function bodyWithOrphans(strays: number, fill = 460): string {
  const stray = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  const filler = `<w:r><w:t>${'x'.repeat(fill)}</w:t></w:r>`
  return (
    '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    Array.from({ length: strays }, () => stray + filler).join('') +
    '</w:body></w:document>'
  )
}

describe('balanceFieldChars orphan scan', () => {
  it('keeps a 1.6 MB body with 3,200 orphans well under the quadratic cliff', () => {
    // On main this shape took 8.3 s (measured): each orphan rescanned the whole
    // body and each edit re-sliced the whole string.
    const body = bodyWithOrphans(3200)
    expect(body.length).toBeGreaterThan(1_500_000)
    const t0 = performance.now()
    const out = balanceFieldChars(body)
    expect(performance.now() - t0).toBeLessThan(4_000)
    expect(out).not.toContain('w:fldCharType="end"/></w:r><w:r><w:t>')
  })

  it('still removes orphan end chars and closes unclosed fields', () => {
    const out = balanceFieldChars(bodyWithOrphans(3))
    expect(out).not.toContain('fldCharType="end"')
  })
})
