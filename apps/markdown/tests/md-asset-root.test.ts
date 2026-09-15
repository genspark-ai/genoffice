import { describe, expect, it } from 'vitest'
import { isInDocDir } from '../src/main/asset-lifecycle'

describe('isInDocDir (md-asset:// serve gate)', () => {
  it('allows sibling images of filesystem-root documents', () => {
    expect(isInDocDir('/a.png', '/', '/')).toBe(true)
    expect(isInDocDir('C:\\a.png', 'C:\\', '\\')).toBe(true)
  })

  it('allows siblings of nested documents and rejects the rest', () => {
    expect(isInDocDir('/docs/a.png', '/docs', '/')).toBe(true)
    expect(isInDocDir('/other/a.png', '/docs', '/')).toBe(false)
    // Prefix attacks: a sibling directory sharing the name prefix is outside
    expect(isInDocDir('/docs-evil/a.png', '/docs', '/')).toBe(false)
    expect(isInDocDir('C:\\docs\\a.png', 'C:\\docs', '\\')).toBe(true)
    expect(isInDocDir('C:\\docs-evil\\a.png', 'C:\\docs', '\\')).toBe(false)
  })

  it('never serves the directory itself', () => {
    expect(isInDocDir('/docs', '/docs', '/')).toBe(false)
    expect(isInDocDir('/', '/', '/')).toBe(false)
  })
})
