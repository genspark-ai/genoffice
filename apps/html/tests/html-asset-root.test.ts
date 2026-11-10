import { describe, expect, it } from 'vitest'
import { isInDocDir } from '../src/main/asset-lifecycle'

describe('isInDocDir (html-asset:// serve gate)', () => {
  it('allows sibling images of filesystem-root documents', () => {
    expect(isInDocDir('/a.png', '/', '/')).toBe(true)
    expect(isInDocDir('D:\\a.png', 'D:\\', '\\')).toBe(true)
  })

  it('allows siblings of nested documents and rejects prefix look-alikes', () => {
    expect(isInDocDir('/docs/a.png', '/docs', '/')).toBe(true)
    expect(isInDocDir('/docs-evil/a.png', '/docs', '/')).toBe(false)
    expect(isInDocDir('C:\\docs\\a.png', 'C:\\docs', '\\')).toBe(true)
    expect(isInDocDir('C:\\docs-evil\\a.png', 'C:\\docs', '\\')).toBe(false)
  })

  it('never serves the directory itself', () => {
    expect(isInDocDir('/docs', '/docs', '/')).toBe(false)
    expect(isInDocDir('/', '/', '/')).toBe(false)
  })
})
