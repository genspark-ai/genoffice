import { describe, expect, it } from 'vitest'
import { isUserVisibleFile } from '../src/main/file-targets'

describe('isUserVisibleFile', () => {
  const sources = {
    insideAnyRoot: (p: string) => p.startsWith('/roots/'),
    trackedPaths: ['/recents/a.docx', '/starred/b.pdf', '/open/c.pptx'],
  }

  it('accepts anything inside a folder root', () => {
    expect(isUserVisibleFile('/roots/sub/deep/file.xlsx', sources)).toBe(true)
  })

  it('accepts a tracked path outside every root (a recent from Downloads)', () => {
    expect(isUserVisibleFile('/recents/a.docx', sources)).toBe(true)
    expect(isUserVisibleFile('/starred/b.pdf', sources)).toBe(true)
    expect(isUserVisibleFile('/open/c.pptx', sources)).toBe(true)
  })

  it('rejects paths the UI never showed — the SSRF-class arbitrary-file case', () => {
    expect(isUserVisibleFile('/Users/me/.ssh/id_rsa', sources)).toBe(false)
    expect(isUserVisibleFile('/etc/hosts', sources)).toBe(false)
    expect(isUserVisibleFile('/recents/', sources)).toBe(false) // prefix, not member
    expect(isUserVisibleFile('', sources)).toBe(false)
  })

  it('tolerates non-string input defensively', () => {
    expect(isUserVisibleFile(undefined as unknown as string, sources)).toBe(false)
  })
})
