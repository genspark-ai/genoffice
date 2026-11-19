import { describe, expect, it } from 'vitest'
import { isMoveSource, isUserVisibleFile } from '../src/main/file-targets'

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

describe('isMoveSource', () => {
  const dirs = new Set(['/roots', '/roots/sub'])
  const folders = new Set(['/roots', '/roots/sub', '/roots/deep/project', '/elsewhere/folder'])
  const sources = {
    insideAnyRoot: (p: string) => [...dirs].some((d) => p.startsWith(`${d}/`)),
    isAnyRoot: (p: string) => dirs.has(p),
    isDirectory: (p: string) => folders.has(p),
    trackedPaths: ['/recents/a.docx'],
  }

  it('accepts a file inside a root', () => {
    expect(isMoveSource('/roots/sub/b.xlsx', sources)).toBe(true)
  })

  it('accepts a file the UI can show outside every root (a recent from Downloads)', () => {
    expect(isMoveSource('/recents/a.docx', sources)).toBe(true)
  })

  it('rejects a file the UI never showed — the arbitrary-path case', () => {
    expect(isMoveSource('/Users/me/.ssh/id_rsa', sources)).toBe(false)
    expect(isMoveSource('/etc/hosts', sources)).toBe(false)
  })

  it('accepts a folder inside a root and rejects a root or an outside folder', () => {
    expect(isMoveSource('/roots/deep/project', sources)).toBe(true)
    expect(isMoveSource('/roots', sources)).toBe(false)
    expect(isMoveSource('/roots/sub', sources)).toBe(false)
    expect(isMoveSource('/elsewhere/folder', sources)).toBe(false)
  })

  it('tolerates non-string input defensively', () => {
    expect(isMoveSource(undefined as unknown as string, sources)).toBe(false)
    expect(isMoveSource('', sources)).toBe(false)
  })
})
