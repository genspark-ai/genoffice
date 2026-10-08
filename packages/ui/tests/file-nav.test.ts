import { describe, expect, it, vi } from 'vitest'
import { FILE_NAV_SCHEME, createFileNav, fileNavHref, parseFileNavHref } from '../src/file-nav'

describe('filenav href generation', () => {
  it('keeps a plain POSIX path readable', () => {
    expect(fileNavHref('/Users/me/report.docx')).toBe('filenav:///Users/me/report.docx')
  })

  it('encodes spaces, parens and CJK so the markdown link regex accepts the href', () => {
    const href = fileNavHref('/Users/王小明/报告 (final v2).docx')
    // nothing left that the inline-link href grammar rejects: no space, no raw paren
    expect(href).not.toMatch(/[\s()]/)
    // encodeURIComponent already handles spaces/CJK; parens are encoded on top
    expect(href).toBe(
      'filenav:///Users/%E7%8E%8B%E5%B0%8F%E6%98%8E/%E6%8A%A5%E5%91%8A%20%28final%20v2%29.docx',
    )
  })

  it('round-trips through parseFileNavHref exactly', () => {
    const paths = [
      '/Users/me/report.docx',
      '/Users/王小明/报告 (final v2).docx',
      '/tmp/50%off (1).md',
      "/tmp/odd'!*.pdf",
    ]
    for (const path of paths) expect(parseFileNavHref(fileNavHref(path))).toBe(path)
  })
})

describe('filenav href parsing', () => {
  it('decodes an encoded href', () => {
    expect(parseFileNavHref('filenav:///a%20b/(1).docx')).toBe('/a b/(1).docx')
  })

  it('passes a raw hand-written href through when it has no valid percent escapes', () => {
    expect(parseFileNavHref('filenav:///a b/report.docx')).toBe('/a b/report.docx')
  })

  it('rejects other schemes and empty paths', () => {
    expect(parseFileNavHref('docnav://block/3')).toBeNull()
    expect(parseFileNavHref('filenav://')).toBeNull()
    expect(parseFileNavHref('https://example.com/x.pdf')).toBeNull()
  })
})

describe('createFileNav', () => {
  it('opens the decoded path of a matching href', () => {
    const openPath = vi.fn()
    const nav = createFileNav(openPath)
    expect(nav.scheme).toBe(FILE_NAV_SCHEME)
    nav.onNavigate(fileNavHref('/Users/me/季度报告 (2).xlsx'))
    expect(openPath).toHaveBeenCalledWith('/Users/me/季度报告 (2).xlsx')
  })

  it('ignores hrefs without a path or of another scheme', () => {
    const openPath = vi.fn()
    createFileNav(openPath).onNavigate('docnav://block/3')
    createFileNav(openPath).onNavigate('filenav://')
    expect(openPath).not.toHaveBeenCalled()
  })
})
