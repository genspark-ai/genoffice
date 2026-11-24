import { describe, expect, it } from 'vitest'
import { linkSourcePath } from '../src/main/link-source'

describe('the local file a link names', () => {
  it('is the path as written, or the path of a file: target', () => {
    expect(linkSourcePath('/Users/a/part.docx', null)).toBe('/Users/a/part.docx')
    expect(linkSourcePath('file:///Users/a/My%20photo.png', null)).toBe('/Users/a/My photo.png')
  })

  it('is read against the document folder when relative', () => {
    expect(linkSourcePath('parts/one.docx', '/Users/a/docs/host.docx')).toBe(
      '/Users/a/docs/parts/one.docx',
    )
    expect(linkSourcePath('parts/one.docx', null)).toBeNull()
  })

  it('is nothing for a web address or an empty source', () => {
    expect(linkSourcePath('https://example.com/a.png', '/Users/a/host.docx')).toBeNull()
    expect(linkSourcePath('  ', '/Users/a/host.docx')).toBeNull()
  })

  it('is nothing for a file on another machine', () => {
    for (const source of [
      '\\\\host\\share\\a.docx',
      '//host/share/a.docx',
      '\\/host/share/a.docx',
      'file://host/share/a.png',
      'FILE:\\\\host\\share\\a.png',
      'file:////host/share/a.png',
    ])
      expect(linkSourcePath(source, '/Users/a/host.docx'), source).toBeNull()
    expect(linkSourcePath('file://localhost/Users/a/a.png', null)).toBe('/Users/a/a.png')
    expect(linkSourcePath('file:/Users/a/a.png', null)).toBe('/Users/a/a.png')
  })
})
