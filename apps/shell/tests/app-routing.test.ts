import { describe, expect, it } from 'vitest'

import { TEXT_RE, appForPath, renameStaysInApp } from '../src/main/app-routing'

describe('appForPath', () => {
  it('routes the markdown app at every extension it opens', () => {
    expect(appForPath('/x/notes.md')).toBe('markdown')
    expect(appForPath('/x/notes.markdown')).toBe('markdown')
    expect(appForPath('/x/notes.txt')).toBe('markdown')
    expect(appForPath('/x/data.json')).toBe('markdown')
  })

  it('leaves the other apps on their own extensions', () => {
    expect(appForPath('/x/a.docx')).toBe('docs')
    expect(appForPath('/x/a.xlsx')).toBe('sheets')
    expect(appForPath('/x/a.pptx')).toBe('slides')
    expect(appForPath('/x/a.pdf')).toBe('pdf')
    expect(appForPath('/x/a.html')).toBe('html')
  })

  it('has no app for an extension nothing opens', () => {
    expect(appForPath('/x/a.xyz')).toBeUndefined()
    expect(appForPath('/x/noextension')).toBeUndefined()
  })

  it('is case-insensitive, because the filesystem is', () => {
    expect(appForPath('/x/NOTES.MD')).toBe('markdown')
    expect(appForPath('/x/DATA.JSON')).toBe('markdown')
  })
})

/**
 * The open route keys on TEXT_RE, and a mismatch between the two is exactly the
 * bug this file was written for: the open dialog listed .txt/.json and the route
 * still tested a markdown-only pattern, so a picked file fell through to the
 * unsupported-file notice. Pinning the regex the route uses keeps the offered
 * extensions and the routed extensions from drifting apart again.
 */
describe("the open route's text-app pattern", () => {
  it('claims every extension the open dialog offers for the text app', () => {
    for (const ext of ['md', 'markdown', 'txt', 'json']) {
      expect(TEXT_RE.test(`/x/notes.${ext}`)).toBe(true)
    }
  })

  it('claims nothing else, so an unknown file still gets the unsupported notice', () => {
    for (const ext of ['docx', 'xlsx', 'pptx', 'pdf', 'html', 'xyz']) {
      expect(TEXT_RE.test(`/x/notes.${ext}`)).toBe(false)
    }
    expect(TEXT_RE.test('/x/noextension')).toBe(false)
  })
})

describe('renameStaysInApp', () => {
  it('refuses a legal name in an extension no app routes to', () => {
    // the character rules accept this: it is the extension that turns an
    // openable file into an unopenable one
    expect(renameStaysInApp('/x/notes.md', 'notes.md')).toBe(true)
    expect(renameStaysInApp('/x/notes.md', 'notes.xyz')).toBe(false)
  })

  it('allows a rename that stays inside the text app', () => {
    expect(renameStaysInApp('/x/notes.md', 'notes.markdown')).toBe(true)
    expect(renameStaysInApp('/x/notes.txt', 'notes.md')).toBe(true)
    expect(renameStaysInApp('/x/data.json', 'data.txt')).toBe(true)
  })

  it('keeps a rename inside every other app too', () => {
    expect(renameStaysInApp('/x/a.docx', 'a.docx')).toBe(true)
    expect(renameStaysInApp('/x/a.docx', 'a.pdf')).toBe(false)
    expect(renameStaysInApp('/x/a.xlsx', 'a.csv')).toBe(true)
    expect(renameStaysInApp('/x/a.png', 'a.png')).toBe(true)
  })

  it('stays put when the file was never openable to begin with', () => {
    // no app to stay in, so this gate is not what blocks it
    expect(renameStaysInApp('/x/a.xyz', 'a.anything')).toBe(true)
  })
})
