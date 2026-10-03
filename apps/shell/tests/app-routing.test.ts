import { describe, expect, it } from 'vitest'

import { appForPath, renameStaysInApp } from '../src/main/app-routing'

describe('appForPath', () => {
  it('routes the text app at the extensions it opens', () => {
    expect(appForPath('/x/notes.md')).toBe('text')
    expect(appForPath('/x/notes.markdown')).toBe('text')
    expect(appForPath('/x/notes.txt')).toBe('text')
    expect(appForPath('/x/data.json')).toBe('text')
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
    expect(appForPath('/x/NOTES.MD')).toBe('text')
    expect(appForPath('/x/DATA.JSON')).toBe('text')
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
