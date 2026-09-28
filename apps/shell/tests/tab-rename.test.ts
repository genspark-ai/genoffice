/**
 * @vitest-environment jsdom
 */
import { describe, expect, it } from 'vitest'
import { fileExtension } from '../src/renderer/src/TabBar'

/**
 * Guards the tab-strip inline rename (apps/shell/src/renderer/src/TabBar.tsx):
 * the extension appended to the new name must come from the basename, so a dot
 * in a directory name cannot leak a path separator into the rename target.
 */
describe('fileExtension', () => {
  it('reads the extension off the basename, not the whole path', () => {
    expect(fileExtension('C:\\Users\\me.v2\\Notes')).toBe('')
    expect(fileExtension('C:\\Users\\me.v2\\Notes.md')).toBe('md')
    expect(fileExtension('/home/me.v2/Notes')).toBe('')
    expect(fileExtension('/home/me/Notes.docx')).toBe('docx')
  })

  it('builds a rename target with no path separator in it', () => {
    const value = 'Meeting'
    const ext = fileExtension('C:\\Users\\me.v2\\Notes')
    expect(ext ? `${value}.${ext}` : value).toBe('Meeting')
  })

  it('still preserves a real extension', () => {
    const value = 'Meeting'
    const ext = fileExtension('C:\\Users\\me.v2\\Notes.md')
    expect(ext ? `${value}.${ext}` : value).toBe('Meeting.md')
  })
})
