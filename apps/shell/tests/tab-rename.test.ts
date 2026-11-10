/**
 * @vitest-environment jsdom
 */
import { describe, expect, it } from 'vitest'
import { fileExtension, renamedFileName } from '../src/renderer/src/TabBar'

/**
 * Guards the tab-strip inline rename (apps/shell/src/renderer/src/TabBar.tsx):
 * the extension read for a rename must come from the basename, so a dot in a
 * directory name cannot leak a path separator into the rename target.
 */
describe('fileExtension', () => {
  it('reads the extension off the basename, not the whole path', () => {
    expect(fileExtension('C:\\Users\\me.v2\\Notes')).toBe('')
    expect(fileExtension('C:\\Users\\me.v2\\Notes.md')).toBe('md')
    expect(fileExtension('/home/me.v2/Notes')).toBe('')
    expect(fileExtension('/home/me/Notes.docx')).toBe('docx')
  })
})

/**
 * The rename input holds the whole name, so a name typed without an extension
 * used to be sent as typed and refused by the shell's rename gate: the app would
 * have stopped being able to open the file it had just renamed. These pin the
 * rule that an extensionless name keeps the file's current extension.
 */
describe('renamedFileName', () => {
  it('keeps the current extension when the typed name has none', () => {
    expect(renamedFileName('/home/me/Notes.md', 'Meeting')).toBe('Meeting.md')
    expect(renamedFileName('/home/me/Notes.txt', 'Meeting')).toBe('Meeting.txt')
    expect(renamedFileName('/home/me/data.json', 'config')).toBe('config.json')
  })

  it('takes a typed extension verbatim, which is how a surface is changed', () => {
    expect(renamedFileName('/home/me/Notes.md', 'Notes.txt')).toBe('Notes.txt')
    expect(renamedFileName('/home/me/Notes.txt', 'Notes.md')).toBe('Notes.md')
    expect(renamedFileName('/home/me/Notes.md', 'notes.MD')).toBe('notes.MD')
  })

  it('leaves a name alone when the file it renames has no extension either', () => {
    expect(renamedFileName('/home/me/README', 'Draft')).toBe('Draft')
  })

  it('does not turn a trailing dot into a doubled extension', () => {
    // "note." is a name the shell rejects outright (Windows strips trailing
    // dots); rewriting it to "note..md" would silently create a different file
    expect(renamedFileName('/home/me/Notes.md', 'note.')).toBe('note.')
  })

  it('reads the extension off the basename, not the whole path', () => {
    expect(renamedFileName('/home/me.v2/Notes.md', 'Meeting')).toBe('Meeting.md')
    expect(renamedFileName('C:\\Users\\me.v2\\Notes.md', 'Meeting')).toBe('Meeting.md')
  })
})
