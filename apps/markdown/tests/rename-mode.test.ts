import { describe, expect, it } from 'vitest'
import { renameAction } from '../src/shared/rename-mode'

describe('renameAction', () => {
  it('keeps the buffer when the extension does not change the surface', () => {
    expect(renameAction('/a/note.md', '/a/renamed.md', false)).toBe('keep')
    expect(renameAction('/a/note.md', '/a/renamed.markdown', true)).toBe('keep')
  })

  it('keeps the buffer for a plain-name move', () => {
    expect(renameAction('/a/note.md', '/b/note.md', true)).toBe('keep')
  })

  it('reloads a clean markdown file renamed to txt', () => {
    expect(renameAction('/a/note.md', '/a/note.txt', false)).toBe('reload')
  })

  it('reloads a clean markdown file renamed to json', () => {
    expect(renameAction('/a/note.md', '/a/note.json', false)).toBe('reload')
  })

  it('reloads a clean txt file renamed to md', () => {
    expect(renameAction('/a/note.txt', '/a/note.md', false)).toBe('reload')
  })

  it('reloads a clean json file renamed to md', () => {
    expect(renameAction('/a/data.json', '/a/data.md', false)).toBe('reload')
  })

  it('refuses a cross-surface rename that would discard unsaved edits', () => {
    expect(renameAction('/a/note.md', '/a/note.txt', true)).toBe('block-dirty')
    expect(renameAction('/a/note.md', '/a/note.json', true)).toBe('block-dirty')
    expect(renameAction('/a/note.txt', '/a/note.md', true)).toBe('block-dirty')
  })

  it('keeps plain and json apart from markdown but not from each other', () => {
    // both are edited as source, so the buffer survives the switch
    expect(renameAction('/a/notes.txt', '/a/notes.json', true)).toBe('keep')
    expect(renameAction('/a/notes.json', '/a/notes.txt', true)).toBe('keep')
  })

  it('follows the new extension for an untitled document', () => {
    expect(renameAction(null, '/a/note.txt', false)).toBe('reload')
    // an untitled document has nothing on disk to lose, so dirty is harmless
    expect(renameAction(null, '/a/note.txt', true)).toBe('reload')
  })

  it('is not fooled by case or by a dot inside the name', () => {
    expect(renameAction('/a/note.MD', '/a/note.TXT', false)).toBe('reload')
    expect(renameAction('/a/a.b.md', '/a/a.b.txt', false)).toBe('reload')
    expect(renameAction('/a/note.md', '/a/note.md.txt', false)).toBe('reload')
  })
})
