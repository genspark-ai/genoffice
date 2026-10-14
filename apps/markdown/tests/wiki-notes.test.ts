import { describe, expect, it } from 'vitest'

import { join } from 'node:path'

import { listWikiNotes, MAX_NOTES } from '../src/main/wiki-notes'
import type { WikiFs } from '../src/main/wiki-resolve'

/** a flat in-memory tree: `list` answers files and dirs per directory */
function fakeFs(tree: Record<string, string[]>): WikiFs {
  const norm = (dir: string): string => dir.replace(/\/+$/, '') || '/'
  return {
    exists: () => false,
    isFile: () => true,
    list: (dir) => {
      const entries = tree[norm(dir)]
      if (!entries) throw new Error('missing dir')
      return {
        files: entries.filter((e) => e.includes('.')),
        dirs: entries.filter((e) => !e.includes('.')),
      }
    },
  }
}

const ROOT = '/vault'

describe('listWikiNotes', () => {
  it('walks breadth-first, so a shallower note precedes a deeper one', () => {
    const notes = listWikiNotes(
      ROOT,
      fakeFs({
        '/vault': ['Deep', 'Top.md', 'cover.png'],
        '/vault/Deep': ['Buried.md'],
      }),
    )
    expect(notes.map((n) => n.relPath)).toEqual(['Top.md', 'cover.png', join('Deep', 'Buried.md')])
  })

  it('names the stem and the file', () => {
    const notes = listWikiNotes(ROOT, fakeFs({ '/vault': ['My Note.md'] }))
    expect(notes[0]).toMatchObject({
      name: 'My Note',
      file: 'My Note.md',
      path: join(ROOT, 'My Note.md'),
    })
  })

  it('skips dependency and app dirs', () => {
    const notes = listWikiNotes(
      ROOT,
      fakeFs({
        '/vault': ['node_modules', '.obsidian', 'Real.md'],
        '/vault/node_modules': ['Junk.md'],
        '/vault/.obsidian': ['Config.md'],
      }),
    )
    expect(notes.map((n) => n.file)).toEqual(['Real.md'])
  })

  it('offers only suggestable file types', () => {
    const notes = listWikiNotes(ROOT, fakeFs({ '/vault': ['note.md', 'data.xlsx', 'app.exe'] }))
    expect(notes.map((n) => n.file)).toEqual(['note.md'])
  })

  it('caps the listing and stops walking once full', () => {
    const tree: Record<string, string[]> = { '/vault': [] }
    for (let i = 0; i < MAX_NOTES + 300; i++) tree['/vault'].push(`Note ${i}.md`)
    const notes = listWikiNotes(ROOT, fakeFs(tree))
    expect(notes).toHaveLength(MAX_NOTES)
  })

  it('returns nothing without a root and survives an unreadable directory', () => {
    expect(listWikiNotes('', fakeFs({ '/vault': ['A.md'] }))).toEqual([])
    const notes = listWikiNotes(ROOT, {
      exists: () => false,
      isFile: () => true,
      list: () => {
        throw new Error('EACCES')
      },
    })
    expect(notes).toEqual([])
  })
})
