import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  KNOWLEDGE_MAX_ENTRIES,
  listKnowledgeEntries,
  pathForEntry,
  readKnowledgeEntry,
} from '../src/main/feature-knowledge-host'

/**
 * The knowledge kind's host: what a source offers, and that an entry can only be
 * one of those names, resolved inside the source.
 */

function source(): { folder: string; outside: string } {
  const root = mkdtempSync(join(tmpdir(), 'genoffice-knowledge-'))
  const folder = join(root, 'notes')
  mkdirSync(folder)
  writeFileSync(join(folder, 'style.md'), '# House style\n\nHeadings are sentence case.')
  writeFileSync(join(folder, 'checklist.md'), '## Checklist')
  writeFileSync(join(folder, 'logo.png'), 'not a document')
  mkdirSync(join(folder, 'archive'))
  const outside = join(root, 'secret.md')
  writeFileSync(outside, 'not in the source')
  return { folder, outside }
}

describe('knowledge host with file rules', () => {
  it('applies include and exclude to the listing', () => {
    const { folder } = source()
    const entries = listKnowledgeEntries({
      source: 'folder',
      path: folder,
      include: ['*.md'],
      exclude: ['check*'],
    })
    expect(entries.map((entry) => entry.name)).toEqual(['style.md'])
  })

  it('a rule-excluded name no longer resolves, like any unlisted one', () => {
    const { folder } = source()
    expect(() =>
      pathForEntry({ source: 'folder', path: folder, exclude: ['style.md'] }, 'style.md'),
    ).toThrow(/not in that source/)
    // and a matching include still resolves
    expect(
      pathForEntry({ source: 'folder', path: folder, include: ['style.md'] }, 'style.md'),
    ).toBe(join(folder, 'style.md'))
  })
})

describe('knowledge host', () => {
  it('lists only what the reader can open, and never a directory', () => {
    const { folder } = source()
    const entries = listKnowledgeEntries({ source: 'folder', path: folder })
    expect(entries.map((entry) => entry.name).sort()).toEqual(['checklist.md', 'style.md'])
    expect(entries.every((entry) => entry.bytes > 0)).toBe(true)
  })

  it('reads one offered entry through the document reader', async () => {
    const { folder } = source()
    const read = await readKnowledgeEntry({ source: 'folder', path: folder }, 'style.md')
    expect(read.name).toBe('style.md')
    expect(read.text).toContain('House style')
    expect(read.chars).toBeGreaterThan(0)
  })

  it('refuses a name the source did not offer', () => {
    const { folder } = source()
    expect(() => pathForEntry({ source: 'folder', path: folder }, 'logo.png')).toThrow(
      /not in that source/,
    )
    expect(() => pathForEntry({ source: 'folder', path: folder }, 'missing.md')).toThrow(
      /not in that source/,
    )
  })

  it('refuses to escape the folder, even with a name that looks local', () => {
    const { folder, outside } = source()
    for (const entry of ['../secret.md', '../../etc/passwd', folder + '/../secret.md']) {
      expect(() => pathForEntry({ source: 'folder', path: folder }, entry)).toThrow()
    }
    // and the file it was reaching for is still unreadable through the source
    expect(() => pathForEntry({ source: 'folder', path: folder }, outside)).toThrow()
  })

  it('reads a single-file source, and only under its own name', async () => {
    const { folder } = source()
    const path = join(folder, 'style.md')
    const entries = listKnowledgeEntries({ source: 'file', path })
    expect(entries).toEqual([{ name: 'style.md', bytes: expect.any(Number) }])
    const read = await readKnowledgeEntry({ source: 'file', path }, 'style.md')
    expect(read.text).toContain('House style')
    expect(() => pathForEntry({ source: 'file', path }, 'other.md')).toThrow(/not in that source/)
  })

  it('reports a source that is not there, rather than an empty list', () => {
    const { folder } = source()
    expect(() => listKnowledgeEntries({ source: 'folder', path: join(folder, 'nope') })).toThrow()
    expect(() => listKnowledgeEntries({ source: 'file', path: folder })).toThrow(/not a file/)
  })

  it('bounds the listing', () => {
    const root = mkdtempSync(join(tmpdir(), 'genoffice-many-'))
    for (let i = 0; i < KNOWLEDGE_MAX_ENTRIES + 12; i++) {
      writeFileSync(join(root, 'note-' + String(i).padStart(3, '0') + '.md'), 'line ' + i)
    }
    expect(listKnowledgeEntries({ source: 'folder', path: root })).toHaveLength(
      KNOWLEDGE_MAX_ENTRIES,
    )
  })
})
