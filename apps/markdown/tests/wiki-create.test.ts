/**
 * Creating the note an unresolved `[[target]]` names (`wiki-create.ts`): where
 * the file lands, what it is named, and the two things it must never do —
 * overwrite what is already there, or reach outside the vault.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { createWikiNote } from '../src/main/wiki-create'

const dirs: string[] = []
afterEach(async () => {
  while (dirs.length > 0) await rm(dirs.pop()!, { recursive: true, force: true })
})

const scratch = async (name = 'vault'): Promise<string> => {
  const dir = await mkdtemp(join(tmpdir(), `wiki-create-${name}-`))
  dirs.push(dir)
  return dir
}

describe('createWikiNote', () => {
  it('a bare name lands beside the open document', async () => {
    const root = await scratch()
    const docDir = await scratch('doc')
    const made = await createWikiNote('Note', { root, fromDir: docDir })
    expect(made).toMatchObject({
      path: join(docDir, 'Note.md'),
      fragment: '',
      created: true,
    })
    expect(await readFile(join(docDir, 'Note.md'), 'utf8')).toBe('')
  })

  it('a directory target is vault-rooted and the directories are made as needed', async () => {
    const root = await scratch()
    const docDir = await scratch('doc')
    const made = await createWikiNote('folder/Deep/Note', { root, fromDir: docDir })
    expect(made?.path).toBe(join(root, 'folder', 'Deep', 'Note.md'))
    // the link keeps the directories: a bare `[[Note]]` could land elsewhere
    expect(made?.target).toBe('folder/Deep/Note')
    expect(await readFile(join(root, 'folder', 'Deep', 'Note.md'), 'utf8')).toBe('')
  })

  it('an untitled document creates at the vault root', async () => {
    const root = await scratch()
    const made = await createWikiNote('Note', { root, fromDir: '' })
    expect(made?.path).toBe(join(root, 'Note.md'))
  })

  it('a file already at the path is opened, never overwritten', async () => {
    const root = await scratch()
    await writeFile(join(root, 'Note.md'), 'precious words', 'utf8')
    const made = await createWikiNote('Note', { root, fromDir: '' })
    expect(made).toEqual({
      path: join(root, 'Note.md'),
      target: 'Note',
      fragment: '',
      created: false,
    })
    expect(await readFile(join(root, 'Note.md'), 'utf8')).toBe('precious words')
  })

  it('characters no file name can carry are replaced, segment by segment', async () => {
    const root = await scratch()
    const made = await createWikiNote('a:b/c?d', { root, fromDir: '' })
    expect(made?.path).toBe(join(root, 'a_b', 'c_d.md'))
    expect(made?.target).toBe('a_b/c_d')
  })

  it('a trailing dot or space is trimmed, so `Note.` cannot spell `Note..md`', async () => {
    const root = await scratch()
    const made = await createWikiNote('Note. ', { root, fromDir: '' })
    expect(made?.path).toBe(join(root, 'Note.md'))
  })

  it('an explicit extension is kept as written, the way resolution reads it', async () => {
    const root = await scratch()
    expect((await createWikiNote('Note.md', { root, fromDir: '' }))?.path).toBe(
      join(root, 'Note.md'),
    )
    expect((await createWikiNote('Figure 1.png', { root, fromDir: '' }))?.path).toBe(
      join(root, 'Figure 1.png'),
    )
  })

  it('the fragment rides back untouched', async () => {
    const root = await scratch()
    const made = await createWikiNote('Note#Some Heading', { root, fromDir: '' })
    expect(made).toMatchObject({ path: join(root, 'Note.md'), fragment: 'Some Heading' })
  })

  it.each([
    ['a same-file ref', '#Heading'],
    ['a climb out of the vault', '../evil'],
    ['a climb hidden in a path', 'folder/../evil'],
    ['an absolute path', '/etc/evil'],
  ])('refuses %s', async (_name, target) => {
    const root = await scratch()
    expect(await createWikiNote(target, { root, fromDir: '' })).toBeNull()
  })

  it('an existing directory where the note would go is not replaced', async () => {
    const root = await scratch()
    await mkdir(join(root, 'Note.md'))
    // the write fails and the caller (the IPC handler) turns that into ok:false
    await expect(createWikiNote('Note', { root, fromDir: '' })).rejects.toThrow()
  })
})
