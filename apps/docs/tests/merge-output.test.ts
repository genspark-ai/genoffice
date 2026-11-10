import { describe, expect, it } from 'vitest'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { mergeOutputFileName, mergeOutputPath } from '../src/main/merge-output'

/**
 * Where a merged document is written: the name comes from the record, so the
 * main process keeps only its last part and gives two records of the same
 * name two different files. genoffice#2097.
 */

describe('mergeOutputFileName', () => {
  it('adds the docx extension', () => {
    expect(mergeOutputFileName('Ada')).toBe('Ada.docx')
  })

  it('keeps a name that already carries the extension', () => {
    expect(mergeOutputFileName('Ada.docx')).toBe('Ada.docx')
  })

  it('keeps only the last part of a name, so it cannot carry a folder', () => {
    expect(mergeOutputFileName('../../etc/passwd')).toBe('passwd.docx')
  })

  it('replaces the characters a file name cannot carry', () => {
    expect(mergeOutputFileName('a:b*c')).toBe('a_b_c.docx')
  })

  it('falls back to a usable name', () => {
    expect(mergeOutputFileName('   ')).toBe('document.docx')
  })
})

describe('mergeOutputPath', () => {
  it('gives two records of the same name two different files', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-merge-out-'))
    await writeFile(join(dir, 'Ada.docx'), 'first')
    const second = mergeOutputPath(dir, 'Ada')
    expect(basename(second)).toBe('Ada-2.docx')
  })

  it('keeps the file inside the folder the user chose', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-merge-out-'))
    const path = mergeOutputPath(dir, '../escape')
    expect(path.startsWith(dir)).toBe(true)
  })
})
