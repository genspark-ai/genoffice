import { closeSync, ftruncateSync, mkdtempSync, openSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { KB_READ_MAX_BYTES, readKnowledgeFilePage, readKnowledgeFileText } from '../src/main/knowledge-base'

let dir: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-kb-read-'))
})
afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('readKnowledgeFileText', () => {
  it('refuses a file over the byte cap before parsing it', async () => {
    // sparse file: the cap check must fire without any bytes on disk, the way
    // a real oversized attachment would be refused before parse
    const big = join(dir, 'huge.txt')
    const fd = openSync(big, 'w')
    ftruncateSync(fd, KB_READ_MAX_BYTES + 1)
    closeSync(fd)
    await expect(readKnowledgeFileText(big)).rejects.toThrow(/too large/)
  })

  it('pages a small file through the shared chunk size', async () => {
    const note = join(dir, 'note.txt')
    writeFileSync(note, 'first chunk second chunk')
    const page = await readKnowledgeFilePage(note, 0)
    expect(page.totalChars).toBe('first chunk second chunk'.length)
    expect(page.text).toBe('first chunk second chunk')
    expect(page.offset).toBe(0)
  })
})
