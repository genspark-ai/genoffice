import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { duplicateDocumentFile } from '../src/main/tab-duplicate'

let directory: string
beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'tab-duplicate-'))
})
afterEach(async () => {
  await rm(directory, { recursive: true, force: true })
})

function callbacks() {
  return { copySuffix: 'Copy', recordRecentFile: vi.fn(), onError: vi.fn() }
}

describe('shared document duplication', () => {
  it('copies through the existing atomic helper and records the new file', async () => {
    const source = join(directory, 'report.docx')
    await writeFile(source, 'saved content')
    const deps = callbacks()
    const target = await duplicateDocumentFile(source, deps)
    expect(target).toBe(join(directory, 'report Copy.docx'))
    expect(await readFile(target!, 'utf8')).toBe('saved content')
    expect(deps.recordRecentFile).toHaveBeenCalledWith(target)
    expect(deps.onError).not.toHaveBeenCalled()
  })

  it('keeps existing copies and uses the localized suffix for the next name', async () => {
    const source = join(directory, 'report.docx')
    const existing = join(directory, 'report Kopie.docx')
    await writeFile(source, 'source')
    await writeFile(existing, 'existing copy')
    const deps = { ...callbacks(), copySuffix: 'Kopie' }
    expect(await duplicateDocumentFile(source, deps)).toBe(join(directory, 'report Kopie 2.docx'))
    expect(await readFile(existing, 'utf8')).toBe('existing copy')
  })

  it('reports a copy failure to the shared user-facing error handler without recording a file', async () => {
    const deps = callbacks()
    expect(await duplicateDocumentFile(join(directory, 'missing.docx'), deps)).toBeUndefined()
    expect(deps.onError).toHaveBeenCalledWith(expect.objectContaining({ code: 'ENOENT' }))
    expect(deps.recordRecentFile).not.toHaveBeenCalled()
  })
})
