// @vitest-environment node
/**
 * The no-clobber move of pdf:auto-rename on a folder where neither a hard link
 * nor the platform copy works (gvfs SMB, WebDAV). Runs in node: under jsdom a
 * mock of node:fs does not reach the main process module.
 */
import {
  constants,
  copyFileSync,
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return { ...actual, copyFileSync: vi.fn(actual.copyFileSync) }
})

vi.mock('electron', () => ({
  app: { on: vi.fn(), whenReady: vi.fn(() => new Promise(() => {})) },
  dialog: {},
  shell: {},
  BrowserWindow: class {},
  WebContentsView: class {},
  ipcMain: { handle: vi.fn(), on: vi.fn(), removeHandler: vi.fn() },
}))

import { movePdfFileNoClobber } from '../src/main/pdf-main'

const tempDirs = new Set<string>()

function makePdfFile(name = 'Untitled PDF.pdf'): string {
  const dir = mkdtempSync(join(tmpdir(), 'pdf-auto-rename-copy-'))
  tempDirs.add(dir)
  const path = join(dir, name)
  writeFileSync(path, '%PDF-1.4\n%%EOF\n')
  return path
}

afterEach(() => {
  vi.mocked(copyFileSync).mockReset()
  for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true })
  tempDirs.clear()
})

describe('pdf auto-rename on a mount without hard links', () => {
  it('moves where neither a hard link nor the platform copy works', () => {
    const source = makePdfFile()
    const target = join(source, '..', 'Plain Copy.pdf')
    const original = readFileSync(source)
    // libuv on a mount that refuses fchmod: the new file is created, removed, then the error thrown
    vi.mocked(copyFileSync).mockImplementation((_from, to, mode) => {
      const path = String(to)
      if (((mode ?? 0) & constants.COPYFILE_EXCL) !== 0 && existsSync(path))
        throw Object.assign(new Error('EEXIST: injected'), { code: 'EEXIST' })
      writeFileSync(path, '')
      rmSync(path)
      throw Object.assign(new Error('ENOTSUP: injected'), { code: 'ENOTSUP' })
    })

    const result = movePdfFileNoClobber(source, target, {
      link: () => {
        throw Object.assign(new Error('hard links unavailable'), { code: 'EPERM' })
      },
    })

    expect(result).toBe('moved')
    expect(existsSync(source)).toBe(false)
    expect(readFileSync(target)).toEqual(original)
  })
})
