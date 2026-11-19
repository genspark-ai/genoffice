import { chmodSync, readFileSync, readdirSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { writeOutput } from '../src/fs'
import { tempDir } from './helpers'

vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return { ...actual, chmodSync: vi.fn(actual.chmodSync) }
})

const actual = await vi.importActual<typeof import('node:fs')>('node:fs')
const errno = (code: string) => Object.assign(new Error(`${code}: injected`), { code })
const dirs: string[] = []

afterEach(() => {
  vi.mocked(chmodSync).mockReset().mockImplementation(actual.chmodSync)
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

describe('writeOutput on a mount without permission support', () => {
  it.each(['ENOTSUP', 'EPERM', 'EACCES'])(
    'still writes the output when the staging file cannot take the target mode (%s)',
    (code) => {
      const dir = tempDir()
      dirs.push(dir)
      const target = join(dir, 'document.docx')
      actual.writeFileSync(target, 'old')
      actual.chmodSync(target, 0o600)
      vi.mocked(chmodSync).mockImplementationOnce(() => {
        throw errno(code)
      })
      writeOutput(target, 'new')
      expect(readFileSync(target, 'utf8')).toBe('new')
      expect(readdirSync(dir)).toEqual(['document.docx'])
      // the staging file was created with the target's mode, so a refused chmod widens nothing
      if (process.platform !== 'win32')
        expect(actual.statSync(target).mode & 0o777 & ~0o600).toBe(0)
    },
  )
})
