import { mkdtempSync, realpathSync, symlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseArgs } from '../src/args'
import { convertCommand } from '../src/commands/convert'
import { EXIT } from '../src/result'

const run = async (argv: string[], cwd: string) => {
  const ctx = { cwd, env: { ...process.env, GENOFFICE_ALLOWED_ROOTS: '' }, log: () => {} }
  try {
    await convertCommand.run(parseArgs(argv, new Set(['force'])), ctx as never)
  } catch (e) {
    return e as { code?: number; reason?: string; message: string }
  }
  return null
}

describe('convert refuses to write over its own input', () => {
  it('rejects --out equal to the input even with --force, also through a symlink', async () => {
    const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-convert-')))
    writeFileSync(join(dir, 'notes.md'), '# hi')
    symlinkSync(join(dir, 'notes.md'), join(dir, 'alias.md'))
    const direct = await run(['notes.md', '--to', 'html', '--out', 'notes.md', '--force'], dir)
    expect(direct).toMatchObject({ code: EXIT.usage, reason: 'invalid_argument' })
    const viaLink = await run(['notes.md', '--to', 'html', '--out', 'alias.md', '--force'], dir)
    expect(viaLink).toMatchObject({ code: EXIT.usage, reason: 'invalid_argument' })
  })
})
