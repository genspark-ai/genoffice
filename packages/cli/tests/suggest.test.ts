import { readFileSync, statSync, writeFileSync, chmodSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { writeOutput } from '../src/fs'
import { CommandRegistry } from '../src/registry'
import { didYouMean } from '../src/suggest'
import { run, tempDir } from './helpers'

describe('didYouMean', () => {
  it('finds one-edit typos and transpositions, and nothing when far away', () => {
    expect(didYouMean('sett', ['set', 'get', 'add'])).toBe('set')
    expect(didYouMean('blod', ['bold', 'italic'])).toBe('bold')
    expect(didYouMean('set_cel', ['setText', 'set_cell', 'set_formula'])).toBe('set_cell')
    expect(didYouMean('setTxt', ['setText', 'setFill'])).toBe('setText')
    expect(didYouMean('frobnicate', ['info', 'convert', 'create'])).toBeUndefined()
    expect(didYouMean('x', ['info'])).toBeUndefined()
  })
})

describe('did-you-mean in errors', () => {
  it('names the closest command and option', async () => {
    const cmd = (await run(['covnert', '--json'])).json()
    expect(cmd.error).toBe('unknown_command')
    expect(cmd.suggestion).toContain('did you mean `convert`?')
    const opt = (await run(['info', 'x.docx', '--pasword', 'p', '--json'])).json()
    expect(opt.error).toBe('unknown_option')
    expect(opt.suggestion).toContain('did you mean `--password`?')
  })
})

describe('warnings channel', () => {
  const registry = new CommandRegistry().register({
    name: 'noisy',
    summary: 'test command',
    usage: 'noisy',
    async run(_args, ctx) {
      ctx.warn({ code: 'from_ctx', message: 'collected later', suggestion: 'ignore me' })
      return { summary: 'done', warnings: [{ code: 'from_result', message: 'returned' }] }
    },
  })
  it('merges result and context warnings into the envelope and the human output', async () => {
    const j = (await run(['noisy', '--json'], { registry })).json()
    expect(j.status).toBe('ok')
    expect(j.warnings.map((w: { code: string }) => w.code)).toEqual(['from_result', 'from_ctx'])
    const h = await run(['noisy'], { registry })
    expect(h.stdout).toContain('  warning: returned')
    expect(h.stdout).toContain('  warning: collected later (ignore me)')
  })
})

describe('writeOutput', () => {
  it('replaces the file whole and keeps its mode', () => {
    const file = join(tempDir(), 'doc.bin')
    writeFileSync(file, 'old')
    chmodSync(file, 0o600)
    writeOutput(file, 'new content')
    expect(readFileSync(file, 'utf-8')).toBe('new content')
    expect(statSync(file).mode & 0o777).toBe(0o600)
    const fresh = join(tempDir(), 'fresh.bin')
    writeOutput(fresh, Buffer.from([1, 2, 3]))
    expect(readFileSync(fresh)).toEqual(Buffer.from([1, 2, 3]))
  })
})
