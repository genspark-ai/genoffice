import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { flagString, parseArgs } from '../src/args'
import { resolveOutput } from '../src/fs'
import { EXIT } from '../src/result'

describe('resolveOutput', () => {
  it('refuses an empty --out rather than falling back onto the input file', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-out-'))
    const input = join(dir, 'deck.pptx')
    writeFileSync(input, 'source')
    const ctx = { cwd: dir, env: { ...process.env, GENOFFICE_ALLOWED_ROOTS: '' } }
    // `--out=` is an empty value, not an absent flag
    const args = parseArgs(['slides', 'apply', 'deck.pptx', '--ops', 'ops.json', '--out='])
    const out = flagString(args, 'out')
    expect(out).toBe('')
    // the fallback is the input file, so accepting '' would overwrite the source
    let err: unknown
    try {
      resolveOutput(out, ctx, { fallback: input })
    } catch (e) {
      err = e
    }
    expect(err).toMatchObject({ code: EXIT.usage, reason: 'missing_argument' })
    expect((err as Error).message).toMatch(/missing --out <path>/)
    // an absent --out still edits in place, and a real one still redirects
    expect(resolveOutput(undefined, ctx, { fallback: input })).toBe(input)
    expect(resolveOutput('out.pptx', ctx, { fallback: input })).toBe(join(dir, 'out.pptx'))
  })
})
