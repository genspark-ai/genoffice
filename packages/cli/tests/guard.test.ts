import { existsSync, mkdirSync, readFileSync, symlinkSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { auditLogPath } from '../src/audit'
import { redactArgv } from '../src/cli'
import { allowedRoots, realizedPath } from '../src/fs'
import { run, tempDir } from './helpers'

const env = (over: Record<string, string>) => ({
  ...process.env,
  GENOFFICE_AUDIT_LOG: 'off',
  ...over,
})

describe('GENOFFICE_ALLOWED_ROOTS', () => {
  it('is unrestricted when unset and blank', () => {
    expect(allowedRoots({})).toBeNull()
    expect(allowedRoots({ GENOFFICE_ALLOWED_ROOTS: '  ' })).toBeNull()
  })

  it('refuses inputs and outputs outside the roots (exit 2, roots in detail)', async () => {
    const inside = tempDir()
    const outside = tempDir()
    const csv = join(outside, 'a.csv')
    writeFileSync(csv, 'x,y\n1,2\n')
    const e = env({ GENOFFICE_ALLOWED_ROOTS: inside })
    const read = await run(['info', csv, '--json'], { env: e })
    expect(read.code).toBe(2)
    expect(read.json().message).toContain('refusing to read')
    expect(read.json().detail.allowed_roots).toEqual([realizedPath(inside)])

    writeFileSync(join(inside, 'b.csv'), 'x,y\n1,2\n')
    const write = await run(
      [
        'convert',
        join(inside, 'b.csv'),
        '--to',
        'xlsx',
        '--out',
        join(outside, 'b.xlsx'),
        '--json',
      ],
      { env: e },
    )
    expect(write.code).toBe(2)
    expect(write.json().message).toContain('refusing to write')
    expect(existsSync(join(outside, 'b.xlsx'))).toBe(false)

    const ok = await run(
      [
        'convert',
        join(inside, 'b.csv'),
        '--to',
        'xlsx',
        '--out',
        join(inside, 'sub/b.xlsx'),
        '--json',
      ],
      { env: e },
    )
    expect(ok.code).toBe(0)
    expect(existsSync(join(inside, 'sub/b.xlsx'))).toBe(true)
  })

  it.skipIf(process.platform === 'win32')('follows symlinks before comparing', async () => {
    const inside = tempDir()
    const outside = tempDir()
    writeFileSync(join(outside, 'real.csv'), 'x\n1\n')
    symlinkSync(join(outside, 'real.csv'), join(inside, 'link.csv'))
    mkdirSync(join(inside, 'escape-dir'))
    symlinkSync(outside, join(inside, 'escape'), 'dir')
    const e = env({ GENOFFICE_ALLOWED_ROOTS: inside })
    expect((await run(['info', join(inside, 'link.csv'), '--json'], { env: e })).code).toBe(2)
    writeFileSync(join(inside, 'b.csv'), 'x\n1\n')
    const viaLink = await run(
      [
        'convert',
        join(inside, 'b.csv'),
        '--to',
        'xlsx',
        '--out',
        join(inside, 'escape/new/b.xlsx'),
      ],
      { env: e },
    )
    expect(viaLink.code).toBe(2)
    expect(realizedPath(join(inside, 'escape/new/b.xlsx'))).toBe(
      join(realizedPath(outside), 'new/b.xlsx'),
    )
  })

  it('keeps names that merely start with ".." inside the root', async () => {
    const inside = tempDir()
    mkdirSync(join(inside, '..hidden'))
    writeFileSync(join(inside, '..hidden', 'a.csv'), 'x\n1\n')
    const e = env({ GENOFFICE_ALLOWED_ROOTS: inside })
    expect(
      (await run(['info', join(inside, '..hidden', 'a.csv'), '--json'], { env: e })).code,
    ).toBe(0)
  })

  it('applies to image files referenced from slide ops', async () => {
    const inside = tempDir()
    const outside = tempDir()
    writeFileSync(join(outside, 'pic.png'), Buffer.from('89504e470d0a1a0a', 'hex'))
    const ops = join(inside, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([{ op: 'addPicture', target: { slide: 0 }, bytes: join(outside, 'pic.png') }]),
    )
    const e = env({ GENOFFICE_ALLOWED_ROOTS: inside })
    const r = await run(
      ['create', '--type', 'pptx', '--ops', ops, '--out', join(inside, 'deck.pptx'), '--json'],
      { env: e },
    )
    expect(r.code).toBe(2)
    expect(r.json().message).toContain('refusing to read')
  })

  it('does not create the output folder on a dry run', async () => {
    const dir = tempDir()
    const table = join(dir, 't.json')
    writeFileSync(
      table,
      JSON.stringify([
        ['a', 'b'],
        [1, 2],
      ]),
    )
    const xlsx = join(dir, 't.xlsx')
    expect(
      (await run(['create', '--type', 'xlsx', '--from', table, '--out', xlsx], { env: env({}) }))
        .code,
    ).toBe(0)
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_range', range: 'A1', values: [['z']] }]))
    const r = await run(
      [
        'sheet',
        'apply',
        xlsx,
        '--ops',
        ops,
        '--dry-run',
        '--out',
        join(dir, 'new-dir/out.xlsx'),
        '--json',
      ],
      { env: env({}) },
    )
    expect(r.code).toBe(0)
    expect(existsSync(join(dir, 'new-dir'))).toBe(false)
  })

  it('accepts several roots separated by the platform delimiter', async () => {
    const a = tempDir()
    const b = tempDir()
    writeFileSync(join(b, 'a.csv'), 'x\n1\n')
    const e = env({
      GENOFFICE_ALLOWED_ROOTS: [a, b].join(process.platform === 'win32' ? ';' : ':'),
    })
    expect((await run(['info', join(b, 'a.csv'), '--json'], { env: e })).code).toBe(0)
  })
})

describe('audit log', () => {
  it('redacts secret flag values in both spellings', () => {
    expect(redactArgv(['convert', 'a.pdf', '--password', 'hunter2', '--to', 'docx'])).toEqual([
      'convert',
      'a.pdf',
      '--password',
      '***',
      '--to',
      'docx',
    ])
    expect(redactArgv(['info', 'a.pdf', '--password=hunter2'])).toEqual([
      'info',
      'a.pdf',
      '--password=***',
    ])
    expect(redactArgv(['info', 'a.pdf', '--password'])).toEqual(['info', 'a.pdf', '--password'])
  })

  it('defaults under ~/.genoffice and honours GENOFFICE_AUDIT_LOG', () => {
    expect(auditLogPath({})).toMatch(/[\\/]\.genoffice[\\/]cli-audit\.jsonl$/)
    expect(auditLogPath({ GENOFFICE_AUDIT_LOG: 'off' })).toBeNull()
    expect(auditLogPath({ GENOFFICE_AUDIT_LOG: '/x/y.jsonl' })).toBe('/x/y.jsonl')
  })

  it('appends one line per executed command, success or failure', async () => {
    const dir = tempDir()
    const log = join(dir, 'nested/audit.jsonl')
    const csv = join(dir, 'a.csv')
    writeFileSync(csv, 'x,y\n1,2\n')
    const e = { ...process.env, GENOFFICE_AUDIT_LOG: log }
    expect((await run(['info', csv, '--json'], { env: e })).code).toBe(0)
    expect((await run(['convert', csv, '--to', 'xlsx', '--json'], { env: e })).code).toBe(0)
    expect((await run(['info', join(dir, 'missing.csv'), '--json'], { env: e })).code).toBe(2)
    expect((await run(['help', 'info'], { env: e })).code).toBe(0)
    const lines = readFileSync(log, 'utf-8')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l))
    expect(lines).toHaveLength(3)
    expect(lines[0]).toMatchObject({
      command: 'info',
      status: 'ok',
      code: 0,
      argv: ['info', csv, '--json'],
    })
    expect(lines[1]).toMatchObject({
      command: 'convert',
      status: 'ok',
      output_path: join(dir, 'a.xlsx'),
    })
    expect(lines[2]).toMatchObject({ command: 'info', status: 'error', code: 2 })
    expect(typeof lines[0].ts).toBe('string')
    expect(typeof lines[0].ms).toBe('number')
  })
})
