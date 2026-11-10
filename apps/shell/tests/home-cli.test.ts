import { mkdir, mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { delimiter, dirname, join, resolve } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  allowedRootsFor,
  validateSources,
  HomeCliError,
  homeCliContext,
  inPlaceTarget,
  runHomeCli,
  stagedPath,
  validateHomeCliArgs,
  validateSessionId,
  validateStagedFiles,
  type HomeCliDeps,
} from '../src/main/home-cli'
import type { CliRunOptions, CliRunOutcome } from '../src/main/mcp/cli-runner'

describe('validateHomeCliArgs', () => {
  it('accepts a document command line', () => {
    expect(
      validateHomeCliArgs(['create', '--type', 'docx', '--from', 'r.md', '--out', 'x.docx']),
    ).toHaveLength(7)
  })

  it('refuses commands that leave the document domain', () => {
    for (const cmd of ['open', 'mcp', 'skill', 'install-cli', 'selection', 'rm']) {
      expect(() => validateHomeCliArgs([cmd])).toThrow(HomeCliError)
    }
  })

  it('refuses --force, drops a stray --json and lets --help through', () => {
    expect(() => validateHomeCliArgs(['convert', 'a.docx', '--to', 'pdf', '--force'])).toThrow(
      /--force/,
    )
    expect(validateHomeCliArgs(['info', 'a.docx', '--json'])).toEqual(['info', 'a.docx'])
    expect(validateHomeCliArgs(['docs', '--help'])).toEqual(['docs', '--help'])
    expect(validateHomeCliArgs(['--help'])).toEqual(['--help'])
  })

  it('names the file an apply / replace without --out rewrites', () => {
    expect(inPlaceTarget(['docs', 'apply', 'a.docx', '--ops', 'o.json'])).toBe('a.docx')
    expect(inPlaceTarget(['slides', 'replace', 'd.pptx', '--slide', '1'])).toBe('d.pptx')
    expect(inPlaceTarget(['docs', 'apply', 'a.docx', '--ops', 'o.json', '--out', 'b.docx'])).toBe(
      null,
    )
    expect(inPlaceTarget(['docs', 'apply', '--help'])).toBe(null)
    expect(inPlaceTarget(['docs', 'read', 'a.docx'])).toBe(null)
    // the file may come after the flags; a flag's value is not the file
    expect(inPlaceTarget(['docs', 'apply', '--ops', 'o.json', 'a.docx'])).toBe('a.docx')
    expect(inPlaceTarget(['docs', 'apply', '--track', '--author', 'AI', 'a.docx'])).toBe('a.docx')
    expect(inPlaceTarget(['sheet', 'apply', '--dry-run', 'b.xlsx', '--ops', 'o'])).toBe('b.xlsx')
    expect(() => inPlaceTarget(['docs', 'apply', '--ops', 'o.json'])).toThrow(HomeCliError)
    expect(inPlaceTarget(['docs', 'apply', 'a.docx', '--ops', 'o.json', '--out=b.docx'])).toBe(null)
    expect(inPlaceTarget(['docs', 'apply', '--author', '--track', 'a.docx', '--ops', 'o'])).toBe(
      'a.docx',
    )
    expect(inPlaceTarget(['docs', 'apply', '--ops=o.json', 'a.docx'])).toBe('a.docx')
    expect(() => inPlaceTarget(['docs', 'apply', 'a.docx', '--track=yes'])).toThrow(HomeCliError)
  })

  it('bounds the argv', () => {
    expect(() => validateHomeCliArgs([])).toThrow(HomeCliError)
    expect(() => validateHomeCliArgs(['info', ''])).toThrow(HomeCliError)
    expect(() => validateHomeCliArgs(['info', 'x'.repeat(1001)])).toThrow(/stage it as a file/)
    expect(() => validateHomeCliArgs(Array.from({ length: 33 }, () => 'a'))).toThrow(/at most/)
  })
})

describe('staging', () => {
  it('keeps staged names inside the work directory', () => {
    expect(stagedPath('/w', 'deck/pages/01.json')).toBe(join('/w', 'deck', 'pages', '01.json'))
    expect(stagedPath('/w', './ops.json')).toBe(join('/w', 'ops.json'))
    expect(() => stagedPath('/w', '../ops.json')).toThrow(HomeCliError)
    expect(() => stagedPath('/w', 'a/../../ops.json')).toThrow(HomeCliError)
    expect(() => stagedPath('/w', '/etc/ops.json')).toThrow(HomeCliError)
    expect(() => stagedPath('/w', 'C:\\x.json')).toThrow(HomeCliError)
  })

  it('validates the staged file list', () => {
    expect(validateStagedFiles(undefined)).toEqual([])
    expect(validateStagedFiles([{ name: 'a.md', content: '#' }])).toEqual([
      { name: 'a.md', content: '#' },
    ])
    expect(() => validateStagedFiles([{ name: 'a.md' }])).toThrow(HomeCliError)
    expect(() => validateStagedFiles('x')).toThrow(HomeCliError)
  })

  it('accepts thread ids and rejects path-like ones', () => {
    expect(validateSessionId('kb-ask')).toBe('kb-ask')
    expect(validateSessionId('kb-ask-1760000000000')).toBeTruthy()
    expect(() => validateSessionId('../x')).toThrow(HomeCliError)
    expect(() => validateSessionId('')).toThrow(HomeCliError)
    expect(() => validateSessionId(42)).toThrow(HomeCliError)
  })
})

describe('runHomeCli', () => {
  let root: string
  const runs: Array<{ args: string[]; options?: CliRunOptions }> = []
  let outcome: CliRunOutcome
  const recent: string[] = []
  let deps: HomeCliDeps

  beforeEach(async () => {
    root = await mkdtemp(join(tmpdir(), 'home-cli-'))
    runs.length = 0
    recent.length = 0
    outcome = { ok: true, code: 0, stdout: '{"status":"ok"}', stderr: '' }
    deps = {
      runner: () => ({
        run: async (args, options) => {
          runs.push({ args, options })
          return outcome
        },
      }),
      workRoot: () => join(root, 'work'),
      outDir: () => join(root, 'Documents'),
      sourcePaths: () => [join(root, 'src', 'a.xlsx'), join(root, 'src', 'b.xlsx')],
      recordRecent: (path) => void recent.push(path),
    }
  })
  afterEach(() => rm(root, { recursive: true, force: true }))

  it('stages files into the thread work directory and confines the CLI to the roots', async () => {
    const out = join(root, 'Documents', 'report.docx')
    outcome.json = { status: 'ok', command: 'create', summary: 'ok', output_path: out }
    const res = await runHomeCli(deps, 'kb-ask', {
      args: ['create', '--type', 'docx', '--from', 'report.md', '--out', out],
      files: [{ name: 'report.md', content: '# Hi' }],
    })
    const workDir = join(root, 'work', 'kb-ask')
    expect(await readFile(join(workDir, 'report.md'), 'utf8')).toBe('# Hi')
    expect(runs[0]!.options?.cwd).toBe(workDir)
    const roots = runs[0]!.options?.env?.GENOFFICE_ALLOWED_ROOTS?.split(delimiter) ?? []
    expect(roots).toEqual([
      resolve(workDir),
      resolve(join(root, 'Documents')),
      resolve(join(root, 'src')),
    ])
    expect(runs[0]!.args.at(-1)).toBe(out)
    expect(res.ok).toBe(true)
    expect(recent).toEqual([out])
  })

  it('does not record files written outside the output directory as recent', async () => {
    outcome.json = {
      status: 'ok',
      command: 'render',
      summary: 'ok',
      output_path: join(root, 'work', 'kb-ask', 'shots'),
    }
    await runHomeCli(deps, 'kb-ask', { args: ['render', 'x.docx', '--out', 'shots'] })
    expect(recent).toEqual([])
  })

  it('edits a file inside the output directory in place after backing it up', async () => {
    await mkdir(join(root, 'Documents'), { recursive: true })
    const target = join(root, 'Documents', 'book.xlsx')
    await writeFile(target, 'original')
    await runHomeCli(deps, 'kb-ask', {
      args: ['sheet', 'apply', target, '--ops', 'ops.json'],
      files: [{ name: 'ops.json', content: '[]' }],
    })
    expect(runs).toHaveLength(1)
    const backups = await readdir(join(root, 'work', 'kb-ask', '.backup'))
    expect(backups).toHaveLength(1)
    expect(backups[0]).toMatch(/-book\.xlsx$/)
    expect(await readFile(join(root, 'work', 'kb-ask', '.backup', backups[0]!), 'utf8')).toBe(
      'original',
    )
  })

  it('refuses an in-place edit of a source outside the output directory', async () => {
    await expect(
      runHomeCli(deps, 'kb-ask', {
        args: ['docs', 'apply', join(root, 'src', 'a.docx'), '--ops', 'ops.json'],
        files: [{ name: 'ops.json', content: '[]' }],
      }),
    ).rejects.toThrow(/only allowed inside the output directory/)
    await expect(
      runHomeCli(deps, 'kb-ask', {
        args: ['docs', 'apply', '--ops', 'ops.json', join(root, 'src', 'a.docx')],
        files: [{ name: 'ops.json', content: '[]' }],
      }),
    ).rejects.toThrow(/only allowed inside the output directory/)
    await expect(
      runHomeCli(deps, 'kb-ask', {
        args: ['docs', 'apply', join(root, 'Documents', 'missing.docx'), '--ops', 'o'],
      }),
    ).rejects.toThrow(/does not exist/)
    expect(runs).toEqual([])
  })

  it('rejects before spawning when the argv or a staged name is bad', async () => {
    await expect(runHomeCli(deps, 'kb-ask', { args: ['open', 'x.docx'] })).rejects.toThrow(
      HomeCliError,
    )
    await expect(
      runHomeCli(deps, 'kb-ask', { args: ['info', 'x'], files: [{ name: '../x', content: '' }] }),
    ).rejects.toThrow(HomeCliError)
    expect(runs).toEqual([])
  })

  it('names the per-thread directories', () => {
    expect(homeCliContext(deps, 'kb-ask-7')).toEqual({
      workDir: join(root, 'work', 'kb-ask-7'),
      outDir: join(root, 'Documents'),
    })
    expect(() => homeCliContext(deps, '')).toThrow(HomeCliError)
  })

  it('attached files widen the roots to their folders when they exist', async () => {
    const att = join(root, 'Downloads', 'brief.pdf')
    await mkdir(dirname(att), { recursive: true })
    await writeFile(att, 'pdf')
    await runHomeCli(deps, 'kb-ask', {
      args: ['convert', att, '--out', join(root, 'Documents', 'brief.docx')],
      sources: [att, join(root, 'gone', 'x.docx'), 'relative.docx', 42 as unknown as string],
    })
    const roots = runs[0]!.options?.env?.GENOFFICE_ALLOWED_ROOTS?.split(delimiter) ?? []
    expect(roots).toContain(resolve(join(root, 'Downloads')))
    expect(roots).not.toContain(resolve(join(root, 'gone')))
    expect(validateSources([att, join(root, 'Downloads')])).toEqual([att])
  })

  it('allowedRootsFor de-duplicates folders', () => {
    const ctx = { workDir: '/w', outDir: '/o' }
    expect(allowedRootsFor(ctx, ['/o/a.docx', '/o/b.docx', '/s/c.pdf']).split(delimiter)).toEqual([
      resolve('/w'),
      resolve('/o'),
      resolve('/s'),
    ])
  })
})
