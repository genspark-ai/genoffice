/**
 * The AI panel's door to installed skills and to the bundled CLI: the argv
 * validation is the security boundary for run_cli, and the skill-text channel
 * must resolve names against installed skills only — never a renderer-supplied
 * path.
 */
import { describe, expect, it, vi, beforeEach } from 'vitest'

const handles = new Map<string, (...args: never[]) => unknown>()
vi.mock('electron', () => ({
  app: { getPath: () => '/app/userData', isPackaged: true, getAppPath: () => '/app' },
  ipcMain: {
    handle: (channel: string, handler: (...args: never[]) => unknown) =>
      handles.set(channel, handler),
  },
  clipboard: {},
  dialog: {},
}))

const usable = vi.fn()
const readBody = vi.fn()
const runnerRun = vi.fn()
vi.mock('../src/main/imported-skills', () => ({
  usableSkills: (...a: unknown[]) => usable(...a),
  readSkillBody: (...a: unknown[]) => readBody(...a),
}))
vi.mock('../src/main/mcp/cli-runner', () => ({
  createCliRunner: () => ({ run: (...a: unknown[]) => runnerRun(...a) }),
}))

import { registerUserSkillsIpc, validateCliArgs } from '../src/main/user-cli'

const invoke = (channel: string, ...args: unknown[]) => handles.get(channel)!(null, ...args)

beforeEach(() => {
  handles.clear()
  usable.mockReset().mockResolvedValue([
    {
      name: 'batch-pdf',
      description: 'd',
      path: '/app/userData/skills/batch-pdf/SKILL.md',
      relevance: { relevant: true, matched: ['pdf'] },
    },
  ])
  readBody.mockReset().mockReturnValue('the body')
  runnerRun
    .mockReset()
    .mockResolvedValue({ ok: true, code: 0, json: undefined, stdout: 'o', stderr: '' })
  registerUserSkillsIpc({ cliRunner: () => ({ run: runnerRun }) as never })
})

describe('validateCliArgs', () => {
  it('accepts a plain command line and passes it through untouched', () => {
    expect(validateCliArgs(['convert', 'in.docx', '--to', 'pdf'])).toEqual([
      'convert',
      'in.docx',
      '--to',
      'pdf',
    ])
  })

  it('rejects the shapes that are not a command line', () => {
    expect(() => validateCliArgs('convert')).toThrow()
    expect(() => validateCliArgs([])).toThrow()
    expect(() => validateCliArgs(['convert', 42])).toThrow()
    expect(() => validateCliArgs(['', 'x'])).toThrow()
    // a flag first is a probe at global flags, not a command
    expect(() => validateCliArgs(['--version'])).toThrow('must be a command')
    expect(() => validateCliArgs(Array.from({ length: 9 }, () => 'x'))).toThrow('at most 8')
    expect(() => validateCliArgs(['convert', 'x'.repeat(201)])).toThrow('200 characters')
  })
})

describe('the ai: channels', () => {
  it('lists installed skills without paths', async () => {
    const out = (await invoke('ai:skills-list')) as { name: string; relevant: boolean }[]
    expect(out).toEqual([{ name: 'batch-pdf', description: 'd', relevant: true }])
  })

  it('resolves skill text by name against the installed list, never a path', async () => {
    const body = (await invoke('ai:skill-text', 'batch-pdf')) as string
    expect(body).toBe('the body')
    expect(readBody).toHaveBeenCalledWith(expect.stringContaining('batch-pdf'))
    await expect(invoke('ai:skill-text', 'nope')).rejects.toThrow('no installed skill named')
    await expect(invoke('ai:skill-text', '/etc/passwd')).rejects.toThrow('no installed skill named')
  })

  it('runs the validated argv on the bundled cli runner', async () => {
    const out = (await invoke('ai:run-cli', ['convert', 'a.docx', '--to', 'pdf'])) as {
      ok: boolean
      code: number
    }
    expect(out).toMatchObject({ ok: true, code: 0 })
    expect(runnerRun).toHaveBeenCalledWith(['convert', 'a.docx', '--to', 'pdf'], {
      timeoutMs: 120_000,
    })
    await expect(invoke('ai:run-cli', ['--help'])).rejects.toThrow('must be a command')
  })
})
