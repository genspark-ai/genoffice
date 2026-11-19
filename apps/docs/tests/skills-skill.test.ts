// @vitest-environment jsdom
/**
 * The skills skill: only run_cli. Skills are never listed or read by the model;
 * the user attaches them from the composer. The bridge (window.desktop) is a
 * fake — the boundary guards live in the main process.
 */
import { describe, expect, it, beforeEach } from 'vitest'
import { createSkillsSkill } from '../src/renderer/ai/skills-skill'
import type { UserCliRun } from '../src/shared/ipc'

const call = (name: string, input: Record<string, unknown> = {}) =>
  ({ id: 't1', name, input }) as never

function installBridge(over: Partial<Record<string, unknown>> = {}): void {
  ;(window as unknown as { desktop: unknown }).desktop = {
    runCli: async (args: string[]): Promise<UserCliRun> => ({
      ok: true,
      code: 0,
      json: { status: 'ok', command: args.join(' '), summary: 'done' },
      stdout: '{"status":"ok"}',
      stderr: '',
    }),
    ...over,
  }
}

describe('createSkillsSkill', () => {
  beforeEach(() => installBridge())

  it('offers run_cli only: the model cannot list or read skills on its own', () => {
    const skill = createSkillsSkill()
    expect(skill.tools.map((t) => t.name)).toEqual(['run_cli'])
    expect(skill.buildContext).toBeUndefined()
    expect(skill.systemPrompt).toContain('Never assume a skill')
  })

  it('reports the cli outcome with its exit code and the json summary', async () => {
    const skill = createSkillsSkill()
    const out = await skill.executeTool!(
      call('run_cli', { args: ['convert', 'a.docx', '--to', 'pdf'] }),
    )
    expect(out.isError).toBeFalsy()
    expect(out.output).toContain('exit 0')
    expect(out.output).toContain('summary: done')
    expect(out.output).toContain('{"status":"ok"}')
    expect(out.mutated).toBe(false)
  })

  it('counts a run that wrote a file as a mutation', async () => {
    installBridge({
      runCli: async (): Promise<UserCliRun> => ({
        ok: true,
        code: 0,
        json: { status: 'ok', command: 'convert', summary: 'done', output_path: '/out/a.pdf' },
        stdout: '',
        stderr: '',
      }),
    })
    const skill = createSkillsSkill()
    const out = await skill.executeTool!(call('run_cli', { args: ['convert', 'a.docx'] }))
    expect(out.mutated).toBe(true)
  })

  it('marks a failing command as an error result', async () => {
    installBridge({
      runCli: async (): Promise<UserCliRun> => ({
        ok: false,
        code: 3,
        json: { status: 'error', command: 'convert x', code: 3, message: 'conversion failed' },
        stdout: '',
        stderr: 'boom',
      }),
    })
    const skill = createSkillsSkill()
    const out = await skill.executeTool!(call('run_cli', { args: ['convert'] }))
    expect(out.isError).toBe(true)
    expect(out.output).toContain('exit 3')
    expect(out.output).toContain('[stderr]\nboom')
  })

  it('turns a missing main-process handler into a tool error, not a throw', async () => {
    installBridge({
      runCli: () => Promise.reject(new Error("No handler registered for 'ai:run-cli'")),
    })
    const skill = createSkillsSkill()
    const out = await skill.executeTool!(call('run_cli', { args: ['info', 'a.docx'] }))
    expect(out.isError).toBe(true)
    expect(out.output).toContain('No handler registered')
  })

  it('rejects a run_cli call whose args are not a string array', async () => {
    const skill = createSkillsSkill()
    const out = await skill.executeTool!(call('run_cli', { args: ['ok', 42] }))
    expect(out.isError).toBe(true)
    expect(out.output).toContain('non-empty array of strings')
  })
})
