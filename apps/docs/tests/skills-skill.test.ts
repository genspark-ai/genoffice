// @vitest-environment jsdom
/**
 * The skills skill: three tools over the user's installed skills, with the
 * list cached into per-turn context once one call has fetched it. The bridge
 * (window.desktop) is a fake — the boundary guards live in the main process.
 */
import { describe, expect, it, beforeEach } from 'vitest'
import { createSkillsSkill } from '../src/renderer/ai/skills-skill'
import type { UserCliRun, UserSkillEntry } from '../src/shared/ipc'

const call = (name: string, input: Record<string, unknown> = {}) =>
  ({ id: 't1', name, input }) as never

const SKILLS: UserSkillEntry[] = [
  { name: 'batch-pdf', description: 'Convert a folder of docx to pdf', relevant: true },
  { name: 'git-flow', description: 'Document a release process', relevant: false },
]

function installBridge(over: Partial<Record<string, unknown>> = {}): void {
  ;(window as unknown as { desktop: unknown }).desktop = {
    skillsList: async () => SKILLS,
    skillText: async (name: string) => `body of ${name}`,
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

  it('lists installed skills and caches them for the per-turn context', async () => {
    const skill = createSkillsSkill()
    expect(skill.buildContext?.()).toBe('')
    const out = await skill.executeTool!(call('list_skills'))
    expect(out.isError).toBeUndefined()
    expect(out.output).toContain('batch-pdf — Convert a folder of docx to pdf')
    expect(out.output).toContain('git-flow — Document a release process')
    expect(out.output).toContain('(not about document formats)')
    expect(skill.buildContext?.()).toContain('batch-pdf')
  })

  it('reads one skill by name through the main-process bridge', async () => {
    const skill = createSkillsSkill()
    const out = await skill.executeTool!(call('read_skill', { name: 'batch-pdf' }))
    expect(out.output).toBe('body of batch-pdf')
    const bad = await skill.executeTool!(call('read_skill', {}))
    expect(bad.isError).toBe(true)
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
      skillsList: () => Promise.reject(new Error("No handler registered for 'ai:skills-list'")),
    })
    const skill = createSkillsSkill()
    const out = await skill.executeTool!(call('list_skills'))
    expect(out.isError).toBe(true)
    expect(out.output).toContain('No handler registered')
  })

  it('rejects a run_cli call whose args are not a string array', async () => {
    const skill = createSkillsSkill()
    const out = await skill.executeTool!(call('run_cli', { args: ['ok', 42] }))
    expect(out.isError).toBe(true)
    expect(out.output).toContain('non-empty array of strings')
  })

  it('says when no skills are installed', async () => {
    installBridge({ skillsList: async () => [] })
    const skill = createSkillsSkill()
    const out = await skill.executeTool!(call('list_skills'))
    expect(out.output).toContain('No skills are installed')
    expect(skill.buildContext?.()).toBe('')
  })
})
