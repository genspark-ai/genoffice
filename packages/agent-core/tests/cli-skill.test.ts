import { describe, expect, it } from 'vitest'
import {
  createCliSkill,
  LOAD_GUIDE_TOOL,
  RUN_GENOFFICE_TOOL,
  type CliRunRequest,
  type CliRunResult,
  type CliSkillDeps,
} from '../src/cli-skill'
import { kbCiteHref } from '../src/knowledge-base-skill'

const CTX = { workDir: '/data/work/kb-ask', outDir: '/Users/me/Documents' }

function makeDeps(
  result: Partial<CliRunResult> = {},
): CliSkillDeps & { runs: CliRunRequest[]; ctx: typeof CTX | null } {
  const runs: CliRunRequest[] = []
  const deps = {
    runs,
    ctx: CTX as typeof CTX | null,
    context: () => deps.ctx,
    run: async (request: CliRunRequest) => {
      runs.push(request)
      return { ok: true, code: 0, stdout: '', stderr: '', ...result }
    },
  }
  return deps
}

const call = (name: string, input: Record<string, unknown>) => ({ id: 't1', name, input })

describe('createCliSkill', () => {
  it('exposes run_genoffice and load_guide and names the directories in context', () => {
    const skill = createCliSkill(makeDeps())
    expect(skill.tools.map((t) => t.name)).toEqual([RUN_GENOFFICE_TOOL, LOAD_GUIDE_TOOL])
    expect(skill.buildContext?.()).toContain(CTX.workDir)
    expect(skill.buildContext?.()).toContain(CTX.outDir)
  })

  it('stays silent in context until the host has resolved the directories', () => {
    const deps = makeDeps()
    deps.ctx = null
    expect(createCliSkill(deps).buildContext?.()).toBe('')
  })

  it('forwards args and staged files and reports the output path with a cite link', async () => {
    const out = '/Users/me/Documents/Weekly report (Q3).docx'
    const deps = makeDeps({
      json: { status: 'ok', command: 'create', summary: 'wrote 1 document', output_path: out },
    })
    const skill = createCliSkill(deps)
    const res = await skill.executeTool(
      call(RUN_GENOFFICE_TOOL, {
        args: ['create', '--type', 'docx', '--from', 'report.md', '--out', out],
        files: [{ name: 'report.md', content: '# Weekly' }],
      }),
    )
    expect(res.isError).toBeFalsy()
    expect(deps.runs[0]).toEqual({
      args: ['create', '--type', 'docx', '--from', 'report.md', '--out', out],
      files: [{ name: 'report.md', content: '# Weekly' }],
    })
    expect(res.output).toContain(`output_path: ${out}`)
    expect(res.output).toContain(`cite: ${kbCiteHref(out)}`)
    expect(skill.createdPaths()).toEqual([out])
  })

  it('marks a failed run as an error and keeps the CLI payload for the model', async () => {
    const skill = createCliSkill(
      makeDeps({
        ok: false,
        code: 2,
        json: { status: 'error', command: 'docs', code: 2, message: 'op_rejected: setFont' },
      }),
    )
    const res = await skill.executeTool(call(RUN_GENOFFICE_TOOL, { args: ['docs', 'apply'] }))
    expect(res.isError).toBe(true)
    expect(res.output).toContain('exit 2')
    expect(res.output).toContain('op_rejected')
    expect(skill.createdPaths()).toEqual([])
  })

  it('rejects malformed args and files without running anything', async () => {
    const deps = makeDeps()
    const skill = createCliSkill(deps)
    expect((await skill.executeTool(call(RUN_GENOFFICE_TOOL, { args: [] }))).isError).toBe(true)
    expect((await skill.executeTool(call(RUN_GENOFFICE_TOOL, { args: 'info' }))).isError).toBe(true)
    expect(
      (await skill.executeTool(call(RUN_GENOFFICE_TOOL, { args: ['info'], files: [{ name: 1 }] })))
        .isError,
    ).toBe(true)
    expect(deps.runs).toEqual([])
  })

  it('load_guide runs guide <domain> [topic] and returns the prose, not JSON', async () => {
    const deps = makeDeps({
      json: {
        status: 'ok',
        command: 'guide',
        summary: 'Word ops (genoffice docs apply)\n  setFont {…}',
      },
    })
    const skill = createCliSkill(deps)
    const res = await skill.executeTool(call(LOAD_GUIDE_TOOL, { domain: 'docs', topic: 'setFont' }))
    expect(deps.runs[0]).toEqual({ args: ['guide', 'docs', 'setFont'] })
    expect(res.output).toContain('setFont {…}')
    expect(res.output).not.toContain('"status"')
    expect((await skill.executeTool(call(LOAD_GUIDE_TOOL, { domain: 'pdf' }))).isError).toBe(true)
  })

  it('verifyResponse pushes back on a saved claim with no successful run', () => {
    const skill = createCliSkill(makeDeps())
    expect(skill.verifyResponse?.('I have saved the workbook to Documents.', [])).toMatch(
      /no run_genoffice call succeeded/,
    )
    expect(
      skill.verifyResponse?.('I have saved the workbook.', [
        { name: RUN_GENOFFICE_TOOL, ok: true },
      ]),
    ).toBeNull()
    expect(skill.verifyResponse?.('The file was created in 2019.', [])).toBeNull()
  })
})
