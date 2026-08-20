import { describe, it, expect, vi } from 'vitest'
import { createWorkspaceSkill, type WorkspaceSearchFn } from '../src/workspace-skill'
import type { WorkspaceSearchResult } from '../src/embed'

const call = (name: string, input: Record<string, unknown>) => ({ id: '1', name, input })

function makeSkill(result: WorkspaceSearchResult) {
  const search = vi.fn<WorkspaceSearchFn>(async () => result)
  return { skill: createWorkspaceSkill(search), search }
}

describe('createWorkspaceSkill', () => {
  it('exposes the workspace_search tool with a query-required schema', () => {
    const { skill } = makeSkill({ ok: true })
    expect(skill.id).toBe('workspace')
    expect(skill.tools.map((t) => t.name)).toEqual(['workspace_search'])
    const schema = skill.tools[0]!.inputSchema as {
      required?: string[]
      properties?: Record<string, unknown>
    }
    expect(schema.required).toEqual(['query'])
    expect(schema.properties).toHaveProperty('query')
    expect(skill.buildContext?.()).toBe('')
    expect(skill.systemPrompt).toContain('workspace_search')
  })

  it('rejects unknown tools', async () => {
    const { skill } = makeSkill({ ok: true })
    const exec = await skill.executeTool(call('other_tool', {}))
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('unknown tool')
  })

  it('requires a non-empty query', async () => {
    const { skill, search } = makeSkill({ ok: true })
    const exec = await skill.executeTool(call('workspace_search', { query: '   ' }))
    expect(exec.isError).toBe(true)
    expect(search).not.toHaveBeenCalled()
  })

  it('passes the bridge the query and a clamped k (default 5, max 10, min 1)', async () => {
    const { skill, search } = makeSkill({ ok: true, query: 'x', results: [] })
    await skill.executeTool(call('workspace_search', { query: 'budget' }))
    expect(search).toHaveBeenCalledWith('budget', 5)
    await skill.executeTool(call('workspace_search', { query: 'budget', k: 20 }))
    expect(search).toHaveBeenCalledWith('budget', 10)
    await skill.executeTool(call('workspace_search', { query: 'budget', k: -3 }))
    expect(search).toHaveBeenCalledWith('budget', 1)
  })

  it('surfaces a search failure as a tool error, never as an empty result', async () => {
    const { skill } = makeSkill({ ok: false, error: 'No embedding model installed' })
    const exec = await skill.executeTool(call('workspace_search', { query: 'budget' }))
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('No embedding model installed')
  })

  it('says plainly when the index has no matches', async () => {
    const { skill } = makeSkill({ ok: true, query: 'budget', results: [] })
    const exec = await skill.executeTool(call('workspace_search', { query: 'budget' }))
    expect(exec.isError).toBeFalsy()
    expect(exec.output).toContain('No matching passages')
    expect(exec.summary).toBe('workspace_search: 0 matches')
  })

  it('formats matches with file names and snippets', async () => {
    const { skill } = makeSkill({
      ok: true,
      query: 'budget',
      results: [
        { file: '/docs/budget.md', snippet: 'Marketing gets $640K.', score: 0.7 },
        { file: '/docs/notes.txt', snippet: 'Cloud costs $60K/mo.', score: 0.5 },
      ],
    })
    const exec = await skill.executeTool(call('workspace_search', { query: 'budget' }))
    expect(exec.output).toContain('Top 2 matches')
    expect(exec.output).toContain('[1] /docs/budget.md')
    expect(exec.output).toContain('Marketing gets $640K.')
    expect(exec.output).toContain('[2] /docs/notes.txt')
    expect(exec.summary).toBe('workspace_search: 2 matches')
  })
})
