import { describe, expect, it, vi } from 'vitest'
import { composeSkills } from '../src/skill'
import type { AgentSkill } from '../src/skill'
import type { AgentToolCall } from '../src/types'
import {
  createProxiedAppsSkill,
  formatProxiedAppResult,
  type ProxiedAppManifest,
  type ProxiedAppsDeps,
} from '../src/proxied-apps-skill'

/**
 * The desktop loop's renderer skill: routing by prefixed tool name, the
 * live-manifest degradation (no open window of an app → its tool is absent,
 * not failing), and the composed shape KbAskPanel builds (kb + attachments +
 * proxied apps in one skill).
 */

const WINDOWS = [
  { id: 'tab-2', kind: 'docs', title: 'report.docx', filePath: '/tmp/report.docx', active: false },
  { id: 'tab-5', kind: 'sheets', title: 'budget.xlsx', filePath: '/tmp/budget.xlsx', active: true },
  { id: 'tab-9', kind: 'pdf', title: 'scan.pdf', filePath: '/tmp/scan.pdf', active: false },
]

const MANIFEST: ProxiedAppManifest = {
  windows: WINDOWS,
  tools: [
    {
      name: 'docs:read_document',
      kind: 'docs',
      description: 'read a Word document',
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'sheets:read_sheet',
      kind: 'sheets',
      description: 'read a workbook',
      inputSchema: { type: 'object', properties: {} },
    },
  ],
}

function skillWith(
  manifest: ProxiedAppManifest | null,
  execute: ProxiedAppsDeps['execute'] = vi.fn(async () => ({
    target: { id: 'x', kind: 'docs', title: 't' },
    result: {},
  })),
): { skill: AgentSkill; execute: ProxiedAppsDeps['execute'] } {
  return { skill: createProxiedAppsSkill({ manifest: () => manifest, execute }), execute }
}

const call = (name: string, input: Record<string, unknown> = {}): AgentToolCall => ({
  id: 'c1',
  name,
  input,
})

describe('the proxied-apps skill surface', () => {
  it('offers exactly the manifest tools and lists the open windows as context', () => {
    const { skill } = skillWith(MANIFEST)
    expect(skill.tools.map((t) => t.name)).toEqual(['docs:read_document', 'sheets:read_sheet'])
    expect(skill.systemPrompt).toMatch(/read-only/i)
    const context = skill.buildContext?.() ?? ''
    expect(context).toContain('tab-2 | docs | report.docx | /tmp/report.docx')
    expect(context).toContain('(active)')
    // a pdf window is context even though no pdf tool exists (the app is a viewer)
    expect(context).toContain('tab-9 | pdf | scan.pdf')
  })

  it('degrades to nothing when no window of any app is open', () => {
    const { skill } = skillWith({ windows: [], tools: [] })
    expect(skill.tools).toEqual([])
    expect(skill.systemPrompt).toBe('')
    expect(skill.buildContext?.()).toBe('')
  })

  it('degrades to nothing before the first manifest refresh', () => {
    const { skill } = skillWith(null)
    expect(skill.tools).toEqual([])
    expect(skill.systemPrompt).toBe('')
  })

  it('routes a prefixed read to the main-process proxy and formats the result', async () => {
    const execute = vi.fn(
      async (tool: string) =>
        ({
          docs: {
            target: { id: 'tab-2', kind: 'docs', title: 'report.docx', path: '/tmp/report.docx' },
            result: { text: '0|heading|Report' },
          },
          sheets: {
            target: { id: 'tab-5', kind: 'sheets', title: 'budget.xlsx' },
            result: { sheets: [{ name: 'Sheet1' }] },
          },
        })[tool.split(':')[0]]!,
    )
    const { skill } = skillWith(MANIFEST, execute)
    const docs = await skill.executeTool(call('docs:read_document', { document: 'tab-2' }))
    expect(docs).toMatchObject({ output: '0|heading|Report', summary: 'read report.docx' })
    expect(execute).toHaveBeenCalledWith('docs:read_document', { document: 'tab-2' })
    const sheets = await skill.executeTool(call('sheets:read_sheet'))
    expect(sheets.output).toContain('"Sheet1"')
    expect(sheets.summary).toBe('read budget.xlsx')
  })

  it('a tool whose app closed mid-run degrades to a do-not-retry error', async () => {
    const execute = vi.fn()
    const { skill } = skillWith(
      { windows: WINDOWS, tools: MANIFEST.tools.filter((t) => t.kind === 'docs') },
      execute,
    )
    const result = await skill.executeTool(call('sheets:read_sheet'))
    expect(result.isError).toBe(true)
    expect(result.output).toMatch(/not available/)
    expect(execute).not.toHaveBeenCalled()
  })

  it('reports a bridge failure as the tool result instead of throwing', async () => {
    const { skill } = skillWith(
      MANIFEST,
      vi.fn(async () => {
        throw new Error('the document tab was closed')
      }),
    )
    const result = await skill.executeTool(call('docs:read_document'))
    expect(result).toMatchObject({ isError: true, output: 'the document tab was closed' })
  })

  it('rejects a non-string document argument', async () => {
    const execute = vi.fn()
    const { skill } = skillWith(MANIFEST, execute)
    const result = await skill.executeTool(call('docs:read_document', { document: 42 }))
    expect(result.isError).toBe(true)
    expect(result.output).toMatch(/document.*string/i)
    expect(execute).not.toHaveBeenCalled()
  })

  it('treats an empty document argument as "the active document"', async () => {
    const execute = vi.fn(async () => ({
      target: { id: 't', kind: 'docs', title: 'x' },
      result: '',
    }))
    const { skill } = skillWith(MANIFEST, execute)
    await skill.executeTool(call('docs:read_document', { document: '' }))
    expect(execute).toHaveBeenCalledWith('docs:read_document', {})
  })
})

describe('composed with the other home-ask skills', () => {
  /** a stand-in kb skill: one tool, its own executor */
  const kbSkill = (): AgentSkill => ({
    id: 'knowledge-base',
    systemPrompt: 'kb prompt',
    tools: [{ name: 'search_knowledge_base', description: 'search', inputSchema: {} }],
    executeTool: async () => ({ output: 'kb result', summary: 'kb' }),
  })

  it('merges the prefixed tools without collisions and routes each call to its owner', async () => {
    const appExecute = vi.fn(async () => ({
      target: { id: 't', kind: 'docs', title: 'x' },
      result: 'app result',
    }))
    const composed = composeSkills('home-ask', '', [
      kbSkill(),
      createProxiedAppsSkill({ manifest: () => MANIFEST, execute: appExecute }),
    ])
    expect(composed.tools.map((t) => t.name)).toEqual([
      'search_knowledge_base',
      'docs:read_document',
      'sheets:read_sheet',
    ])
    expect(composed.systemPrompt).toContain('kb prompt')
    expect(composed.systemPrompt).toMatch(/read-only/i)

    const viaKb = await composed.executeTool(call('search_knowledge_base', { query: 'lease' }))
    expect(viaKb.output).toBe('kb result')
    const viaApp = await composed.executeTool(call('docs:read_document', {}))
    expect(viaApp.output).toBe('app result')
    expect(appExecute).toHaveBeenCalledTimes(1)

    // a tool no sub-skill owns (an app whose tabs all closed) is a routing
    // miss at the compose layer: unknown tool, never executed
    const unknown = await composed.executeTool(call('slides:read_deck'))
    expect(unknown.isError).toBe(true)
    expect(unknown.output).toMatch(/Unknown tool/)
    expect(appExecute).toHaveBeenCalledTimes(1)
  })

  it('collapses to the kb-only surface when nothing is open', () => {
    const composed = composeSkills('home-ask', '', [
      kbSkill(),
      createProxiedAppsSkill({ manifest: () => null, execute: vi.fn() }),
    ])
    expect(composed.tools.map((t) => t.name)).toEqual(['search_knowledge_base'])
    expect(composed.systemPrompt).not.toMatch(/read-only/i)
  })
})

describe('formatProxiedAppResult', () => {
  it('passes text through, unwraps {text}, pretty-prints structures', () => {
    expect(formatProxiedAppResult('plain text')).toBe('plain text')
    expect(formatProxiedAppResult({ text: 'block list' })).toBe('block list')
    const json = formatProxiedAppResult({ sheets: [{ name: 'Sheet1', rows: 20 }] })
    expect(json).toContain('\n')
    expect(json).toContain('"rows": 20')
  })

  it('stringifies what JSON cannot', () => {
    const cyclic: Record<string, unknown> = {}
    cyclic.self = cyclic
    expect(formatProxiedAppResult(cyclic)).toBe(String(cyclic))
  })
})
