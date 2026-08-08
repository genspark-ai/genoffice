import { describe, expect, it, vi } from 'vitest'

import { FORGET_TOOL, REMEMBER_TOOL } from '../src/memory'
import { createWebSkill, type WebSkillBridge } from '../src/web-skill'

function skill(overrides: Partial<WebSkillBridge> = {}) {
  const bridge: WebSkillBridge = {
    browsePage: vi.fn(),
    extractPages: vi.fn(),
    loadSkill: vi.fn(),
    remember: vi.fn().mockResolvedValue(true),
    forget: vi.fn().mockResolvedValue(true),
    ...overrides,
  }
  return {
    bridge,
    skill: createWebSkill({
      bridge,
      surface: 'docx',
      hasUserSkills: () => false,
      instructionsPrompt: () => '',
    }),
  }
}

const call = (name: string, input: Record<string, unknown>) => ({ id: 't1', name, input })

describe('memory tools', () => {
  it('offers remember and forget even when the user has no skills', () => {
    const names = skill().skill.tools.map((t) => t.name)
    expect(names).toContain(REMEMBER_TOOL)
    expect(names).toContain(FORGET_TOOL)
  })

  it('tells the model what not to record, on the tool itself', () => {
    const remember = skill().skill.tools.find((t) => t.name === REMEMBER_TOOL)!
    // the guidance has to travel with the tool: a model that never reads the
    // system prompt section still sees this
    expect(remember.description.toLowerCase()).toContain('never record document contents')
  })

  it('passes the text through to the bridge', async () => {
    const { bridge, skill: s } = skill()
    const result = await s.executeTool!(call(REMEMBER_TOOL, { text: '  prefers tables  ' }))
    expect(bridge.remember).toHaveBeenCalledWith('prefers tables')
    expect(result.output).toBe('Remembered.')
    // recording a preference is not a document edit
    expect(result.mutated).toBe(false)
  })

  it('forgets by wording', async () => {
    const { bridge, skill: s } = skill()
    const result = await s.executeTool!(call(FORGET_TOOL, { text: 'prefers tables' }))
    expect(bridge.forget).toHaveBeenCalledWith('prefers tables')
    expect(result.output).toBe('Forgotten.')
  })

  it('rejects empty text as an error the model should not retry blindly', async () => {
    const { bridge, skill: s } = skill()
    const result = await s.executeTool!(call(REMEMBER_TOOL, { text: '   ' }))
    expect(result.isError).toBe(true)
    expect(bridge.remember).not.toHaveBeenCalled()
  })

  it('reports a declined store as plain output, not an error', async () => {
    // the store refusing (too long, nothing matched) is a normal answer — an
    // isError result would push the loop into a retry that cannot succeed
    const { skill: s } = skill({ remember: vi.fn().mockResolvedValue(false) })
    const result = await s.executeTool!(call(REMEMBER_TOOL, { text: 'x' }))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('Not stored')
  })

  it('reports a miss on forget without failing the turn', async () => {
    const { skill: s } = skill({ forget: vi.fn().mockResolvedValue(false) })
    const result = await s.executeTool!(call(FORGET_TOOL, { text: 'never said this' }))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('Nothing recorded matches')
  })
})
