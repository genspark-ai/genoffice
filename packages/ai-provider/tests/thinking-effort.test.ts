import { describe, expect, it } from 'vitest'

import {
  anthropicThinkingFor,
  geminiThinkingLevelFor,
  openAiEffortFor,
  resolveThinkingEffort,
  thinkingShapeFor,
  toggleThinkingBodyFor,
} from '../src/thinking-effort'
import { sanitizeAiSettings } from '../src/ai-settings-guard'

describe('thinkingShapeFor', () => {
  it('openai: the 6.x ladder with none gated on the tiers that refuse it', () => {
    expect(thinkingShapeFor('gpt-6.1-terra').levels).toEqual([
      'none',
      'low',
      'medium',
      'high',
      'xhigh',
      'max',
    ])
    expect(thinkingShapeFor('gpt-6-astra').levels).toEqual([
      'low',
      'medium',
      'high',
      'xhigh',
      'max',
    ])
    expect(thinkingShapeFor('gpt-6-sol').levels).not.toContain('none')
    expect(thinkingShapeFor('gpt-6-luna').levels).toContain('none')
    expect(thinkingShapeFor('gpt-5.6-terra').levels).toEqual(['none', 'low', 'medium', 'high'])
    expect(thinkingShapeFor('gpt-5.5').kind).toBe('openai-effort')
  })

  it('anthropic: adaptive effort, 5.x extends past high', () => {
    expect(thinkingShapeFor('claude-opus-5-5')).toEqual({
      kind: 'anthropic-adaptive',
      levels: ['none', 'low', 'medium', 'high', 'xhigh', 'max'],
    })
    expect(thinkingShapeFor('claude-sonnet-4-6').levels).toEqual(['none', 'low', 'medium', 'high'])
  })

  it('gemini 3/4: low and high only', () => {
    expect(thinkingShapeFor('gemini-4-argon')).toEqual({
      kind: 'gemini-level',
      levels: ['low', 'high'],
    })
    expect(thinkingShapeFor('gemini-3.8-flash').levels).toEqual(['low', 'high'])
  })

  it('deepseek and kimi take reasoning_effort with their own ladders', () => {
    expect(thinkingShapeFor('deepseek-v4-pro')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'high', 'max'],
    })
    expect(thinkingShapeFor('deepseek-flash').levels).toEqual(['low', 'high', 'max'])
    expect(thinkingShapeFor('kimi-k3.1-pro')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'high', 'max'],
    })
    expect(thinkingShapeFor('kimi-k3').levels).toEqual(['low', 'high', 'max'])
  })

  it('glm: the full ladder through effort, none only where thinking can stop', () => {
    expect(thinkingShapeFor('glm-5.2').levels).toEqual([
      'none',
      'low',
      'medium',
      'high',
      'xhigh',
      'max',
    ])
    expect(thinkingShapeFor('glm-5.3').levels).toEqual(['low', 'medium', 'high', 'xhigh', 'max'])
    expect(thinkingShapeFor('glm-5.3-flash').levels).not.toContain('none')
  })

  it('qwen: enable_thinking plus a budget per level', () => {
    expect(thinkingShapeFor('qwen3.8-max')).toEqual({
      kind: 'qwen-budget',
      levels: ['none', 'low', 'medium', 'high'],
    })
    expect(
      toggleThinkingBodyFor({ model: 'qwen3.8-max', thinking: { 'qwen3.8-max': 'low' } }),
    ).toEqual({ enable_thinking: true, thinking_budget: 4096 })
  })

  it('ling is a thinking switch; grok speaks effort; longcat stays knobless', () => {
    expect(thinkingShapeFor('ling-3.0-flash')).toEqual({
      kind: 'toggle',
      levels: ['none', 'medium'],
    })
    expect(
      toggleThinkingBodyFor({ model: 'ling-3.0-flash', thinking: { 'ling-3.0-flash': 'none' } }),
    ).toEqual({ thinking: { type: 'disabled' } })
    expect(thinkingShapeFor('grok-4.7')).toEqual({
      kind: 'openai-effort',
      levels: ['none', 'low', 'medium', 'high'],
    })
    expect(thinkingShapeFor('LongCat-2.0').levels).toEqual([])
  })

  it('minimax carries the five effort rungs without none; mimo is a switch', () => {
    expect(thinkingShapeFor('MiniMax-M3.1-Flash').levels).toEqual([
      'low',
      'medium',
      'high',
      'xhigh',
      'max',
    ])
    expect(thinkingShapeFor('MiniMax-M3').levels).not.toContain('none')
    expect(
      openAiEffortFor({ model: 'MiniMax-M3.1-Flash', thinking: { MiniMax: 'xhigh' } }),
    ).toBeUndefined()
    expect(
      openAiEffortFor({ model: 'MiniMax-M3.1-Flash', thinking: { 'MiniMax-M3.1-Flash': 'xhigh' } }),
    ).toBe('xhigh')
    expect(thinkingShapeFor('mimo-v2.6-pro')).toEqual({
      kind: 'toggle',
      levels: ['none', 'medium'],
    })
  })

  it('doubao, hunyuan, spark and mistral-medium carry their documented surfaces', () => {
    expect(thinkingShapeFor('doubao-seed-2-1-pro-260628').levels).toEqual([
      'none',
      'low',
      'medium',
      'high',
      'xhigh',
      'max',
    ])
    expect(thinkingShapeFor('hy4-preview')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'high'],
    })
    expect(thinkingShapeFor('hy3').levels).toEqual(['low', 'high'])
    expect(thinkingShapeFor('spark-x2.5')).toEqual({ kind: 'toggle', levels: ['none', 'medium'] })
    expect(thinkingShapeFor('mistral-medium-latest').levels).toEqual(['low', 'medium', 'high'])
    expect(thinkingShapeFor('mistral-large-latest').levels).toEqual([])
  })
})

describe('resolveThinkingEffort', () => {
  it('resolves the stored pick only when the family lists it', () => {
    expect(
      resolveThinkingEffort({ model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'max' } }),
    ).toBe('max')
    expect(
      resolveThinkingEffort({ model: 'gpt-6-astra', thinking: { 'gpt-6-astra': 'none' } }),
    ).toBeUndefined()
    expect(
      resolveThinkingEffort({ model: 'gemini-4-argon', thinking: { 'gemini-4-argon': 'high' } }),
    ).toBe('high')
    expect(
      resolveThinkingEffort({ model: 'gemini-4-argon', thinking: { 'gemini-4-argon': 'max' } }),
    ).toBeUndefined()
    expect(
      resolveThinkingEffort({ model: 'deepseek-v4-pro', thinking: { 'deepseek-v4-pro': 'none' } }),
    ).toBeUndefined()
    expect(resolveThinkingEffort({ model: 'kimi-k3', thinking: { 'kimi-k3': 'high' } })).toBe(
      'high',
    )
    expect(resolveThinkingEffort({ model: 'glm-5.2', thinking: { 'glm-5.2': 'none' } })).toBe(
      'none',
    )
    expect(
      resolveThinkingEffort({ model: 'glm-5.3', thinking: { 'glm-5.3': 'none' } }),
    ).toBeUndefined()
    expect(resolveThinkingEffort({ model: 'gpt-6.1-terra', thinking: {} })).toBeUndefined()
  })
})

describe('per-protocol translators', () => {
  it('openai effort rides both routes; nothing for other families', () => {
    expect(
      openAiEffortFor({ model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'xhigh' } }),
    ).toBe('xhigh')
    expect(openAiEffortFor({ model: 'grok-4.7', thinking: { 'grok-4.7': 'low' } })).toBe('low')
    expect(
      openAiEffortFor({ model: 'claude-opus-5-5', thinking: { 'claude-opus-5-5': 'high' } }),
    ).toBeUndefined()
  })

  it('anthropic: adaptive record, none means no thinking block', () => {
    expect(
      anthropicThinkingFor({ model: 'claude-opus-5-5', thinking: { 'claude-opus-5-5': 'max' } }),
    ).toEqual({
      type: 'adaptive',
      effort: 'max',
    })
    expect(
      anthropicThinkingFor({
        model: 'claude-sonnet-4-6',
        thinking: { 'claude-sonnet-4-6': 'none' },
      }),
    ).toBeUndefined()
    expect(
      anthropicThinkingFor({
        model: 'claude-sonnet-4-6',
        thinking: { 'claude-sonnet-4-6': 'xhigh' },
      }),
    ).toBeUndefined()
  })

  it('gemini: thinkingLevel passthrough for the listed levels', () => {
    expect(
      geminiThinkingLevelFor({
        model: 'gemini-3.8-flash',
        thinking: { 'gemini-3.8-flash': 'low' },
      }),
    ).toBe('low')
    expect(
      geminiThinkingLevelFor({
        model: 'gemini-3.8-flash',
        thinking: { 'gemini-3.8-flash': 'none' },
      }),
    ).toBeUndefined()
  })

  it('toggles: qwen flips enable_thinking, doubao rides effort now', () => {
    expect(
      toggleThinkingBodyFor({
        model: 'doubao-seed-2-1-pro-260628',
        thinking: { 'doubao-seed-2-1-pro-260628': 'medium' },
      }),
    ).toBeUndefined()
    expect(
      toggleThinkingBodyFor({ model: 'qwen3.8-max', thinking: { 'qwen3.8-max': 'none' } }),
    ).toEqual({
      enable_thinking: false,
    })
    expect(
      toggleThinkingBodyFor({ model: 'qwen3.8-max', thinking: { 'qwen3.8-max': 'medium' } }),
    ).toEqual({ enable_thinking: true, thinking_budget: 8192 })
  })
})

describe('sanitizeAiSettings keeps a valid thinking map', () => {
  it('stores clean picks and drops unknown levels and oversized keys', () => {
    const settings = sanitizeAiSettings({
      provider: 'openai',
      providers: {
        openai: {
          apiKey: 'sk',
          model: 'gpt-6.1-terra',
          thinking: {
            'gpt-6.1-terra': 'max',
            'gpt-5.5': 'nonsense',
            '': 'low',
            [`${'x'.repeat(300)}`]: 'low',
          },
        },
      },
    })
    expect(settings?.providers.openai.thinking).toEqual({ 'gpt-6.1-terra': 'max' })
  })
})
