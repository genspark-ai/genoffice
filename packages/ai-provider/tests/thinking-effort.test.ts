import { afterEach, describe, expect, it, vi } from 'vitest'

import { sanitizeAiSettings } from '../src/ai-settings-guard'
import { chatForProvider } from '../src/chat'
import {
  anthropicThinkingFields,
  geminiThinkingLevelFor,
  openAiEffortFor,
  resolveThinkingEffort,
  thinkingShapeFor,
  toggleThinkingBodyFor,
} from '../src/thinking-effort'
import { jsonResponse } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

/** the JSON body of the most recent mocked fetch call */
function lastBody(fetchMock: ReturnType<typeof vi.fn>): Record<string, unknown> {
  const call = fetchMock.mock.calls.at(-1)!
  return JSON.parse(call[1].body as string) as Record<string, unknown>
}

describe('thinkingShapeFor', () => {
  it('openai: the 6.x ladder, with the upper rungs gated on the Responses route', () => {
    expect(thinkingShapeFor('gpt-6.1-terra', 'openai-responses').levels).toEqual([
      'none',
      'low',
      'medium',
      'high',
      'xhigh',
      'max',
    ])
    // the same id behind a chat-completions gateway only takes low..high
    expect(thinkingShapeFor('gpt-6.1-terra', 'openai-compatible').levels).toEqual([
      'none',
      'low',
      'medium',
      'high',
    ])
    // no route known yet: the conservative surface
    expect(thinkingShapeFor('gpt-6.1-terra').levels).toEqual(['none', 'low', 'medium', 'high'])
    expect(thinkingShapeFor('gpt-6-astra', 'openai-responses').levels).toEqual([
      'low',
      'medium',
      'high',
      'xhigh',
      'max',
    ])
    expect(thinkingShapeFor('gpt-6-sol', 'openai-responses').levels).not.toContain('none')
    expect(thinkingShapeFor('gpt-6.1-sol', 'openai-responses').levels).not.toContain('none')
    expect(thinkingShapeFor('gpt-6.1-astra', 'openai-responses').levels).not.toContain('none')
    expect(thinkingShapeFor('gpt-6-luna', 'openai-responses').levels).toContain('none')
    expect(thinkingShapeFor('gpt-5.6-terra', 'openai-responses').levels).toEqual([
      'none',
      'low',
      'medium',
      'high',
    ])
    expect(thinkingShapeFor('gpt-5.5').kind).toBe('openai-effort')
    expect(thinkingShapeFor('o3').kind).toBe('openai-effort')
  })

  it('anthropic: adaptive on 4.6+, nothing below (4.5 and older 400 on adaptive)', () => {
    expect(thinkingShapeFor('claude-opus-5-5')).toEqual({
      kind: 'anthropic-adaptive',
      levels: ['low', 'medium', 'high', 'xhigh', 'max'],
    })
    expect(thinkingShapeFor('claude-sonnet-4-6').levels).toEqual(['low', 'medium', 'high'])
    // 4.5 and older reject `{ type: 'adaptive' }`: no knob at all
    expect(thinkingShapeFor('claude-sonnet-4-5').levels).toEqual([])
    expect(thinkingShapeFor('claude-opus-4-5').levels).toEqual([])
    expect(thinkingShapeFor('claude-haiku-4-5').levels).toEqual([])
    expect(thinkingShapeFor('claude-3-7-sonnet').levels).toEqual([])
    // a date suffix is not a minor version: 4-20250514 is 4.0, not 4.20250514
    expect(thinkingShapeFor('claude-sonnet-4-20250514').levels).toEqual([])
    expect(thinkingShapeFor('claude-opus-4-1-20250805').levels).toEqual([])
    expect(thinkingShapeFor('claude-sonnet-4-6').kind).toBe('anthropic-adaptive')
  })

  it('gemini 3/4: low, medium and high', () => {
    expect(thinkingShapeFor('gemini-4-argon')).toEqual({
      kind: 'gemini-level',
      levels: ['low', 'medium', 'high'],
    })
    expect(thinkingShapeFor('gemini-3.8-flash').levels).toEqual(['low', 'medium', 'high'])
  })

  it('families whose surface is still documentation-only carry no knob', () => {
    for (const id of [
      'kimi-k3.1-pro',
      'glm-5.2',
      'glm-5.3',
      'doubao-seed-2-1-pro-260628',
      'hy4-preview',
      'mistral-medium-latest',
    ]) {
      expect(thinkingShapeFor(id).levels, id).toEqual([])
    }
  })

  it('minimax m3.1 takes the documented effort rungs, without an off; plain m3 stays knobless', () => {
    // the vendor guide documents reasoning_effort for M3.1-Flash only; the
    // plain M3 has no documented surface, so per the module rule it waits
    expect(thinkingShapeFor('MiniMax-M3')).toEqual({ kind: 'openai-effort', levels: [] })
    expect(thinkingShapeFor('minimax-m3.1-flash')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'medium', 'high', 'xhigh', 'max'],
    })
    expect(thinkingShapeFor('MiniMax-M3.1')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'medium', 'high', 'xhigh', 'max'],
    })
  })

  it('deepseek takes the OpenAI effort ladder on both id spellings, every rung probed', () => {
    expect(thinkingShapeFor('deepseek-flash')).toEqual({
      kind: 'openai-effort',
      levels: ['none', 'low', 'medium', 'high', 'xhigh', 'max'],
    })
    expect(thinkingShapeFor('deepseek-v4-pro')).toEqual({
      kind: 'openai-effort',
      levels: ['none', 'low', 'medium', 'high', 'xhigh', 'max'],
    })
    // the GenOffice-facing alias spells it with the hyphen (DEEPSEEK_V41_FLASH)
    expect(thinkingShapeFor('deep-seek-v4.1-flash')).toEqual({
      kind: 'openai-effort',
      levels: ['none', 'low', 'medium', 'high', 'xhigh', 'max'],
    })
    // DeepSeek's chat-completions route accepts the upper rungs as-is (probed
    // 200, 2026-11-08), so no Responses-only gating applies
    expect(thinkingShapeFor('deepseek-flash', 'openai-compatible')).toEqual(
      thinkingShapeFor('deepseek-flash'),
    )
  })

  it('ling/spark/mimo are switches; qwen and grok keep their documented fields', () => {
    expect(thinkingShapeFor('ling-3.0-flash')).toEqual({
      kind: 'toggle',
      levels: ['none', 'medium'],
    })
    expect(thinkingShapeFor('spark-x2.5')).toEqual({ kind: 'toggle', levels: ['none', 'medium'] })
    expect(thinkingShapeFor('mimo-v2.6-pro')).toEqual({
      kind: 'toggle',
      levels: ['none', 'medium'],
    })
    expect(thinkingShapeFor('qwen3.8-max')).toEqual({
      kind: 'qwen-budget',
      levels: ['none', 'low', 'medium', 'high'],
    })
    // grok documents low/medium/high, xhigh from 4.6 on, and no none
    expect(thinkingShapeFor('grok-4.7')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'medium', 'high', 'xhigh'],
    })
    expect(thinkingShapeFor('grok-4.6')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'medium', 'high', 'xhigh'],
    })
    expect(thinkingShapeFor('grok-4.5')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'medium', 'high'],
    })
    expect(thinkingShapeFor('grok-4-0709')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'medium', 'high'],
    })
    // o-series ids document low/medium/high only as well
    expect(thinkingShapeFor('o3')).toEqual({
      kind: 'openai-effort',
      levels: ['low', 'medium', 'high'],
    })
    expect(thinkingShapeFor('o4-mini').levels).toEqual(['low', 'medium', 'high'])
    expect(thinkingShapeFor('LongCat-2.0').levels).toEqual([])
  })
})

describe('resolveThinkingEffort', () => {
  it('resolves the stored pick only when the family and route accept it', () => {
    expect(
      resolveThinkingEffort(
        { model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'max' } },
        'openai-responses',
      ),
    ).toBe('max')
    // the same pick is unsendable through chat-completions
    expect(
      resolveThinkingEffort(
        { model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'max' } },
        'openai-compatible',
      ),
    ).toBeUndefined()
    expect(
      resolveThinkingEffort(
        { model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'high' } },
        'openai-compatible',
      ),
    ).toBe('high')
    expect(
      resolveThinkingEffort({ model: 'gpt-6-astra', thinking: { 'gpt-6-astra': 'none' } }),
    ).toBeUndefined()
    expect(
      resolveThinkingEffort({ model: 'gemini-4-argon', thinking: { 'gemini-4-argon': 'max' } }),
    ).toBeUndefined()
    // none is not on any adaptive Claude ladder
    expect(
      resolveThinkingEffort({ model: 'claude-opus-5-5', thinking: { 'claude-opus-5-5': 'none' } }),
    ).toBeUndefined()
    expect(
      resolveThinkingEffort({
        model: 'claude-sonnet-4-6',
        thinking: { 'claude-sonnet-4-6': 'none' },
      }),
    ).toBeUndefined()
    expect(
      resolveThinkingEffort({
        model: 'claude-sonnet-4-6',
        thinking: { 'claude-sonnet-4-6': 'high' },
      }),
    ).toBe('high')
    // 4.5 has no knob: its old file picks resolve to nothing sent
    expect(
      resolveThinkingEffort({
        model: 'claude-sonnet-4-5',
        thinking: { 'claude-sonnet-4-5': 'high' },
      }),
    ).toBeUndefined()
    expect(
      resolveThinkingEffort({ model: 'grok-4.7', thinking: { 'grok-4.7': 'none' } }),
    ).toBeUndefined()
    expect(resolveThinkingEffort({ model: 'o3', thinking: { o3: 'none' } })).toBeUndefined()
    expect(
      resolveThinkingEffort({ model: 'hy4-preview', thinking: { 'hy4-preview': 'high' } }),
    ).toBeUndefined()
    expect(resolveThinkingEffort({ model: 'gpt-6.1-terra', thinking: {} })).toBeUndefined()
  })
})

describe('per-protocol translators', () => {
  it('openai effort is gated on the Responses route for the upper rungs', () => {
    expect(
      openAiEffortFor(
        { model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'xhigh' } },
        'openai-responses',
      ),
    ).toBe('xhigh')
    expect(
      openAiEffortFor(
        { model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'xhigh' } },
        'openai-compatible',
      ),
    ).toBeUndefined()
    expect(openAiEffortFor({ model: 'grok-4.7', thinking: { 'grok-4.7': 'low' } })).toBe('low')
    expect(openAiEffortFor({ model: 'grok-4.7', thinking: { 'grok-4.7': 'xhigh' } })).toBe('xhigh')
    expect(
      openAiEffortFor({ model: 'grok-4.5', thinking: { 'grok-4.5': 'xhigh' } }),
    ).toBeUndefined()
    expect(
      openAiEffortFor({ model: 'claude-opus-5-5', thinking: { 'claude-opus-5-5': 'high' } }),
    ).toBeUndefined()
  })

  it('anthropic: the adaptive record carries no level; effort is top-level', () => {
    expect(
      anthropicThinkingFields({ model: 'claude-opus-5-5', thinking: { 'claude-opus-5-5': 'max' } }),
    ).toEqual({ thinking: { type: 'adaptive' }, effort: 'max' })
    expect(
      anthropicThinkingFields({
        model: 'claude-sonnet-4-6',
        thinking: { 'claude-sonnet-4-6': 'none' },
      }),
    ).toBeUndefined()
    expect(
      anthropicThinkingFields({
        model: 'claude-sonnet-4-6',
        thinking: { 'claude-sonnet-4-6': 'high' },
      }),
    ).toEqual({ thinking: { type: 'adaptive' }, effort: 'high' })
    // 4.5 rejects the adaptive record outright: nothing is sent
    expect(
      anthropicThinkingFields({
        model: 'claude-opus-4-5',
        thinking: { 'claude-opus-4-5': 'high' },
      }),
    ).toBeUndefined()
    expect(
      anthropicThinkingFields({
        model: 'claude-sonnet-4-6',
        thinking: { 'claude-sonnet-4-6': 'xhigh' },
      }),
    ).toBeUndefined()
    expect(
      anthropicThinkingFields({ model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'high' } }),
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
    expect(
      geminiThinkingLevelFor({
        model: 'gemini-3.8-flash',
        thinking: { 'gemini-3.8-flash': 'medium' },
      }),
    ).toBe('medium')
  })

  it('toggles: qwen flips enable_thinking and its budget; trimmed families send nothing', () => {
    expect(
      toggleThinkingBodyFor({ model: 'qwen3.8-max', thinking: { 'qwen3.8-max': 'low' } }),
    ).toEqual({ enable_thinking: true, thinking_budget: 4096 })
    expect(
      toggleThinkingBodyFor({ model: 'qwen3.8-max', thinking: { 'qwen3.8-max': 'none' } }),
    ).toEqual({ enable_thinking: false })
    expect(
      toggleThinkingBodyFor({ model: 'ling-3.0-flash', thinking: { 'ling-3.0-flash': 'none' } }),
    ).toEqual({ thinking: { type: 'disabled' } })
    expect(
      toggleThinkingBodyFor({
        model: 'doubao-seed-2-1-pro-260628',
        thinking: { 'doubao-seed-2-1-pro-260628': 'medium' },
      }),
    ).toBeUndefined()
  })
})

describe('request wiring', () => {
  it('anthropic sends the adaptive record and the level as output_config.effort', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ content: [{ type: 'text', text: 'ok' }] }))
    vi.stubGlobal('fetch', fetchMock)
    const result = await chatForProvider(
      'anthropic',
      { apiKey: 'k', model: 'claude-opus-5-5', thinking: { 'claude-opus-5-5': 'high' } },
      'sys',
      'hi',
    )
    expect(result).toEqual({ ok: true, content: 'ok' })
    const body = lastBody(fetchMock)
    expect(body.thinking).toEqual({ type: 'adaptive' })
    expect(body.output_config).toEqual({ effort: 'high' })
  })

  it('anthropic sends neither thinking field without a pick', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ content: [{ type: 'text', text: 'ok' }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider('anthropic', { apiKey: 'k', model: 'claude-opus-5-5' }, 'sys', 'hi')
    const body = lastBody(fetchMock)
    expect(body.thinking).toBeUndefined()
    expect(body.output_config).toBeUndefined()
  })

  it('gemini nests the level under generationConfig.thinkingConfig', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ candidates: [{ content: { parts: [{ text: 'ok' }] } }] }))
    vi.stubGlobal('fetch', fetchMock)
    const result = await chatForProvider(
      'gemini',
      { apiKey: 'k', model: 'gemini-4-argon', thinking: { 'gemini-4-argon': 'medium' } },
      'sys',
      'hi',
    )
    expect(result).toEqual({ ok: true, content: 'ok' })
    const generationConfig = lastBody(fetchMock).generationConfig as Record<string, unknown>
    expect(generationConfig.thinkingConfig).toEqual({ thinkingLevel: 'medium' })
    // a flat level is an unknown generationConfig field and 400s
    expect(generationConfig.thinkingLevel).toBeUndefined()
  })

  it('gemini sends no thinkingConfig without a pick', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ candidates: [{ content: { parts: [{ text: 'ok' }] } }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider('gemini', { apiKey: 'k', model: 'gemini-4-argon' }, 'sys', 'hi')
    expect(lastBody(fetchMock).generationConfig).not.toHaveProperty('thinkingConfig')
  })

  it('openai: xhigh rides the Responses route only', async () => {
    const responsesMock = vi.fn().mockResolvedValue(
      jsonResponse({
        output: [{ type: 'message', content: [{ type: 'output_text', text: 'ok' }] }],
      }),
    )
    vi.stubGlobal('fetch', responsesMock)
    await chatForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'xhigh' } },
      'sys',
      'hi',
    )
    expect(responsesMock.mock.calls[0]![0]).toBe('https://api.openai.com/v1/responses')
    expect(lastBody(responsesMock).reasoning).toEqual({ summary: 'auto', effort: 'xhigh' })

    // the same model through a chat-completions gateway drops the unsendable pick
    const chatMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    vi.stubGlobal('fetch', chatMock)
    await chatForProvider(
      'openrouter',
      { apiKey: 'k', model: 'gpt-6.1-terra', thinking: { 'gpt-6.1-terra': 'xhigh' } },
      'sys',
      'hi',
    )
    expect(lastBody(chatMock).reasoning_effort).toBeUndefined()
  })

  it('a stored pick outranks a profile bodyExtras field, for both merge paths', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    vi.stubGlobal('fetch', fetchMock)
    const base = { apiKey: 'k', baseUrl: 'https://gw.example.com/v1' }

    await chatForProvider(
      'custom',
      {
        ...base,
        model: 'gpt-5.5',
        profile: { bodyExtras: { reasoning_effort: 'low' } },
        thinking: { 'gpt-5.5': 'high' },
      },
      'sys',
      'hi',
    )
    expect(lastBody(fetchMock).reasoning_effort).toBe('high')

    // no stored pick: the profile's own field still rides through untouched
    await chatForProvider(
      'custom',
      { ...base, model: 'gpt-5.5', profile: { bodyExtras: { reasoning_effort: 'low' } } },
      'sys',
      'hi',
    )
    expect(lastBody(fetchMock).reasoning_effort).toBe('low')

    await chatForProvider(
      'custom',
      {
        ...base,
        model: 'qwen3.8-max',
        profile: { bodyExtras: { enable_thinking: false } },
        thinking: { 'qwen3.8-max': 'medium' },
      },
      'sys',
      'hi',
    )
    const toggleBody = lastBody(fetchMock)
    expect(toggleBody.enable_thinking).toBe(true)
    expect(toggleBody.thinking_budget).toBe(8192)
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
