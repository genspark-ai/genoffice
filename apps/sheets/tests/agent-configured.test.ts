import { describe, expect, it } from 'vitest'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { agentConfigured, sendRoute } from '../src/ai/agent-configured'
import {
  isOfflinePrompt,
  planPrompt,
  UnsupportedPromptError,
} from '../src/ai/deterministic-planner'

// Codex CLI stores no API key and may leave the model empty; requiring either
// sent Codex users to the offline planner instead of the agent.
describe('agentConfigured', () => {
  it('is false without settings', () => {
    expect(agentConfigured(undefined)).toBe(false)
  })

  it('accepts genspark without a stored key (main injects the gsk key)', () => {
    const settings = defaultAiSettings()
    expect(settings.provider).toBe('genspark')
    expect(agentConfigured(settings)).toBe(true)
  })

  it('accepts codex with no API key and no model (CLI-side auth and default model)', () => {
    const settings = defaultAiSettings()
    settings.provider = 'codex'
    expect(settings.providers.codex.apiKey).toBe('')
    expect(settings.providers.codex.model).toBe('')
    expect(agentConfigured(settings)).toBe(true)
  })

  it('accepts codex with a picked model, still keyless', () => {
    const settings = defaultAiSettings()
    settings.provider = 'codex'
    settings.providers.codex.model = 'gpt-6.1-sol'
    expect(agentConfigured(settings)).toBe(true)
  })

  it('still requires a key for API-key providers', () => {
    const settings = defaultAiSettings()
    settings.provider = 'anthropic'
    expect(agentConfigured(settings)).toBe(false)
    settings.providers.anthropic.apiKey = 'sk-ant-test'
    expect(agentConfigured(settings)).toBe(true)
  })

  it('still requires a model for API-key providers', () => {
    const settings = defaultAiSettings()
    settings.provider = 'anthropic'
    settings.providers.anthropic.apiKey = 'sk-ant-test'
    settings.providers.anthropic.model = ''
    expect(agentConfigured(settings)).toBe(false)
  })

  it('accepts a custom endpoint whose key is named by environment variable', () => {
    const settings = defaultAiSettings()
    settings.provider = 'custom'
    settings.providers.custom.model = 'local-model'
    expect(agentConfigured(settings)).toBe(false)
    settings.providers.custom.profile = { apiKeyEnv: 'MY_GATEWAY_KEY' }
    expect(agentConfigured(settings)).toBe(true)
  })
})

// genoffice#2967: a provider picked without a key answered every message with
// the offline planner's syntax hint instead of the missing-key error.
describe('sendRoute', () => {
  const keyless = () => {
    const settings = defaultAiSettings()
    settings.provider = 'openai'
    settings.providers.openai.model = 'gpt-6.2-sol'
    return settings
  }

  it('runs the agent for a configured provider, whatever the message', () => {
    const settings = keyless()
    settings.providers.openai.apiKey = 'sk-test'
    expect(sendRoute(settings, 'hello')).toBe('agent')
    expect(sendRoute(settings, 'set A1 to 42')).toBe('agent')
  })

  it('sends free text to the agent path when the setup is incomplete, so the real error shows', () => {
    expect(sendRoute(keyless(), 'hello')).toBe('agent')
    const noModel = keyless()
    noModel.providers.openai.apiKey = 'sk-test'
    noModel.providers.openai.model = ''
    expect(sendRoute(noModel, 'sum column B')).toBe('agent')
  })

  it('keeps the offline planner for the commands it understands', () => {
    expect(sendRoute(keyless(), 'set A1 to 42')).toBe('offline')
    expect(sendRoute(keyless(), 'formula B1 = SUM(A1:A10)')).toBe('offline')
    expect(sendRoute(keyless(), 'rename sheet to Budget')).toBe('offline')
  })

  it('stays offline until settings have loaded', () => {
    expect(sendRoute(null, 'hello')).toBe('offline')
  })
})

describe('planPrompt', () => {
  const options = { revision: 0, sheetId: 'sheet-1' }

  it('accepts an example pasted back with the quotes the hint wraps it in', () => {
    const quoted = planPrompt('\u201cset A1 to 42\u201d', options)
    expect(quoted.operations).toEqual([
      { op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 42 },
    ])
    expect(planPrompt('"rename sheet to Budget"', options).operations[0]).toMatchObject({
      op: 'rename_sheet',
      name: 'Budget',
    })
  })

  it('does not unwrap several quoted examples pasted together', () => {
    const pasted =
      '\u201cset A1 to 42\u201d, \u201cformula B1 = SUM(A1:A10)\u201d, or \u201crename sheet to Budget\u201d'
    expect(() => planPrompt(pasted, options)).toThrow(UnsupportedPromptError)
    expect(isOfflinePrompt(pasted)).toBe(false)
  })

  it('leaves quotes inside a value alone', () => {
    expect(planPrompt('set A1 to "hi"', options).operations[0]).toMatchObject({ value: '"hi"' })
  })
})
