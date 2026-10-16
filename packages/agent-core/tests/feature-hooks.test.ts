import { describe, expect, it } from 'vitest'
import { HOOK_EVENTS, describeHooks, hooksFor, validateHook } from '../src/feature-hooks'
import type { FeatureRecord } from '../src/features-skill'

function record(id: string, on?: string): FeatureRecord {
  return {
    id,
    name: id,
    description: 'x',
    kind: on ? 'hook' : '',
    ...(on ? { hook: { on: on as never, script: { source: '$.log(1)' } } } : {}),
    createdAt: 'x',
    updatedAt: 'x',
  }
}

describe('hook contract', () => {
  it('accepts the events the host raises', () => {
    for (const event of HOOK_EVENTS) {
      const result = validateHook({ on: event, script: { source: '$.log(input.event.path)' } })
      expect(result.ok).toBe(true)
    }
  })

  it('refuses an event nobody raises, naming the field', () => {
    const result = validateHook({ on: 'document.printed', script: { source: '$.log(1)' } })
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.errors[0]!.field).toBe('hook.on')
      expect(result.errors[0]!.message).toContain('document.opened')
    }
  })

  it('validates the script with the script contract, and names it as hook.source', () => {
    const missing = validateHook({ on: 'document.opened' })
    expect(missing.ok).toBe(false)
    if (!missing.ok) expect(missing.errors[0]!.field).toBe('hook.script')

    const module = validateHook({ on: 'document.opened', script: { source: 'export default 1' } })
    expect(module.ok).toBe(false)
    if (!module.ok) expect(module.errors[0]!.field).toBe('hook.script.source')

    expect(validateHook('opened').ok).toBe(false)
    expect(validateHook(null).ok).toBe(false)
  })

  it('picks the hooks for one event, and describes them for the assistant', () => {
    const features = [record('a', 'document.opened'), record('b', 'document.closed'), record('c')]
    expect(hooksFor(features, 'document.opened').map((f) => f.id)).toEqual(['a'])
    expect(hooksFor(features, 'document.closed').map((f) => f.id)).toEqual(['b'])
    expect(describeHooks(features)).toEqual([
      'a observes document.opened',
      'b observes document.closed',
    ])
  })
})
