import { describe, expect, it } from 'vitest'
import { SCRIPT_API_SUMMARY, SCRIPT_LIMITS, scriptsIn, validateScript } from '../src/feature-script'
import type { FeatureRecord } from '../src/features-skill'

function record(id: string, source?: string): FeatureRecord {
  return {
    id,
    name: id,
    description: 'x',
    kind: source ? 'script' : '',
    ...(source ? { script: { source } } : {}),
    createdAt: 'x',
    updatedAt: 'x',
  }
}

describe('script contract', () => {
  it('accepts a body that reads and proposes', () => {
    const result = validateScript({
      source: [
        'var blocks = $.docs.blocks()',
        '$.log("blocks=" + blocks.length)',
        '$.docs.propose([{ op: "findReplace", find: "a", replace: "b" }])',
      ].join('\n'),
    })
    expect(result.ok).toBe(true)
  })

  it('refuses an empty script and a module', () => {
    expect(validateScript({ source: '   ' }).ok).toBe(false)
    expect(validateScript({}).ok).toBe(false)
    expect(validateScript('$.log(1)').ok).toBe(false)
    const module = validateScript({ source: 'export default function () {}' })
    expect(module.ok).toBe(false)
    if (!module.ok) {
      expect(module.errors[0]!.field).toBe('script.source')
      expect(module.errors[0]!.message).toContain('body of a function')
    }
  })

  it('caps the source at the byte limit', () => {
    const big = validateScript({ source: 'x'.repeat(SCRIPT_LIMITS.maxSourceBytes + 1) })
    expect(big.ok).toBe(false)
    if (!big.ok) expect(big.errors[0]!.message).toContain('more than the')
  })

  it('tells the assistant what a script may do', () => {
    expect(SCRIPT_API_SUMMARY).toContain('$.docs.propose')
    expect(SCRIPT_API_SUMMARY).toContain('read-only')
  })

  it('lists the scripts a registry holds', () => {
    expect(scriptsIn([record('a', '$.log(1)'), record('b')]).map((f) => f.id)).toEqual(['a'])
    expect(scriptsIn([])).toEqual([])
  })
})
