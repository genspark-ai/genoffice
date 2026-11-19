import { describe, expect, it } from 'vitest'
import {
  customChanges,
  customEntries,
  customValueFits,
  summaryChanges,
  summaryValues,
} from '../src/renderer/components/PropertiesDialog'

const SAVED = {
  core: { title: 'Report', creator: 'Ada', revision: '4' },
  app: { company: 'Analytical Engines' },
}

describe('File > Properties', () => {
  it('shows the saved properties with the edits waiting for the next save on top', () => {
    const values = summaryValues(SAVED, { core: { title: 'Final report', creator: null } })
    expect(values['core.title']).toBe('Final report')
    expect(values['core.creator']).toBe('')
    expect(values['app.company']).toBe('Analytical Engines')
    expect(values['core.subject']).toBe('')
  })

  it('writes only what changed, and an emptied field removes its property', () => {
    const before = summaryValues(SAVED, null)
    expect(summaryChanges(before, before)).toBeNull()
    const after = { ...before, 'core.title': ' Plan ', 'core.creator': '', 'app.manager': 'Grace' }
    expect(summaryChanges(before, after)).toEqual({
      core: { title: 'Plan', creator: null },
      app: { manager: 'Grace' },
    })
    expect(summaryChanges(before, { ...before, 'core.keywords': 'a, b' })).toEqual({
      core: { keywords: 'a, b' },
    })
  })

  const CUSTOM = {
    custom: {
      Client: 'Example Ltd',
      Pages: '42',
      Final: 'true',
      Filed: '2026-01-02T00:00:00Z',
      Parts: 'a',
    },
    customTypes: {
      Client: 'lpwstr',
      Pages: 'i4',
      Final: 'bool',
      Filed: 'filetime',
      Parts: 'vector',
    },
  }

  it('lists the custom properties by the four types of the tab, pending edits on top', () => {
    expect(customEntries(CUSTOM, null)).toEqual([
      { name: 'Client', kind: 'text', value: 'Example Ltd' },
      { name: 'Pages', kind: 'number', value: '42' },
      { name: 'Final', kind: 'bool', value: 'true' },
      { name: 'Filed', kind: 'date', value: '2026-01-02' },
      // a stored type the tab has no type for
      { name: 'Parts', kind: null, value: 'a' },
    ])
    const pending = {
      custom: { Client: null, Pages: '7', Due: '2026-11-17', Count: 3, Done: false },
      customKinds: { Pages: 'text', Due: 'date' },
    } as const
    expect(customEntries(CUSTOM, pending)).toEqual([
      { name: 'Pages', kind: 'text', value: '7' },
      { name: 'Final', kind: 'bool', value: 'true' },
      { name: 'Filed', kind: 'date', value: '2026-01-02' },
      { name: 'Parts', kind: null, value: 'a' },
      { name: 'Due', kind: 'date', value: '2026-11-17' },
      { name: 'Count', kind: 'number', value: '3' },
      { name: 'Done', kind: 'bool', value: 'false' },
    ])
  })

  it('writes the custom properties that were added, changed or deleted, each with its type', () => {
    const before = customEntries(CUSTOM, null)
    expect(customChanges(before, before)).toBeNull()
    const after = [
      // Client deleted, Pages retyped, Final flipped, Filed and Parts untouched, Due added
      { name: 'Pages', kind: 'text', value: '42' },
      { name: 'Final', kind: 'bool', value: 'false' },
      before[3],
      before[4],
      { name: 'Due', kind: 'date', value: '2026-11-17' },
    ] as const
    expect(customChanges(before, after)).toEqual({
      custom: { Client: null, Pages: '42', Final: 'false', Due: '2026-11-17' },
      customKinds: { Pages: 'text', Final: 'bool', Due: 'date' },
    })
    // a deletion alone
    expect(customChanges(before, before.slice(1))).toEqual({
      custom: { Client: null },
      customKinds: {},
    })
  })

  it('a value has to fit its type before it can be added', () => {
    expect(customValueFits('text', 'x')).toBe(true)
    expect(customValueFits('text', '  ')).toBe(false)
    expect(customValueFits('number', '-12.5')).toBe(true)
    expect(customValueFits('number', '12 apples')).toBe(false)
    expect(customValueFits('number', '')).toBe(false)
    expect(customValueFits('date', '2026-11-17')).toBe(true)
    expect(customValueFits('date', '2026-13-40')).toBe(false)
    expect(customValueFits('date', 'tomorrow')).toBe(false)
    expect(customValueFits('bool', 'true')).toBe(true)
    expect(customValueFits('bool', 'yes')).toBe(false)
  })
})
