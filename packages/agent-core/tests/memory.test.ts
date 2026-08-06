import { describe, expect, it } from 'vitest'
import { createMemoryTools, type MemoryStoreAdapter } from '../src/memory'

function fakeAdapter(): MemoryStoreAdapter & { entriesList: { id: string; text: string; ts: string }[] } {
  const entriesList: { id: string; text: string; ts: string }[] = []
  return {
    entriesList,
    available: () => true,
    entries: () => entriesList,
    add: async (text) => {
      const entry = { id: `id-${entriesList.length}`, text, ts: new Date().toISOString() }
      entriesList.push(entry)
      return entry
    },
    remove: async (id) => {
      const idx = entriesList.findIndex((e) => e.id === id)
      if (idx >= 0) entriesList.splice(idx, 1)
    },
  }
}

describe('createMemoryTools', () => {
  it('exposes remember_memory and forget_memory tool definitions', () => {
    const mt = createMemoryTools(fakeAdapter())
    expect(mt.tools.map((t) => t.name)).toEqual(['remember_memory', 'forget_memory'])
  })

  it('remember_memory persists an entry and reports its id', async () => {
    const adapter = fakeAdapter()
    const mt = createMemoryTools(adapter)
    const result = await mt.execute({
      id: 't1',
      name: 'remember_memory',
      input: { text: '  prefers US date formats  ' },
    })
    expect(result.isError).not.toBe(true)
    expect(result.output).toContain('id id-0')
    expect(adapter.entriesList).toHaveLength(1)
    expect(adapter.entriesList[0].text).toBe('prefers US date formats')
  })

  it('remember_memory rejects empty text', async () => {
    const mt = createMemoryTools(fakeAdapter())
    const result = await mt.execute({
      id: 't2',
      name: 'remember_memory',
      input: { text: '   ' },
    })
    expect(result.isError).toBe(true)
  })

  it('forget_memory deletes an entry by id', async () => {
    const adapter = fakeAdapter()
    const mt = createMemoryTools(adapter)
    await mt.execute({ id: 't1', name: 'remember_memory', input: { text: 'alpha' } })
    const result = await mt.execute({ id: 't2', name: 'forget_memory', input: { id: 'id-0' } })
    expect(result.isError).not.toBe(true)
    expect(adapter.entriesList).toHaveLength(0)
  })

  it('contextSection lists entries and is empty when unavailable', () => {
    const adapter = fakeAdapter()
    const mt = createMemoryTools(adapter)
    expect(mt.contextSection()).toContain('(empty)')
    adapter.available = () => false
    expect(mt.contextSection()).toBe('')
  })

  it('contextSection shows saved entries with ids', async () => {
    const adapter = fakeAdapter()
    const mt = createMemoryTools(adapter)
    await mt.execute({ id: 't1', name: 'remember_memory', input: { text: 'uses 12pt font' } })
    const section = mt.contextSection()
    expect(section).toContain('uses 12pt font (id: id-0)')
  })

  it('unknown memory tool errors', async () => {
    const mt = createMemoryTools(fakeAdapter())
    const result = await mt.execute({ id: 't9', name: 'nope', input: {} })
    expect(result.isError).toBe(true)
  })
})
