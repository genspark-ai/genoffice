import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { MAX_MEMORIES } from '@genoffice/agent-core'

import { AgentInstructionsStore } from '../src/agent-instructions-store'

let dir: string
let store: AgentInstructionsStore

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'agent-memory-'))
  store = new AgentInstructionsStore(dir)
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

const memoryFile = () => join(dir, 'agent-memory.json')

describe('memory storage', () => {
  it('reads as empty before anything is recorded', () => {
    expect(store.readMemories()).toEqual([])
  })

  it('records a preference and reads it back', () => {
    const saved = store.addMemory('  prefers British   English ')
    expect(saved?.text).toBe('prefers British English')
    expect(store.readMemories().map((m) => m.text)).toEqual(['prefers British English'])
  })

  it('refuses text that is not worth storing', () => {
    expect(store.addMemory('   ')).toBeNull()
    expect(store.addMemory(null)).toBeNull()
    expect(store.readMemories()).toEqual([])
  })

  it('refreshes a restated preference instead of storing it twice', () => {
    const first = store.addMemory('dislikes emoji')
    const again = store.addMemory('Dislikes Emoji')

    const all = store.readMemories()
    expect(all).toHaveLength(1)
    expect(all[0]!.id).toBe(first!.id)
    expect(again!.createdAt).toBeGreaterThanOrEqual(first!.createdAt)
  })

  it('returns newest first', () => {
    store.addMemory('first')
    store.addMemory('second')
    expect(store.readMemories()[0]!.text).toBe('second')
  })

  it('deletes by id and reports whether anything went', () => {
    const saved = store.addMemory('temporary')
    expect(store.deleteMemory(saved!.id)).toBe(true)
    expect(store.readMemories()).toEqual([])
    expect(store.deleteMemory(saved!.id)).toBe(false)
  })

  it('keeps the file bounded however much the agent writes', () => {
    for (let i = 0; i < MAX_MEMORIES + 25; i++) store.addMemory(`preference ${i}`)
    expect(store.readMemories()).toHaveLength(MAX_MEMORIES)
  })

  it('survives a corrupt file rather than taking AI features down', () => {
    writeFileSync(memoryFile(), '{ not json', 'utf-8')
    expect(store.readMemories()).toEqual([])
  })

  it('skips entries that lost their text or id rather than surfacing blanks', () => {
    writeFileSync(
      memoryFile(),
      JSON.stringify([
        { id: 'ok', text: 'kept', createdAt: 2 },
        { id: '', text: 'no id', createdAt: 1 },
        { id: 'blank', text: '   ', createdAt: 1 },
        'not even an object',
      ]),
      'utf-8',
    )
    expect(store.readMemories().map((m) => m.text)).toEqual(['kept'])
  })

  it('leaves rules and skills alone', () => {
    store.writeRules({ global: 'always cite sources' })
    store.addMemory('prefers tables over prose')

    expect(store.readRules()).toEqual({ global: 'always cite sources' })
    expect(JSON.parse(readFileSync(memoryFile(), 'utf-8'))).toHaveLength(1)
  })
})
