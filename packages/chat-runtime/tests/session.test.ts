/**
 * Persistence tests — JSONL round-trip and in-memory adapter contract.
 */
import { describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { JsonlChatPersistence, MemoryChatPersistence } from '../src/persistence.js'
import type { ChatMessage, ChatSession } from '../src/types.js'

function makeSession(id: string): ChatSession {
  return {
    id,
    app: 'docs',
    title: 'Test',
    createdAt: 1700000000000,
    updatedAt: 1700000000000,
    messages: [],
  }
}

function makeMessage(id: string, content: string, ts: number): ChatMessage {
  return { id, role: 'user', content, timestamp: ts, status: 'done' }
}

describe('MemoryChatPersistence', () => {
  it('round-trips a session', async () => {
    const p = new MemoryChatPersistence()
    const session = makeSession('m1')
    await p.save(session)
    const loaded = await p.load('m1')
    expect(loaded?.id).toBe('m1')
    expect(loaded?.messages).toEqual([])
  })

  it('appends messages in order', async () => {
    const p = new MemoryChatPersistence()
    await p.save(makeSession('m2'))
    await p.append('m2', makeMessage('msg-1', 'first', 1))
    await p.append('m2', makeMessage('msg-2', 'second', 2))
    const loaded = await p.load('m2')
    expect(loaded?.messages.map(m => m.content)).toEqual(['first', 'second'])
  })
})

describe('JsonlChatPersistence', () => {
  it('round-trips messages through disk', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'chat-runtime-'))
    try {
      const p = new JsonlChatPersistence({ dir })
      await p.save(makeSession('j1'))
      await p.append('j1', makeMessage('msg-1', 'hello', 100))
      await p.append('j1', makeMessage('msg-2', 'world', 200))

      const loaded = await p.load('j1')
      expect(loaded?.messages).toHaveLength(2)
      expect(loaded?.messages[0].content).toBe('hello')
      expect(loaded?.messages[1].content).toBe('world')

      // A second load returns the same data.
      const reloaded = await p.load('j1')
      expect(reloaded?.messages).toHaveLength(2)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('returns null for missing session ids', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'chat-runtime-'))
    try {
      const p = new JsonlChatPersistence({ dir })
      expect(await p.load('not-there')).toBeNull()
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
