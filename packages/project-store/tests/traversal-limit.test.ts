import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'

function makeTempDir(): string {
  return mkdtempSync(join(tmpdir(), 'project-store-traversal-limit-'))
}

// Seed N assistant messages so they are written to disk immediately
function seedAssistantMessages(store: ProjectStore, chatId: string, count: number): void {
  for (let i = 0; i < count; i++) {
    store.appendChatMessage('default', chatId, { role: 'assistant', text: `msg${i}` })
  }
}

describe('projectId/chatId traversal rejection', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('loadChat rejects ../evil chatId', () => {
    expect(() => store.loadChat('default', '../evil')).toThrow(/Invalid chatId/)
  })

  it('loadChat rejects traversal projectId', () => {
    expect(() => store.loadChat('../evil', 'chat1')).toThrow(/Invalid projectId/)
  })

  it('loadChat rejects slash and backslash in chatId', () => {
    expect(() => store.loadChat('default', 'a/b')).toThrow(/Invalid chatId/)
    expect(() => store.loadChat('default', 'a\\b')).toThrow(/Invalid chatId/)
  })

  it('loadChat rejects dot-dot and empty ids', () => {
    expect(() => store.loadChat('default', '..')).toThrow(/Invalid chatId/)
    expect(() => store.loadChat('default', '')).toThrow(/Invalid chatId/)
  })

  it('appendChatMessage throws fail-closed on traversal ids', () => {
    expect(() =>
      store.appendChatMessage('default', '../evil', { role: 'assistant', text: 'x' }),
    ).toThrow(/Invalid chatId/)
    expect(() =>
      store.appendChatMessage('../evil', 'chat1', { role: 'assistant', text: 'x' }),
    ).toThrow(/Invalid projectId/)
  })

  it('listChats rejects traversal projectId', () => {
    expect(() => store.listChats('../evil')).toThrow(/Invalid projectId/)
  })

  it('valid ids still work', () => {
    store.appendChatMessage('default', 'ok-chat_123', { role: 'assistant', text: 'hi' })
    expect(store.loadChat('default', 'ok-chat_123')).toHaveLength(1)
  })
})

describe('loadChat limit clamping', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('normal limit returns the most recent N messages', () => {
    seedAssistantMessages(store, 'chat-normal', 5)
    const msgs = store.loadChat('default', 'chat-normal', 2)
    expect(msgs).toHaveLength(2)
    expect(msgs.map((m) => m.text)).toEqual(['msg3', 'msg4'])
  })

  it('limit=0 returns 1 message instead of all', () => {
    seedAssistantMessages(store, 'chat-zero', 5)
    const msgs = store.loadChat('default', 'chat-zero', 0)
    expect(msgs).toHaveLength(1)
    expect(msgs[0].text).toBe('msg4')
  })

  it('negative limit clamps to 1', () => {
    seedAssistantMessages(store, 'chat-neg', 5)
    const msgs = store.loadChat('default', 'chat-neg', -1)
    expect(msgs).toHaveLength(1)
    expect(msgs[0].text).toBe('msg4')
  })

  it('NaN falls back to the default limit', () => {
    seedAssistantMessages(store, 'chat-nan', 250)
    const msgs = store.loadChat('default', 'chat-nan', NaN)
    expect(msgs).toHaveLength(200)
    expect(msgs[0].text).toBe('msg50')
  })

  it('Infinity falls back to the default limit', () => {
    seedAssistantMessages(store, 'chat-inf', 250)
    const msgs = store.loadChat('default', 'chat-inf', Infinity)
    expect(msgs).toHaveLength(200)
    expect(msgs[0].text).toBe('msg50')
  })

  it('fractional limits are floored', () => {
    seedAssistantMessages(store, 'chat-frac', 5)
    const msgs = store.loadChat('default', 'chat-frac', 2.9)
    expect(msgs).toHaveLength(2)
  })

  it('oversized limits are capped and do not crash', () => {
    seedAssistantMessages(store, 'chat-big', 5)
    const msgs = store.loadChat('default', 'chat-big', 20000)
    expect(msgs).toHaveLength(5)
  })
})
