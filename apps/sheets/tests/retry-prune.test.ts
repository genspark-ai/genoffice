import { describe, expect, it } from 'vitest'
import type { AiChatMessage } from '../src/renderer/ai/AiChatPanel'
import { pruneFailedExchange } from '../src/renderer/ai/retry-prune'

const user = (text: string, undelivered?: boolean): AiChatMessage => ({
  role: 'user',
  text,
  tools: [],
  ...(undelivered ? { undelivered: true } : {}),
})

const assistant = (text: string, isError?: boolean): AiChatMessage => ({
  role: 'assistant',
  text,
  tools: [],
  ...(isError ? { isError: true } : {}),
})

describe('pruneFailedExchange', () => {
  it('drops the failed user bubble and its paired error reply', () => {
    const chat = [user('a'), assistant('ok'), user('b', true), assistant('HTTP 429', true)]
    expect(pruneFailedExchange(chat, 2)).toEqual([user('a'), assistant('ok')])
  })

  it('drops a solo undelivered bubble when no error reply follows', () => {
    const chat = [user('a'), assistant('ok'), user('b', true)]
    expect(pruneFailedExchange(chat, 2)).toEqual([user('a'), assistant('ok')])
  })

  it('keeps a following assistant reply that is not an error', () => {
    const chat = [user('b', true), assistant('answer')]
    expect(pruneFailedExchange(chat, 0)).toEqual([assistant('answer')])
  })

  it('prunes an older failed exchange mid-transcript', () => {
    const chat = [user('a', true), assistant('boom', true), user('b'), assistant('ok')]
    expect(pruneFailedExchange(chat, 0)).toEqual([user('b'), assistant('ok')])
  })

  it('is a no-op for delivered messages, assistant entries, and bad indexes', () => {
    const chat = [user('a'), assistant('ok')]
    expect(pruneFailedExchange(chat, 0)).toBe(chat)
    expect(pruneFailedExchange(chat, 1)).toBe(chat)
    expect(pruneFailedExchange(chat, 7)).toBe(chat)
  })
})
