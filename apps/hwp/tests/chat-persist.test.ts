import { describe, expect, it } from 'vitest'
import {
  mapPersistedChat,
  shouldApplyRestoredChat,
  shouldLoadAfterRebind,
} from '../src/renderer/ai/chat-persist'

describe('shouldLoadAfterRebind', () => {
  it('reloads when an unsaved bind becomes a file chat and the UI is empty', () => {
    expect(
      shouldLoadAfterRebind({
        previousChatId: 'unsaved-1',
        nextChatId: 'file-abc',
        chatAlreadyHasMessages: false,
        skipped: false,
      }),
    ).toBe(true)
  })

  it('does not clobber a live transcript after the first save', () => {
    expect(
      shouldLoadAfterRebind({
        previousChatId: 'unsaved-1',
        nextChatId: 'file-abc',
        chatAlreadyHasMessages: true,
        skipped: false,
      }),
    ).toBe(false)
  })

  it('skips after New chat', () => {
    expect(
      shouldLoadAfterRebind({
        previousChatId: 'unsaved-1',
        nextChatId: 'file-abc',
        chatAlreadyHasMessages: false,
        skipped: true,
      }),
    ).toBe(false)
  })
})

describe('shouldApplyRestoredChat', () => {
  it('applies a non-empty history onto an idle empty panel', () => {
    expect(
      shouldApplyRestoredChat({
        messageCount: 2,
        chatAlreadyHasMessages: false,
        loopBusy: false,
        skipped: false,
      }),
    ).toBe(true)
  })

  it('ignores an empty history', () => {
    expect(
      shouldApplyRestoredChat({
        messageCount: 0,
        chatAlreadyHasMessages: false,
        loopBusy: false,
        skipped: false,
      }),
    ).toBe(false)
  })
})

describe('mapPersistedChat', () => {
  it('keeps role, text, and tool chips', () => {
    expect(
      mapPersistedChat([
        { role: 'user', text: 'hi' },
        { role: 'assistant', text: 'ok', tools: [{ name: 'web_search', summary: '검색', isError: false }] },
      ]),
    ).toEqual([
      { role: 'user', text: 'hi' },
      { role: 'assistant', text: 'ok', tools: [{ name: 'web_search', summary: '검색', isError: false }] },
    ])
  })
})
