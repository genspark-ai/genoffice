import { describe, expect, it } from 'vitest'

import type { ChatMessage, ChatMeta, ProjectApi } from '@genoffice/project-store'

import {
  claimThread,
  firstQuestionOf,
  listThreads,
  restoreAskMessages,
  type AskMessage,
} from '../src/renderer/src/KbAskPanel'

function stored(over: Partial<ChatMessage> = {}): ChatMessage {
  return {
    projectId: 'default',
    chatId: 'kb-ask',
    role: 'user',
    text: 'who is in the photograph',
    ts: '2026-10-16T09:00:00.000Z',
    seq: 0,
    ...over,
  }
}

describe('restoring a stored knowledge-base thread', () => {
  it('keeps the turns in order with their role and text', () => {
    const restored = restoreAskMessages([
      stored({ role: 'user', text: 'first question', seq: 0 }),
      stored({ role: 'assistant', text: 'first answer', seq: 1 }),
      stored({ role: 'user', text: 'second question', seq: 2 }),
    ])
    expect(restored).toEqual([
      { role: 'user', text: 'first question', toolNotes: [] },
      { role: 'assistant', text: 'first answer', toolNotes: [] },
      { role: 'user', text: 'second question', toolNotes: [] },
    ])
  })

  it('drops a role the panel cannot render', () => {
    // The store types role as user|assistant, but the JSONL on disk is not
    // ours to trust: an unreadable turn is skipped, not shown as a blank one.
    const restored = restoreAskMessages([
      stored({ role: 'user', text: 'kept', seq: 0 }),
      stored({ role: 'system', text: 'dropped', seq: 1 } as Partial<ChatMessage>),
    ])
    expect(restored).toEqual([{ role: 'user', text: 'kept', toolNotes: [] }])
  })

  it('restores no badges, because a stored turn has none', () => {
    // The badges, token chip and save outcome are this session's state; a
    // turn read back from disk has none of them, and inventing them would
    // show a search or a saved file that did not happen here.
    const restored = restoreAskMessages([
      stored({
        role: 'assistant',
        text: 'an answer',
        seq: 0,
        tools: [{ name: 'kb_search', summary: 'searched 3 files' }],
      }),
    ])
    expect(restored[0]?.toolNotes).toEqual([])
    expect(Object.keys(restored[0] ?? {}).sort()).toEqual(['role', 'text', 'toolNotes'])
  })

  it('is empty for an empty transcript', () => {
    expect(restoreAskMessages([])).toEqual([])
  })
})

/** A project-store that answers, and records what it was asked. */
function recorder(storedMessages: ChatMessage[] = []) {
  const calls: Array<{ method: string; args: unknown }> = []
  const api = {
    resolveChat: async (args: unknown) => {
      calls.push({ method: 'resolveChat', args })
      return { projectId: 'default', chatId: 'kb-ask' }
    },
    appendChat: async (args: unknown) => {
      calls.push({ method: 'appendChat', args })
    },
    loadChat: async (args: unknown) => {
      calls.push({ method: 'loadChat', args })
      return storedMessages
    },
    rebindChat: async (args: unknown) => {
      calls.push({ method: 'rebindChat', args })
      return { projectId: 'default', chatId: 'kb-ask' }
    },
  } as unknown as ProjectApi
  return { api, calls }
}

describe('claiming the panel thread', () => {
  it('asks for a file-less chat under the fixed id, not a fresh one', async () => {
    // The store takes a file-less caller's id verbatim, so a constant is what
    // makes the next launch find this thread again.
    const { api, calls } = recorder([stored({ role: 'user', text: 'older question' })])
    const shown: AskMessage[][] = []
    await claimThread(
      api,
      'kb-ask',
      [],
      (m) => shown.push(m),
      () => true,
    )

    expect(calls[0]).toEqual({
      method: 'resolveChat',
      args: { filePath: null, tempChatId: 'kb-ask' },
    })
    expect(shown[0]).toEqual([{ role: 'user', text: 'older question', toolNotes: [] }])
  })

  it('writes the turns asked before the ids landed, before reading back', async () => {
    const { api, calls } = recorder([])
    await claimThread(
      api,
      'kb-ask',
      [
        { role: 'user', text: 'first' },
        { role: 'assistant', text: 'first answer' },
      ],
      () => undefined,
      () => true,
    )
    expect(calls.map((c) => c.method)).toEqual([
      'resolveChat',
      'appendChat',
      'appendChat',
      'loadChat',
    ])
    expect(calls[1]?.args).toEqual({
      projectId: 'default',
      chatId: 'kb-ask',
      role: 'user',
      text: 'first',
    })
  })

  it('shows nothing when the panel was torn down mid-flight', async () => {
    const { api } = recorder([stored({ text: 'too late to show' })])
    const shown: AskMessage[][] = []
    await claimThread(
      api,
      'kb-ask',
      [],
      (m) => shown.push(m),
      () => false,
    )
    expect(shown).toEqual([])
  })

  it('returns the ids so the panel can write the turns that come next', async () => {
    const { api } = recorder([])
    await expect(
      claimThread(
        api,
        'kb-ask',
        [],
        () => undefined,
        () => true,
      ),
    ).resolves.toEqual({
      projectId: 'default',
      chatId: 'kb-ask',
    })
  })
})

describe('listing the stored threads', () => {
  const meta = (chatId: string, updatedAt: string): ChatMeta => ({
    chatId,
    updatedAt,
    approxCount: 4,
  })

  it('puts the most recent conversation on top', () => {
    const listed = listThreads(
      [meta('old', '2026-10-15T09:00:00.000Z'), meta('new', '2026-10-17T09:00:00.000Z')],
      () => undefined,
    )
    expect(listed.map((t) => t.chatId)).toEqual(['new', 'old'])
  })

  it('labels a conversation by its own first question', () => {
    const first: Record<string, string> = {
      a: 'who is in the photograph',
      b: 'what did the letter say',
    }
    const listed = listThreads(
      [meta('a', '2026-10-17T09:00:00.000Z'), meta('b', '2026-10-16T09:00:00.000Z')],
      (chatId) => first[chatId],
    )
    expect(listed.map((t) => t.label)).toEqual([
      'who is in the photograph',
      'what did the letter say',
    ])
  })

  it('collapses a question that wrapped and truncates a long one', () => {
    expect(
      listThreads([meta('a', '2026-10-17T09:00:00.000Z')], () => 'a\n\n  b   c')[0]?.label,
    ).toBe('a b c')
    const long = 'x'.repeat(200)
    expect(listThreads([meta('a', '2026-10-17T09:00:00.000Z')], () => long)[0]?.label).toHaveLength(
      60,
    )
  })

  it('falls back to the id when a conversation has no question to name it by', () => {
    // An empty thread, or one whose read failed: the id is unique and
    // copyable, which a blank row would not be.
    expect(
      listThreads([meta('kb-ask-1731', '2026-10-17T09:00:00.000Z')], () => undefined)[0]?.label,
    ).toBe('kb-ask-1731')
    expect(
      listThreads([meta('kb-ask-1731', '2026-10-17T09:00:00.000Z')], () => '   ')[0]?.label,
    ).toBe('kb-ask-1731')
  })

  it('is empty when nothing was ever asked', () => {
    expect(listThreads([], () => 'anything')).toEqual([])
  })
})

describe('naming a conversation by its first question', () => {
  it('takes the first user turn, not the first turn of any kind', () => {
    expect(
      firstQuestionOf([
        stored({ role: 'assistant', text: 'an answer that came first', seq: 0 }),
        stored({ role: 'user', text: 'the actual question', seq: 1 }),
      ]),
    ).toBe('the actual question')
  })

  it('is undefined when the thread opens with answers', () => {
    expect(
      firstQuestionOf([stored({ role: 'assistant', text: 'orphan answer', seq: 0 })]),
    ).toBeUndefined()
    expect(firstQuestionOf([])).toBeUndefined()
  })
})
