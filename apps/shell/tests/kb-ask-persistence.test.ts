import { describe, expect, it } from 'vitest'

import type { ChatMessage, ProjectApi } from '@genoffice/project-store'

import { claimThread, restoreAskMessages, type AskMessage } from '../src/renderer/src/KbAskPanel'

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
