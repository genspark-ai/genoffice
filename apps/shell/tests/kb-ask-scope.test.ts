// @vitest-environment jsdom
/**
 * The Ask panel's round scope, end to end through the real send path
 * (AgentLoop + IPC transport + knowledge-base skill) with a stubbed main
 * process: a pending query carrying scopePaths (the desktop's multi-select →
 * Ask AI) attaches the corpus chip, every ask while it is attached carries
 * `onlyPaths` on the search call, the chip surfaces the not-indexed hint from
 * the scoped answer, and clearing it — or leaving the thread — drops back to
 * the default corpus. A preload/main that predates onlyPaths answers the bare
 * array shape and must not take the panel down.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { KbAskPanel, shownCorpusCount } from '../src/renderer/src/KbAskPanel'
import { LocaleProvider } from '../src/renderer/src/locale'
import type { PendingAskQuery } from '../src/renderer/src/ask-intent'
import type { ChatMessage, ProjectApi } from '@genoffice/project-store'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const CORPUS = [
  { path: 'C:/kb/notes.md', folder: 'kb', name: 'notes.md', ext: 'md', sizeBytes: 512 },
]

const SCOPE = ['C:/kb/notes.md', 'C:/kb/unindexed.docx']

interface StartedRequest {
  requestId: string
  messages: Array<{ role: string; text?: string }>
  tools: Array<{ name: string }>
}

type ChunkListener = (chunk: {
  requestId: string
  type: string
  text?: string
  toolCall?: unknown
}) => void

/** the search calls that crossed the aiOffice bridge, query payloads verbatim */
let searchQueries: Array<{ q: string; limit?: number; onlyPaths?: readonly string[] }>
let searchAnswer: unknown
/** the read calls that crossed the bridge, (path, offset, onlyPaths) verbatim */
let reads: Array<[string, number, readonly string[] | undefined]>
let chunkListener: ChunkListener | null

beforeAll(() => {
  if (!Element.prototype.scrollTo) Element.prototype.scrollTo = () => {}
})

const projectApiStub = (): ProjectApi => {
  const chats = new Map<string, ChatMessage[]>([
    ['kb-ask', []],
    ['kb-ask-2', []],
  ])
  return {
    resolveChat: async (args: { filePath: string | null; tempChatId?: string }) => ({
      projectId: 'default',
      chatId: args.tempChatId ?? 'kb-ask',
    }),
    loadChat: async (args: { chatId: string }) => [...(chats.get(args.chatId) ?? [])],
    appendChat: async (args: { chatId: string; role: 'user' | 'assistant'; text: string }) => {
      const list = chats.get(args.chatId) ?? []
      list.push({
        projectId: 'default',
        chatId: args.chatId,
        seq: list.length,
        ts: '2026-10-30T09:00:00.000Z',
        role: args.role,
        text: args.text,
      })
      chats.set(args.chatId, list)
    },
    deleteChat: async () => {},
    listChats: async () =>
      [...chats.keys()].map((chatId, i) => ({
        chatId,
        updatedAt: `2026-10-30T09:0${i}:00.000Z`,
        approxCount: chats.get(chatId)?.length ?? 0,
        firstQuestion: chats.get(chatId)?.[0]?.text,
      })),
  } as unknown as ProjectApi
}

let host: HTMLDivElement
let root: Root
let started: StartedRequest[]
let pending: PendingAskQuery | null
let consumed: number
let project: ProjectApi
let requestSeq: number

beforeEach(() => {
  started = []
  pending = null
  consumed = 0
  requestSeq = 0
  searchQueries = []
  searchAnswer = { hits: [], notIndexed: [] }
  reads = []
  chunkListener = null
  project = projectApiStub()
  ;(window as unknown as { projectApi?: unknown }).projectApi = project
  ;(window as unknown as { aiOfficeTabs?: unknown }).aiOfficeTabs = {
    onChanged: () => () => {},
  }
  ;(window as unknown as { aiOffice: unknown }).aiOffice = {
    skillsList: vi.fn(async () => []),
    skillText: vi.fn(async () => ''),
    aiGetSettings: async () => null,
    aiGskStatus: async () => ({ loggedIn: true }),
    onAiSettingsChanged: () => () => {},
    kbList: async () => CORPUS,
    kbSearch: async (query: { q: string; limit?: number; onlyPaths?: readonly string[] }) => {
      searchQueries.push(query)
      return searchAnswer
    },
    kbRead: async (path: string, offset: number, onlyPaths?: readonly string[]) => {
      reads.push([path, offset, onlyPaths])
      return { text: 'body', totalChars: 4, offset }
    },
    kbRunCli: vi.fn(async () => ({ ok: true, code: 0, stdout: '{"status":"ok"}', stderr: '' })),
    appSkillManifest: async () => ({ windows: [], tools: [] }),
    appSkillExecute: vi.fn(async () => {
      throw new Error('nothing is open')
    }),
    onAiStream: (listener: ChunkListener) => {
      chunkListener = listener
      return () => {
        chunkListener = null
      }
    },
    // an alternating stub: each ask's first request gets a search tool call,
    // its follow-up (carrying the tool result) gets the answer text
    aiStream: (request: StartedRequest) => {
      started.push(request)
      const firstTurn = requestSeq++ % 2 === 0
      queueMicrotask(() => {
        if (firstTurn) {
          chunkListener?.({
            requestId: request.requestId,
            type: 'tool-call',
            toolCall: { id: 't1', name: 'search_knowledge_base', input: { query: 'notes' } },
          })
        } else {
          chunkListener?.({ requestId: request.requestId, type: 'delta', text: 'answer text' })
        }
        chunkListener?.({ requestId: request.requestId, type: 'done' })
      })
    },
    aiStreamCancel: () => {},
  }
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  delete (window as unknown as { aiOffice?: unknown }).aiOffice
  delete (window as unknown as { projectApi?: unknown }).projectApi
  delete (window as unknown as { aiOfficeTabs?: unknown }).aiOfficeTabs
})

function render(query: PendingAskQuery | null = null): void {
  pending = query
  act(() =>
    root.render(
      createElement(
        LocaleProvider,
        { initial: 'en' },
        createElement(KbAskPanel, {
          onOpenPath: () => {},
          onGoToRecents: () => {},
          onOpenModelSettings: () => {},
          pendingQuery: pending,
          onPendingQueryConsumed: () => {
            consumed += 1
          },
        }),
      ),
    ),
  )
}

/** re-render with a fresh pending query (the seq is what makes it consumable) */
const handoff = (query: PendingAskQuery): void => {
  act(() => {
    pending = query
    root.render(
      createElement(
        LocaleProvider,
        { initial: 'en' },
        createElement(KbAskPanel, {
          onOpenPath: () => {},
          onGoToRecents: () => {},
          onOpenModelSettings: () => {},
          pendingQuery: pending,
          onPendingQueryConsumed: () => {
            consumed += 1
          },
        }),
      ),
    )
  })
}

async function settle(): Promise<void> {
  await act(async () => {
    for (let i = 0; i < 12; i++) await Promise.resolve()
    await new Promise((resolve) => setTimeout(resolve, 0))
    for (let i = 0; i < 12; i++) await Promise.resolve()
  })
}

async function ask(question: string): Promise<void> {
  const ta = host.querySelector('textarea')!
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    setter.call(ta, question)
    ta.dispatchEvent(new Event('input', { bubbles: true }))
  })
  act(() => {
    ta.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
    )
  })
  await settle()
}

const click = (selector: string) => {
  const el = host.querySelector(selector)
  expect(el).toBeTruthy()
  act(() => (el as HTMLElement).click())
}

const chipText = () => host.querySelector('.kb-ask-scope-label')?.textContent ?? ''

describe('the round-scope chip', () => {
  it('a blank scoped handoff opens the panel with the chip attached and asks nothing', async () => {
    render({ text: '', seq: 1, origin: 'palette', scopePaths: SCOPE })
    await settle()
    expect(host.querySelector('.kb-ask-scope')).toBeTruthy()
    expect(chipText()).toBe('Corpus: 2 selected files')
    // the headbar's count follows the scope, not the index window
    expect(host.querySelector('.kb-ask-headbar span')?.textContent).toBe('Corpus: 2 selected files')
    // nothing sent itself: no model request, no search
    expect(started).toHaveLength(0)
    expect(searchQueries).toHaveLength(0)
    expect(consumed).toBe(1)
    // and the handoff is consumed once even though renders repeat
    await settle()
    expect(consumed).toBe(1)
  })

  it('a scoped question submits scoped: the search tool call carries onlyPaths', async () => {
    render({ text: 'what do these say', seq: 1, origin: 'palette', scopePaths: SCOPE })
    await settle()
    expect(host.querySelector('.kb-ask-scope')).toBeTruthy()
    // the run happened and its search crossed the bridge with the selection
    expect(started).toHaveLength(2)
    expect(searchQueries).toHaveLength(1)
    expect(searchQueries[0]).toMatchObject({ q: 'notes', onlyPaths: SCOPE })
    // the model's requests carried the knowledge tools
    expect(started[0]!.tools.map((t) => t.name)).toContain('search_knowledge_base')
  })

  it('the not-indexed hint surfaces from the scoped answer', async () => {
    searchAnswer = {
      hits: [{ path: 'C:/kb/notes.md', name: 'notes.md', ext: 'md', folder: 'kb', snippet: 'x' }],
      notIndexed: ['C:/kb/unindexed.docx'],
    }
    render({ text: 'what do these say', seq: 1, origin: 'palette', scopePaths: SCOPE })
    await settle()
    const miss = host.querySelector('.kb-ask-scope-miss')
    expect(miss).toBeTruthy()
    expect(miss?.textContent).toBe('1 not indexed')
    expect(miss?.getAttribute('title')).toBe('C:/kb/unindexed.docx')
  })

  it('clearing the chip drops back to the default corpus for the next ask', async () => {
    render({ text: '', seq: 1, origin: 'palette', scopePaths: SCOPE })
    await settle()
    click('.kb-ask-scope-clear')
    await settle()
    expect(host.querySelector('.kb-ask-scope')).toBeNull()
    // the headbar count is the index window's again
    expect(host.querySelector('.kb-ask-headbar span')?.textContent).toBe('Knowledge base: 1 files')
    await ask('anything')
    expect(searchQueries).toHaveLength(1)
    expect(searchQueries[0]).not.toHaveProperty('onlyPaths')
  })

  it('a preload that predates onlyPaths answers the bare array and nothing breaks', async () => {
    // the old main ignores the field and answers a plain hit array
    searchAnswer = [
      { path: 'C:/kb/notes.md', name: 'notes.md', ext: 'md', folder: 'kb', snippet: 'x' },
    ]
    render({ text: 'what do these say', seq: 1, origin: 'palette', scopePaths: SCOPE })
    await settle()
    // the answer completed and rendered; no not-indexed hint to fake
    expect(host.textContent).toContain('answer text')
    expect(host.querySelector('.kb-ask-scope-miss')).toBeNull()
    // the old kbRead signature (two args) simply ignores the scope
    expect(reads).toHaveLength(0)
  })

  it('switching threads leaves the selection behind', async () => {
    render({ text: '', seq: 1, origin: 'palette', scopePaths: SCOPE })
    await settle()
    expect(host.querySelector('.kb-ask-scope')).toBeTruthy()
    // the stored second thread shows as a chip row
    const threadButton = host.querySelector('.kb-ask-thread-chip > button')
    expect(threadButton).toBeTruthy()
    act(() => (threadButton as HTMLElement).click())
    await settle()
    expect(host.querySelector('.kb-ask-scope')).toBeNull()
  })

  it('a scope attached mid-conversation governs only the asks that follow', async () => {
    render()
    await settle()
    await ask('first, unscoped')
    expect(searchQueries[0]).not.toHaveProperty('onlyPaths')
    handoff({ text: '', seq: 7, origin: 'palette', scopePaths: ['C:/kb/notes.md'] })
    await settle()
    expect(chipText()).toBe('Corpus: 1 selected files')
    await ask('now scoped')
    expect(searchQueries[1]).toMatchObject({ onlyPaths: ['C:/kb/notes.md'] })
  })

  it('a scope parking keeps the draft the user is typing and hands the focus to the composer', async () => {
    render()
    await settle()
    const ta = host.querySelector('textarea')!
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
      setter.call(ta, 'my half-typed question')
      ta.dispatchEvent(new Event('input', { bubbles: true }))
    })
    // the desktop hands over a selection with no question: the chip parks,
    // the half-typed draft survives it (not blanked, not overwritten)
    handoff({ text: '', seq: 7, origin: 'palette', scopePaths: SCOPE })
    await settle()
    expect(chipText()).toBe('Corpus: 2 selected files')
    expect(ta.value).toBe('my half-typed question')
    expect(document.activeElement).toBe(ta)
    // nothing sent itself
    expect(started).toHaveLength(0)
    // the kept draft is the user's to finish and send scoped
    await ask('my half-typed question')
    expect(searchQueries[0]).toMatchObject({ onlyPaths: SCOPE })
  })
})

describe('shownCorpusCount', () => {
  it('the selection replaces the index count while attached', () => {
    expect(shownCorpusCount(['a', 'b'], 100)).toBe(2)
    expect(shownCorpusCount([], 100)).toBe(100)
    expect(shownCorpusCount([], null)).toBeNull()
    expect(shownCorpusCount(['a'], null)).toBe(1)
  })
})
