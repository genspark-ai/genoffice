// @vitest-environment jsdom
/**
 * The Ask panel's safety layer end to end: the inline allow/deny card around
 * risk:'high' tools, and the background-task list backed by the project
 * store's first-class runs API (appendRunEvent / listRuns / deleteRun over the
 * renderer bridges). These drive the real send path (AgentLoop + IPC
 * transport) with a stubbed main process, asserting on what the model receives
 * and on the run events that cross the runs bridge.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { KbAskPanel } from '../src/renderer/src/KbAskPanel'
import { LocaleProvider } from '../src/renderer/src/locale'
import { TOOL_DENIED_OUTPUT } from '@genoffice/agent-core'
import type { ChatMessage, ProjectApi, RunSummary } from '@genoffice/project-store'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const CORPUS = [
  { path: 'C:/kb/notes.md', folder: 'kb', name: 'notes.md', ext: 'md', sizeBytes: 512 },
]

interface StartedRequest {
  requestId: string
  messages: Array<{ role: string; text?: string; results?: Array<{ output: string }> }>
  tools: Array<{ name: string; risk?: string }>
}

type ChunkListener = (chunk: {
  requestId: string
  type: string
  text?: string
  toolCall?: unknown
}) => void

/** stores the conversations and records every call, like the main process would */
function projectApiStub(storedRuns: string[]) {
  const calls: Array<{ method: string; args: Record<string, unknown> }> = []
  const chats = new Map<string, ChatMessage[]>()
  for (const text of storedRuns) {
    const list = chats.get(RUNS_CHAT_ID) ?? []
    list.push({
      seq: list.length,
      ts: `2026-10-30T09:0${list.length}:00.000Z`,
      role: 'assistant',
      text,
    })
    chats.set(RUNS_CHAT_ID, list)
  }
  const api = {
    resolveChat: async (args: { filePath: string | null; tempChatId?: string }) => {
      calls.push({ method: 'resolveChat', args: args as Record<string, unknown> })
      return { projectId: 'default', chatId: args.tempChatId ?? 'kb-ask' }
    },
    appendChat: async (args: Record<string, unknown>) => {
      calls.push({ method: 'appendChat', args })
      const chatId = args.chatId as string
      const list = chats.get(chatId) ?? []
      list.push({
        seq: list.length,
        ts: '2026-10-30T09:00:00.000Z',
        role: args.role as 'user' | 'assistant',
        text: args.text as string,
      })
      chats.set(chatId, list)
    },
    loadChat: async (args: { chatId: string; limit?: number }) => {
      calls.push({ method: 'loadChat', args: args as Record<string, unknown> })
      return [...(chats.get(args.chatId) ?? [])]
    },
    deleteChat: async (args: Record<string, unknown>) => {
      calls.push({ method: 'deleteChat', args })
      chats.delete(args.chatId as string)
    },
  } as unknown as ProjectApi
  return { api, calls, chats }
}

let host: HTMLDivElement
let root: Root
let started: StartedRequest[]
let chunkListener: ChunkListener | null
let runCli: ReturnType<typeof vi.fn>
let project: ReturnType<typeof projectApiStub>
let runs: ReturnType<typeof runsApiStub>

/**
 * The runs half of the aiOffice bridge, over an in-memory event log: the
 * panel sees the same contract the shell preload forwards
 * (project:appendRunEvent / listRuns / deleteRun), with the store's fold
 * reduced to what these tests need (a start fixes the row, the last terminal
 * settles it).
 */
function runsApiStub(storedRuns: RunSummary[]) {
  const calls: Array<{ method: string; args: Record<string, unknown> }> = []
  const eventsByRun = new Map<string, Array<Record<string, unknown>>>()
  const summaries = [...storedRuns]
  const fold = (): RunSummary[] => {
    const folded: RunSummary[] = []
    for (const [runId, events] of eventsByRun) {
      const start = events.find((e) => e.kind === 'start')
      if (!start) continue
      const terminal = [...events].reverse().find((e) => e.kind === 'finish' || e.kind === 'error')
      folded.push({
        runId,
        instruction: String(start.instruction ?? ''),
        ...(start.chatId ? { chatId: String(start.chatId) } : {}),
        status: terminal
          ? terminal.kind === 'error'
            ? 'error'
            : terminal.outcome === 'cancelled'
              ? 'cancelled'
              : terminal.outcome === 'discarded'
                ? 'discarded'
                : 'done'
          : 'interrupted',
        startedAt: '2026-10-30T09:00:00.000Z',
      })
    }
    return folded
  }
  const api = {
    appendRunEvent: async (args: Record<string, unknown>) => {
      calls.push({ method: 'appendRunEvent', args })
      const runId = String(args.runId)
      const list = eventsByRun.get(runId) ?? []
      list.push({ ...args })
      eventsByRun.set(runId, list)
    },
    listRuns: async (args: Record<string, unknown>) => {
      calls.push({ method: 'listRuns', args })
      return [...summaries, ...fold()]
    },
    deleteRun: async (args: Record<string, unknown>) => {
      calls.push({ method: 'deleteRun', args })
      eventsByRun.delete(String(args.runId))
      const at = summaries.findIndex((run) => run.runId === args.runId)
      if (at >= 0) summaries.splice(at, 1)
    },
  }
  return { api, calls, summaries }
}

beforeAll(() => {
  if (!Element.prototype.scrollTo) Element.prototype.scrollTo = () => {}
})

beforeEach(() => {
  started = []
  chunkListener = null
  runCli = vi.fn(async () => ({ ok: true, code: 0, stdout: '{"status":"ok"}', stderr: '' }))
  project = projectApiStub([])
  runs = runsApiStub([])
  ;(window as unknown as { projectApi?: unknown }).projectApi = project.api
  // the tab-strip bridge the desktop loop subscribes to for manifest refreshes
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
    kbRunCli: runCli,
    // the desktop loop's app-skill proxy: nothing open, so no cross-app tools
    appSkillManifest: async () => ({ windows: [], tools: [] }),
    appSkillExecute: vi.fn(async () => {
      throw new Error('nothing is open')
    }),
    appendRunEvent: runs.api.appendRunEvent,
    listRuns: runs.api.listRuns,
    deleteRun: runs.api.deleteRun,
    onAiStream: (listener: ChunkListener) => {
      chunkListener = listener
      return () => {
        chunkListener = null
      }
    },
    aiStream: (request: StartedRequest) => {
      started.push(request)
      queueMicrotask(() => chunkListener?.({ requestId: request.requestId, type: 'done' }))
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

function render(): void {
  act(() =>
    root.render(
      createElement(
        LocaleProvider,
        { initial: 'en' },
        createElement(KbAskPanel, {
          onOpenPath: () => {},
          onGoToRecents: () => {},
          onOpenModelSettings: () => {},
        }),
      ),
    ),
  )
}

async function settle(): Promise<void> {
  await act(async () => {
    for (let i = 0; i < 12; i++) await Promise.resolve()
    await new Promise((resolve) => setTimeout(resolve, 0))
    for (let i = 0; i < 12; i++) await Promise.resolve()
  })
}

function type(text: string): void {
  const ta = host.querySelector('textarea')!
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    setter.call(ta, text)
    ta.setSelectionRange(text.length, text.length)
    ta.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function press(key: string): void {
  act(() => {
    host
      .querySelector('textarea')!
      .dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
  })
}

async function ask(question: string): Promise<void> {
  type(question)
  press('Enter')
  await settle()
}

const click = (selector: string) => {
  const el = host.querySelector(selector)
  expect(el).toBeTruthy()
  act(() => (el as HTMLElement).click())
}

/** the run events that crossed the runs bridge, as (method, args) pairs */
const runEventCalls = (): Array<Record<string, unknown>> =>
  runs.calls.filter((c) => c.method === 'appendRunEvent').map((c) => c.args)

describe('KbAskPanel tool confirmation', () => {
  it('offers risk-marked tools to the model', () => {
    render()
    return settle().then(() => {
      const runGenoffice = started.length === 0 ? null : null // no request yet; the tools ride the first one
      void runGenoffice
      // the corpus gate is open; ask once to see the request's tool table
      return ask('build a file').then(() => {
        const tools = started[0]!.tools
        const risky = tools.filter((t) => t.risk === 'high').map((t) => t.name)
        expect(risky).toContain('run_genoffice')
        expect(risky).toContain('create_document')
      })
    })
  })

  it('shows the allow/deny card and holds the next request until decided', async () => {
    render()
    await settle()
    // turn 1: the model calls the CLI; the run parks at the approval gate
    ;(window as unknown as { aiOffice: { aiStream: unknown } }).aiOffice.aiStream = (
      request: StartedRequest,
    ) => {
      started.push(request)
      const turn = started.length
      queueMicrotask(() => {
        if (turn === 1) {
          chunkListener?.({
            requestId: request.requestId,
            type: 'tool-call',
            toolCall: { id: 't1', name: 'run_genoffice', input: { args: ['guide', 'docs'] } },
          })
        } else {
          chunkListener?.({ requestId: request.requestId, type: 'delta', text: 'ok then' })
        }
        chunkListener?.({ requestId: request.requestId, type: 'done' })
      })
    }

    await ask('build a guide')
    expect(started).toHaveLength(1)
    // the card is in the conversation, awaiting the decision
    expect(host.querySelector('.ai-confirm-card')).toBeTruthy()
    expect(host.querySelector('.ai-confirm-allow')).toBeTruthy()
    // no second request: the loop is parked at the gate

    click('.ai-confirm-deny')
    await settle()
    expect(started).toHaveLength(2)
    // the model received the denial as an error tool result it can act on
    const toolMessage = started[1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).toBe(TOOL_DENIED_OUTPUT)
    // the settled card stays in the flow as the record of the decision
    expect(host.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe('denied')
    expect(host.querySelector('.ai-confirm-deny')).toBeNull()
    // the CLI was never spawned
    expect(runCli).not.toHaveBeenCalled()
  })

  it('allowing executes the tool through the host bridge', async () => {
    render()
    await settle()
    ;(window as unknown as { aiOffice: { aiStream: unknown } }).aiOffice.aiStream = (
      request: StartedRequest,
    ) => {
      started.push(request)
      const turn = started.length
      queueMicrotask(() => {
        if (turn === 1) {
          chunkListener?.({
            requestId: request.requestId,
            type: 'tool-call',
            toolCall: { id: 't1', name: 'run_genoffice', input: { args: ['guide', 'docs'] } },
          })
        } else {
          chunkListener?.({ requestId: request.requestId, type: 'delta', text: 'ok then' })
        }
        chunkListener?.({ requestId: request.requestId, type: 'done' })
      })
    }

    await ask('build a guide')
    click('.ai-confirm-allow')
    await settle()
    expect(runCli).toHaveBeenCalledTimes(1)
    // the model round-trip continued: the tool result went back to it
    expect(started).toHaveLength(2)
    const toolMessage = started[1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).toContain('exit 0')
    expect(host.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe('allowed')
  })

  it('Stop settles an open confirmation as denied', async () => {
    render()
    await settle()
    ;(window as unknown as { aiOffice: { aiStream: unknown } }).aiOffice.aiStream = (
      request: StartedRequest,
    ) => {
      started.push(request)
      const turn = started.length
      queueMicrotask(() => {
        if (turn === 1) {
          chunkListener?.({
            requestId: request.requestId,
            type: 'tool-call',
            toolCall: { id: 't1', name: 'run_genoffice', input: { args: ['guide', 'docs'] } },
          })
        } else {
          chunkListener?.({ requestId: request.requestId, type: 'delta', text: 'ok then' })
        }
        chunkListener?.({ requestId: request.requestId, type: 'done' })
      })
    }

    await ask('build a guide')
    expect(host.querySelector('.ai-confirm-card')).toBeTruthy()
    // the composer's Stop button while the card is open
    click('.ai-stop-btn')
    await settle()
    // the request settled as denied, the loop freed, no second turn
    expect(host.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe('denied')
    expect(runCli).not.toHaveBeenCalled()
  })
})

describe('KbAskPanel background tasks', () => {
  it('lists a run while it streams and settles its events when it finishes', async () => {
    render()
    await settle()
    expect(host.querySelector('.kb-ask-tasks')).toBeNull() // nothing to show yet

    await ask('summarize the notes')
    // the run completed within the same settle: the row came and went
    expect(host.querySelector('.kb-ask-task')).toBeNull()
    const events = runEventCalls()
    expect(events).toHaveLength(2)
    expect(events[0]).toMatchObject({
      projectId: 'default',
      kind: 'start',
      instruction: 'summarize the notes',
      chatId: 'kb-ask',
    })
    expect(events[1]).toMatchObject({ runId: events[0]!.runId, kind: 'finish', outcome: 'done' })
  })

  it('shows the run left unfinished by a closed window as interrupted, and dismissable', async () => {
    runs = runsApiStub([
      {
        runId: 'run-9',
        chatId: 'kb-ask',
        instruction: 'build the deck',
        status: 'interrupted',
        startedAt: '2026-10-29T18:00:00.000Z',
      },
    ])
    ;(window as unknown as { aiOffice: Record<string, unknown> }).aiOffice.appendRunEvent =
      runs.api.appendRunEvent
    ;(window as unknown as { aiOffice: Record<string, unknown> }).aiOffice.listRuns =
      runs.api.listRuns
    ;(window as unknown as { aiOffice: Record<string, unknown> }).aiOffice.deleteRun =
      runs.api.deleteRun
    render()
    await settle()

    const row = host.querySelector('.kb-ask-task[data-status="interrupted"]')
    expect(row).toBeTruthy()
    expect(row?.querySelector('.kb-ask-task-text')?.textContent).toBe('build the deck')

    click('.kb-ask-task[data-status="interrupted"] .kb-ask-task-stop')
    await settle()
    // the dismissal settled the record as discarded and dropped the row
    expect(host.querySelector('.kb-ask-task[data-status="interrupted"]')).toBeNull()
    const discard = runEventCalls().at(-1)
    expect(discard).toMatchObject({
      runId: 'run-9',
      kind: 'finish',
      outcome: 'discarded',
    })
  })

  it('clear all empties the interrupted list through the store', async () => {
    runs = runsApiStub([
      {
        runId: 'run-9',
        instruction: 'first interrupted',
        status: 'interrupted',
        startedAt: '2026-10-29T18:00:00.000Z',
      },
      {
        runId: 'run-8',
        instruction: 'second interrupted',
        status: 'interrupted',
        startedAt: '2026-10-29T17:00:00.000Z',
      },
    ])
    ;(window as unknown as { aiOffice: Record<string, unknown> }).aiOffice.appendRunEvent =
      runs.api.appendRunEvent
    ;(window as unknown as { aiOffice: Record<string, unknown> }).aiOffice.listRuns =
      runs.api.listRuns
    ;(window as unknown as { aiOffice: Record<string, unknown> }).aiOffice.deleteRun =
      runs.api.deleteRun
    render()
    await settle()
    expect(host.querySelectorAll('.kb-ask-task[data-status="interrupted"]')).toHaveLength(2)

    click('.kb-ask-task-clear')
    await settle()
    expect(host.querySelector('.kb-ask-tasks')).toBeNull()
    // the wipe went through the runs API: every stored run deleted by id
    const deleted = runs.calls
      .filter((c) => c.method === 'deleteRun')
      .map((c) => (c.args as { runId: string }).runId)
    expect(deleted.sort()).toEqual(['run-8', 'run-9'])
  })
})
