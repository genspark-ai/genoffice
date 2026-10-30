// The html panel's approval gate: risk:'high' source-editing tools
// (apply_ops, write_document) park at an inline allow/deny card before they
// execute. Denying must leave the document untouched and hand the model a
// do-not-retry result; allowing runs the tool. Driven end-to-end through the
// stubbed window.htmlApi stream, like ai-panel-stream-gating.test.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { AiPanel, type HtmlAiDeps } from '../src/renderer/ai/AiPanel'
import { TOOL_DENIED_OUTPUT } from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import type { HtmlApi } from '../src/shared/ipc'
import type { ParseMap } from '../src/renderer/document/parse-map'

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  if (!Element.prototype.scrollTo) Element.prototype.scrollTo = () => {}
})
const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup())
  vi.unstubAllGlobals()
})

type Chunk = { requestId: string; type: string; text?: string; toolCall?: unknown }
type Message = { role: string; results?: Array<{ output: string }> }

let streamListener: ((chunk: Chunk) => void) | null = null
let requests: Array<{ requestId: string; messages: Message[] }> = []
let nonceSeq = 0

const htmlApi = {
  getAiSettings: () => Promise.resolve({} as AiSettings),
  setAiSettings: () => Promise.resolve(),
  onAiSettingsChanged: () => () => {},
  aiGskStatus: () => Promise.resolve({ loggedIn: false }),
  openAiModelSettings: () => Promise.resolve(),
  openSkillsSettings: () => Promise.resolve(),
  skillsList: () => Promise.resolve([]),
  skillText: () => Promise.resolve(''),
  setAiPanelPrefs: () => Promise.resolve(),
  onAiPanelPrefsChanged: () => () => {},
  getPathForFile: () => null,
  readAttachmentImage: () => Promise.resolve({ ok: false, error: 'stub' }),
  aiStream: (request: { requestId: string; messages: Message[] }) => {
    requests.push(request)
    return Promise.resolve()
  },
  aiStreamCancel: () => {},
  onAiStream: (listener: (chunk: Chunk) => void) => {
    streamListener = listener
    return () => {}
  },
} as unknown as HtmlApi

window.htmlApi = htmlApi

const emptyMap: ParseMap = { version: 1, elements: [], bySid: new Map(), errorCount: 0 }

/** deps whose applyOps spy is the document-write assertion point */
function makeDeps(applyOps: ReturnType<typeof vi.fn>): HtmlAiDeps {
  return {
    access: {
      getText: () => '<!DOCTYPE html>\n<html><body><p>original body</p></body></html>\n',
      getVersion: () => 1,
      getMap: () => emptyMap,
      getLastManualVersion: () => 1,
      getFilePath: () => null,
      getSelectedSid: () => null,
      applyOps: applyOps as unknown as HtmlAiDeps['access']['applyOps'],
      replaceAll: () => {},
    },
    getSnapshot: () => ({ text: '' }),
    restoreSnapshot: () => {},
    onPrompt: () => {},
    onRunDone: () => {},
    clearHighlights: () => {},
    navigateTo: () => {},
    onBriefConfirmed: () => {},
    previewDraft: () => {},
  }
}

function mountPanel(deps: HtmlAiDeps): HTMLElement {
  streamListener = null
  requests = []
  const dock = document.createElement('div')
  dock.className = 'ai-dock'
  document.body.appendChild(dock)
  const root: Root = createRoot(dock)
  act(() =>
    root.render(
      createElement(AiPanel, {
        deps,
        filePath: null,
        preset: { text: 'set the title text', nonce: ++nonceSeq },
        editQueue: [],
        onQueueEditInstruction: () => {},
        onQueueRemove: () => {},
        onQueueClear: () => {},
        onQueueFocus: () => {},
        onQueueConsume: () => {},
        onCollapse: () => {},
      }),
    ),
  )
  cleanups.push(() => {
    act(() => root.unmount())
    dock.remove()
  })
  return dock
}

/** let the send pipeline (async pre-run reads) reach the first stream request */
async function settle(): Promise<void> {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

/** push one IPC chunk to the in-flight turn and let the loop react */
async function push(chunk: Omit<Chunk, 'requestId'>): Promise<void> {
  await act(async () => {
    const request = requests.at(-1)
    if (request && streamListener) streamListener({ requestId: request.requestId, ...chunk })
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

/** answer the current turn with an apply_ops call, then close it */
async function answerWithApplyOps(): Promise<void> {
  await push({
    type: 'tool-call',
    toolCall: {
      id: 't1',
      name: 'apply_ops',
      input: {
        ops: [{ op: 'set_text', sid: 3, text: 'AI edited' }],
        summary: 'set the paragraph text',
      },
    },
  })
  await push({ type: 'done' })
}

describe('AI panel tool confirmation', () => {
  it('parks a write tool at the allow/deny card; denying leaves the document untouched', async () => {
    const applyOps = vi.fn(() => ({ ok: true, ranges: [[1, 10]] as [number, number][] }))
    const dock = mountPanel(makeDeps(applyOps))
    await settle()

    await answerWithApplyOps()
    // the card is in the conversation and the loop holds at the gate
    expect(dock.querySelector('.ai-confirm-card')).not.toBeNull()
    expect(requests).toHaveLength(1)

    const deny = dock.querySelector<HTMLButtonElement>('.ai-confirm-deny')!
    await act(async () => {
      deny.click()
    })

    // the tool never ran: the document was never handed an op batch
    expect(applyOps).not.toHaveBeenCalled()
    // the model received the denial as an error tool result it can act on
    expect(requests).toHaveLength(2)
    const toolMessage = requests[1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).toBe(TOOL_DENIED_OUTPUT)
    // the settled card stays in the flow as the record of the decision
    expect(dock.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe('denied')
    expect(dock.querySelector('.ai-confirm-deny')).toBeNull()
  })

  it('allowing executes the write tool', async () => {
    const applyOps = vi.fn(() => ({ ok: true, ranges: [[1, 10]] as [number, number][] }))
    const dock = mountPanel(makeDeps(applyOps))
    await settle()

    await answerWithApplyOps()
    const allow = dock.querySelector<HTMLButtonElement>('.ai-confirm-allow')!
    await act(async () => {
      allow.click()
    })

    // the write happened through the real access surface
    expect(applyOps).toHaveBeenCalledTimes(1)
    expect(applyOps.mock.calls[0]![0]).toEqual([{ op: 'set_text', sid: 3, text: 'AI edited' }])
    // the tool result went back to the model (its round-trip continued)
    expect(requests).toHaveLength(2)
    const toolMessage = requests[1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).not.toBe(TOOL_DENIED_OUTPUT)
    expect(dock.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe('allowed')
  })

  it('a read tool never asks: it runs straight through', async () => {
    const applyOps = vi.fn(() => ({ ok: true, ranges: [] as [number, number][] }))
    const dock = mountPanel(makeDeps(applyOps))
    await settle()

    await push({
      type: 'tool-call',
      toolCall: { id: 't1', name: 'read_source', input: { start_line: 1, end_line: 5 } },
    })
    await push({ type: 'done' })
    await settle()
    expect(dock.querySelector('.ai-confirm-card')).toBeNull()
    expect(applyOps).not.toHaveBeenCalled()
  })

  it('unmounting the panel denies the pending request', async () => {
    const applyOps = vi.fn(() => ({ ok: true, ranges: [] as [number, number][] }))
    const dock = mountPanel(makeDeps(applyOps))
    await settle()

    await answerWithApplyOps()
    expect(dock.querySelector('.ai-confirm-card')).not.toBeNull()
    // the first registered cleanup unmounts the panel
    cleanups.splice(0)[0]!()
    await settle()
    expect(applyOps).not.toHaveBeenCalled()
    expect(requests).toHaveLength(1)
  })
})
