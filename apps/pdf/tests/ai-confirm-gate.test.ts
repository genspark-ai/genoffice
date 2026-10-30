// The pdf panel's approval gate: risk:'high' document-editing tools (see the
// policy note in ai/tools.ts — set_watermark, edit_text, apply_ops, the image
// ops, create_document, …) park at an inline allow/deny card before they
// execute. Denying must leave the document untouched and hand the model a
// do-not-retry result; allowing runs the tool. The file-level operations
// (replace_pages & co.) keep their own confirmFileOp card and are not gated a
// second time. Runs are driven through the real transport chain with a
// stubbed window.pdfApi stream channel, like ai-panel-prefs.test.ts.
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { applyAiPanelPrefs, DEFAULT_AI_PANEL_PREFS } from '@genoffice/ui'
import { TOOL_DENIED_OUTPUT } from '@genoffice/agent-core'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import type { PdfAppDeps } from '../src/renderer/ai/tools'
import type { StampConfig } from '../src/renderer/edit-state'

type StreamChunk = {
  requestId: string
  type: 'delta' | 'reasoning' | 'tool-call' | 'usage' | 'done' | 'error' | 'ping'
  text?: string
  stopReason?: string
  toolCall?: { id: string; name: string; input: Record<string, unknown> }
}

type Message = { role: string; results?: Array<{ output: string }> }
type StreamListener = (chunk: StreamChunk) => void

const streamListeners = new Set<StreamListener>()
let startedRequestId = ''
/** every aiStream request with its message list, for the denial assertions */
const requests: Array<{ requestId: string; messages: Message[] }> = []

const pdfApi = {
  getAiSettings: () => Promise.resolve({}),
  setAiSettings: () => Promise.resolve(),
  onAiSettingsChanged: () => () => {},
  openAiModelSettings: () => Promise.resolve(),
  openSkillsSettings: () => Promise.resolve(),
  skillsList: () => Promise.resolve([]),
  skillText: () => Promise.resolve(''),
  gskStatus: () => Promise.resolve({ loggedIn: false }),
  aiStream: (request: { requestId: string; messages: Message[] }) => {
    startedRequestId = request.requestId
    requests.push(request)
    return Promise.resolve()
  },
  aiStreamCancel: () => Promise.resolve(),
  onAiStream: (handler: StreamListener) => {
    streamListeners.add(handler)
    return () => streamListeners.delete(handler)
  },
  setAiPanelPrefs: () => Promise.resolve(DEFAULT_AI_PANEL_PREFS),
  listPageImages: () => Promise.resolve([]),
  pageImageBytes: () => Promise.resolve(null),
  openSourcePath: () => Promise.resolve(),
}

const projectApi = {
  resolveChat: () => Promise.resolve({ projectId: 'p', chatId: 'c' }),
  appendChat: () => Promise.resolve(),
  loadChat: () => Promise.resolve([]),
  rebindChat: () => Promise.resolve(null),
}

/** the stamp surface set_watermark writes through; the spy is the assertion point */
const setStamps = vi.fn((cfg: StampConfig | null) => {
  void cfg
})

const api = {
  doc: () => null,
  fileName: () => 'sample.pdf',
  pageCount: () => 1,
  currentPage: () => 1,
  readOnly: () => false,
  ocrText: () => null,
  selection: () => null,
  pendingSummary: () => '',
  outline: () => null,
  searchIndex: () => null,
  isDeleted: () => false,
  gotoPage: () => true,
  addMarkup: () => {},
  annotationsOn: () => Promise.resolve({ threads: [], markups: [] }),
  annotationSummary: () => '',
  metadata: () => ({}),
  addNote: () => Promise.resolve(null),
  deletePages: () => Promise.resolve({ ok: true, pageCount: 1 }),
  rotatePages: () => Promise.resolve({ ok: true, pageCount: 1 }),
  setPageSize: () => Promise.resolve({ ok: true, pageCount: 1 }),
  cropPages: () => Promise.resolve({ ok: true, pageCount: 1 }),
  replacePages: () => Promise.resolve({ ok: true, pageCount: 1 }),
  extractPages: () => Promise.resolve({ ok: true, savedPath: '' }),
  splitPdf: () => Promise.resolve({ ok: true, savedDir: '', count: 0 }),
  splitPages: () => Promise.resolve({ ok: true, savedPath: '' }),
  mergePages: () => Promise.resolve({ ok: true, savedPath: '' }),
  stamps: () => null,
  setStamps,
} as unknown as PdfAppDeps

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  Object.defineProperty(window, 'projectApi', { value: projectApi, configurable: true })
  Object.defineProperty(window, 'pdfApi', { value: pdfApi, configurable: true })
  Element.prototype.scrollTo = () => {}
  // jsdom's crypto may lack randomUUID (the IPC transport mints stream ids)
  const cryptoObj = globalThis.crypto as Crypto & { randomUUID?: () => string }
  if (typeof cryptoObj.randomUUID !== 'function') {
    Object.defineProperty(cryptoObj, 'randomUUID', {
      value: () => `test-${Math.random().toString(36).slice(2)}`,
      configurable: true,
    })
  }
})

beforeEach(() => {
  streamListeners.clear()
  setStamps.mockClear()
  requests.length = 0
  startedRequestId = ''
  applyAiPanelPrefs(DEFAULT_AI_PANEL_PREFS)
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
})

const flush = async (): Promise<void> => {
  for (let i = 0; i < 6; i++) await Promise.resolve()
}

async function renderPanel(): Promise<HTMLDivElement> {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(
      createElement(AiPanel, {
        api,
        filePath: 'sample.pdf',
        onCollapse: () => {},
        preset: null,
      }),
    )
    await flush()
  })
  await act(async () => {
    await flush()
  })
  return container
}

/** send the first quick-action prompt (starts a real run through the stubbed stream) */
const sendQuick = async (): Promise<void> => {
  const btn = container!.querySelector<HTMLButtonElement>('.ai-quick-btn')
  expect(btn, 'quick action button on the empty chat').not.toBeNull()
  await act(async () => {
    btn!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    await flush()
  })
  expect(startedRequestId, 'the transport started a stream').not.toBe('')
}

const emit = async (chunk: Omit<StreamChunk, 'requestId'>): Promise<void> => {
  await act(async () => {
    const full = { ...chunk, requestId: startedRequestId }
    for (const listener of [...streamListeners]) listener(full)
    await flush()
  })
}

/** answer the current turn with a set_watermark call, then close it */
const answerWithWatermark = async (): Promise<void> => {
  await emit({
    type: 'tool-call',
    toolCall: { id: 't1', name: 'set_watermark', input: { text: 'CONFIDENTIAL' } },
  })
  await emit({ type: 'done', stopReason: 'end_turn' })
}

describe('AI panel tool confirmation', () => {
  it('parks a write tool at the allow/deny card; denying leaves the document untouched', async () => {
    await renderPanel()
    await sendQuick()
    await answerWithWatermark()

    // the card is in the conversation and the loop holds at the gate
    expect(container!.querySelector('.ai-confirm-card')).not.toBeNull()
    expect(requests).toHaveLength(1)

    const deny = container!.querySelector<HTMLButtonElement>('.ai-confirm-deny')!
    await act(async () => {
      deny.click()
      await flush()
    })

    // the tool never ran: no stamp configuration was queued
    expect(setStamps).not.toHaveBeenCalled()
    // the model received the denial as an error tool result it can act on
    expect(requests).toHaveLength(2)
    const toolMessage = requests[1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).toBe(TOOL_DENIED_OUTPUT)
    // the settled card stays in the flow as the record of the decision
    expect(container!.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe(
      'denied',
    )
    expect(container!.querySelector('.ai-confirm-deny')).toBeNull()
  })

  it('allowing executes the write tool', async () => {
    await renderPanel()
    await sendQuick()
    await answerWithWatermark()

    const allow = container!.querySelector<HTMLButtonElement>('.ai-confirm-allow')!
    await act(async () => {
      allow.click()
      await flush()
    })

    // the write happened through the real stamp surface
    expect(setStamps).toHaveBeenCalledTimes(1)
    expect((setStamps.mock.calls[0]![0] as { wm: { text: string } }).wm.text).toBe('CONFIDENTIAL')
    // the tool result went back to the model (its round-trip continued)
    expect(requests).toHaveLength(2)
    const toolMessage = requests[1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).not.toBe(TOOL_DENIED_OUTPUT)
    expect(container!.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe(
      'allowed',
    )
  })

  it('a read tool never asks: it runs straight through', async () => {
    await renderPanel()
    await sendQuick()
    await emit({
      type: 'tool-call',
      toolCall: { id: 't1', name: 'read_pages', input: { start: 1 } },
    })
    await emit({ type: 'done', stopReason: 'end_turn' })
    await act(async () => {
      await flush()
    })
    expect(container!.querySelector('.ai-confirm-card')).toBeNull()
    expect(setStamps).not.toHaveBeenCalled()
  })

  it('unmounting the panel denies the pending request', async () => {
    await renderPanel()
    await sendQuick()
    await answerWithWatermark()
    expect(container!.querySelector('.ai-confirm-card')).not.toBeNull()

    await act(async () => {
      root!.unmount()
      await flush()
    })
    root = null
    // the tool never ran, and the denial still reaches the model as a
    // do-not-retry result the next turn can act on
    expect(setStamps).not.toHaveBeenCalled()
    expect(requests.length).toBeGreaterThanOrEqual(2)
    const toolMessage = requests[1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).toBe(TOOL_DENIED_OUTPUT)
  })
})
