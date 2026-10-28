/**
 * AI panel display prefs (shell Settings → AI panel): showStreaming holds the
 * reply text until the run completes, showThinking renders the streamed
 * reasoning in a collapsible block (kept after the run ends), showToolCalls
 * hides the tool activity entries. Runs are driven through the real transport
 * chain with a stubbed window.pdfApi stream channel, so the reasoning tee in
 * createElectronTransport is exercised end to end.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest'
import { applyAiPanelPrefs, DEFAULT_AI_PANEL_PREFS } from '@genoffice/ui'

import { AiPanel } from '../src/renderer/ai/AiPanel'
import type { PdfAppDeps } from '../src/renderer/ai/tools'

type StreamChunk = {
  requestId: string
  type: 'delta' | 'reasoning' | 'tool-call' | 'usage' | 'done' | 'error' | 'ping'
  text?: string
  stopReason?: string
}

type StreamListener = (chunk: StreamChunk) => void

const streamListeners = new Set<StreamListener>()
let startedRequestId = ''

const pdfApi = {
  getAiSettings: () => Promise.resolve({}),
  setAiSettings: () => Promise.resolve(),
  onAiSettingsChanged: () => () => {},
  openAiModelSettings: () => Promise.resolve(),
  openSkillsSettings: () => Promise.resolve(),
  skillsList: () => Promise.resolve([]),
  skillText: () => Promise.resolve(''),
  gskStatus: () => Promise.resolve({ loggedIn: false }),
  aiStream: (request: { requestId: string }) => {
    startedRequestId = request.requestId
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

/** history the shared chat store hands back on open; per-test via chatHistory */
let chatHistory: Array<{
  role: 'user' | 'assistant'
  text: string
  tools?: Array<{ name: string; summary: string }>
}> = []

const projectApi = {
  resolveChat: () => Promise.resolve({ projectId: 'p', chatId: 'c' }),
  appendChat: () => Promise.resolve(),
  loadChat: () => Promise.resolve(chatHistory),
  rebindChat: () => Promise.resolve(null),
}

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
  chatHistory = []
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

describe('AiPanel display prefs', () => {
  it('renders streamed reasoning in a collapsible block and keeps it after the run', async () => {
    await renderPanel()
    await sendQuick()
    await emit({ type: 'reasoning', text: 'weighing the ' })
    await emit({ type: 'reasoning', text: 'options' })

    const block = container!.querySelector('.ai-reasoning')
    expect(block, 'reasoning block on the assistant message').not.toBeNull()
    expect(block!.textContent).toContain('weighing the options')
    expect(block!.querySelector('.ai-reasoning-summary')!.textContent).toContain('思考过程')
    // live while streaming: expanded without any click
    expect(block!.querySelector('.ai-reasoning-body')!.className).toContain('open')

    await emit({ type: 'delta', text: 'The answer' })
    expect(container!.querySelector('.ai-msg-assistant')!.textContent).toContain('The answer')

    await emit({ type: 'done', stopReason: 'end_turn' })
    // the block survives stream end and collapses; the manual toggle re-opens it
    const after = container!.querySelector('.ai-reasoning')!
    expect(after).not.toBeNull()
    expect(after.querySelector('.ai-reasoning-body')!.className).not.toContain('open')
    await act(async () => {
      after
        .querySelector<HTMLButtonElement>('.ai-reasoning-summary')!
        .dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })
    expect(container!.querySelector('.ai-reasoning-body')!.className).toContain('open')
  })

  it('hides the reasoning block entirely while showThinking is off', async () => {
    await act(async () => {
      applyAiPanelPrefs({ showThinking: false })
      await Promise.resolve()
    })
    await renderPanel()
    await sendQuick()
    await emit({ type: 'reasoning', text: 'secret thoughts' })
    await emit({ type: 'delta', text: 'The answer' })
    await emit({ type: 'done', stopReason: 'end_turn' })

    expect(container!.querySelector('.ai-reasoning')).toBeNull()
    expect(container!.querySelector('.ai-msg-assistant')!.textContent).toContain('The answer')
    // reasoning is still on the message: flipping the pref back reveals it
    await act(async () => {
      applyAiPanelPrefs({ showThinking: true })
      await Promise.resolve()
    })
    expect(container!.querySelector('.ai-reasoning')!.textContent).toContain('secret thoughts')
  })

  it('holds the reply pending while showStreaming is off and reveals it at completion', async () => {
    await act(async () => {
      applyAiPanelPrefs({ showStreaming: false })
      await Promise.resolve()
    })
    await renderPanel()
    await sendQuick()
    await emit({ type: 'reasoning', text: 'hmm' })

    await emit({ type: 'delta', text: 'The answer' })
    // no text deltas in the chat; the pending indicator is still up
    expect(container!.querySelector('.ai-msg-assistant')!.textContent).not.toContain('The answer')
    expect(container!.querySelector('.ai-typing')).not.toBeNull()

    await emit({ type: 'delta', text: ', in full' })
    expect(container!.querySelector('.ai-msg-assistant')!.textContent).not.toContain('The answer')

    await emit({ type: 'done', stopReason: 'end_turn' })
    const reply = container!.querySelector('.ai-msg-assistant')!
    expect(reply.textContent).toContain('The answer, in full')
    expect(container!.querySelector('.ai-typing')).toBeNull()
  })

  it('hides the tool activity entries while showToolCalls is off', async () => {
    chatHistory = [
      {
        role: 'assistant',
        text: 'Done',
        tools: [{ name: 'read_pages', summary: '读取第 1-2 页' }],
      },
    ]
    await renderPanel()
    expect(container!.querySelector('.ai-work-group'), 'tool entries by default').not.toBeNull()

    await act(async () => {
      applyAiPanelPrefs({ showToolCalls: false })
      await Promise.resolve()
    })
    expect(container!.querySelector('.ai-work-group')).toBeNull()
    // only the entries hide — the reply text stays
    expect(container!.querySelector('.ai-msg-assistant')!.textContent).toContain('Done')

    await act(async () => {
      applyAiPanelPrefs({ showToolCalls: true })
      await Promise.resolve()
    })
    expect(container!.querySelector('.ai-work-group')).not.toBeNull()
  })
})
