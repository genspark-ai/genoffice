// Slides: the approval gate for risk:'high' deck-editing tools
// (execute_slide_script, apply_ops, regenerate_slide, generate_deck,
// insert_web_image, replace_image, edit_chart — see slides-skill.ts). A write
// tool parks at the inline allow/deny card before it executes: denying must
// leave the deck untouched and hand the model a do-not-retry result, allowing
// runs the tool. The real AgentLoop runs against a stubbed window.slidesApi
// stream (the panel drives the run through its preset auto-send).
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { TOOL_DENIED_OUTPUT } from '@genoffice/agent-core'
import type { RenderSlide } from '@genoffice/pptx-render'

// react-konva's node entry requires the native 'canvas' package; these tests do not draw.
vi.mock('react-konva', () => {
  const stub = () => null
  return {
    Stage: stub,
    Layer: stub,
    Rect: stub,
    Group: stub,
    Transformer: stub,
    Line: stub,
    Arrow: stub,
    Text: stub,
    Ellipse: stub,
    Image: stub,
    Path: stub,
    Circle: stub,
    Arc: stub,
  }
})

import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((provider) => [provider.id, { apiKey: '', model: provider.defaultModel }]),
  ) as AiSettings['providers'],
}

/** one minimal deck page: a single text element the script can rewrite */
const slide = {
  nodes: [
    {
      type: 'text',
      durableId: 't1',
      sourceId: 't1',
      fill: { kind: 'none' },
      text: {
        lines: [
          {
            runs: [{ text: 'original body', color: '#000000', fontSizePx: 18, isBullet: false }],
          },
        ],
      },
      box: { x: 40, y: 40, w: 320, h: 40, rotationDeg: 0 },
    },
  ],
  widthPx: 960,
  heightPx: 540,
} as unknown as RenderSlide

type Message = { role: string; results?: Array<{ output: string }> }
type Chunk = {
  requestId: string
  type: string
  text?: string
  toolCall?: { id: string; name: string; input: Record<string, unknown> }
}

let listener: ((chunk: Chunk) => void) | null = null
const requests: Array<{ requestId: string; messages: Message[] }> = []
const applyEditScript = vi.fn(async () => ({ slide }))

function installSlidesApi(): void {
  Object.defineProperty(window, 'slidesApi', {
    configurable: true,
    value: {
      aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
      beginHistoryBatch: vi.fn(async () => false),
      endHistoryBatch: vi.fn(async () => null),
      aiLogRunFailure: vi.fn(async () => undefined),
      applyEditScript,
      readDeckImage: vi.fn(async () => null),
      onAiStream: (l: (chunk: Chunk) => void) => {
        listener = l
        return () => {
          listener = null
        }
      },
      aiStream: (request: { requestId: string; messages: Message[] }) => {
        requests.push(request)
        return Promise.resolve()
      },
      aiStreamCancel: vi.fn(),
    },
  })
}

const mountedRoots: Array<{ root: Root; container: HTMLElement }> = []
let nonce = 0
/** the App's applySlide: the write-landed assertion point */
const applySlide = vi.fn()

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  Element.prototype.scrollTo ??= () => {}
})

beforeEach(() => {
  requests.length = 0
  applyEditScript.mockClear()
  applySlide.mockClear()
  applyEditScript.mockImplementation(async () => ({ slide }))
  installSlidesApi()
})

afterEach(() => {
  for (const { root, container } of mountedRoots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
  listener = null
  delete (window as unknown as { slidesApi?: unknown }).slidesApi
})

function mountPanel(): HTMLDivElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() =>
    root.render(
      createElement(AiPanel, {
        slides: [slide],
        current: 0,
        selectedIds: [],
        images: new Map<string, HTMLImageElement>(),
        applySlide,
        applyDeck: () => {},
        fitWidthPx: 960,
        settings,
        open: true,
        onExpand: () => {},
        onCollapse: () => {},
        preset: { text: `rewrite the title ${nonce}`, nonce: ++nonce, autoRun: true },
      }),
    ),
  )
  mountedRoots.push({ root, container })
  return container
}

const flush = async (): Promise<void> => {
  for (let i = 0; i < 6; i++) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 1))
    })
  }
}

/** answer the current turn with an execute_slide_script call, then close it */
async function answerWithScript(): Promise<void> {
  await act(async () => {
    const turn = requests.at(-1)!
    listener?.({
      requestId: turn.requestId,
      type: 'tool-call',
      toolCall: {
        id: 't1',
        name: 'execute_slide_script',
        input: {
          slideIndex: 0,
          code: "setText('t1', 'AI edited')",
          explanation: 'rewrite the title text',
        },
      },
    })
    listener?.({ requestId: turn.requestId, type: 'done' })
    await Promise.resolve()
  })
}

describe('AI panel tool confirmation', () => {
  it('parks a write tool at the allow/deny card; denying leaves the deck untouched', async () => {
    const container = mountPanel()
    await flush()
    expect(requests, 'the transport started a stream').not.toHaveLength(0)
    const before = requests.length

    await answerWithScript()
    // the card is in the conversation and the loop holds at the gate
    expect(container.querySelector('.ai-confirm-card')).not.toBeNull()
    expect(requests).toHaveLength(before)

    const deny = container.querySelector<HTMLButtonElement>('.ai-confirm-deny')!
    await act(async () => {
      deny.click()
    })
    await flush()

    // the tool never ran: neither the edit-script IPC nor the deck landing
    expect(applyEditScript).not.toHaveBeenCalled()
    expect(applySlide).not.toHaveBeenCalled()
    // the model received the denial as an error tool result it can act on
    expect(requests.length).toBeGreaterThan(before)
    const toolMessage = requests[requests.length - 1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).toBe(TOOL_DENIED_OUTPUT)
    // the settled card stays in the flow as the record of the decision
    expect(container.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe('denied')
    expect(container.querySelector('.ai-confirm-deny')).toBeNull()
  })

  it('allowing executes the write tool', async () => {
    const container = mountPanel()
    await flush()
    expect(requests).not.toHaveLength(0)
    const before = requests.length

    await answerWithScript()
    const allow = container.querySelector<HTMLButtonElement>('.ai-confirm-allow')!
    await act(async () => {
      allow.click()
    })
    await flush()

    // the write happened: the script's edits went through the edit-script IPC
    // and landed on the deck through the panel's applySlide
    expect(applyEditScript).toHaveBeenCalledTimes(1)
    expect(applySlide).toHaveBeenCalledWith(0, slide)
    // the tool result went back to the model (its round-trip continued)
    expect(requests.length).toBeGreaterThan(before)
    const toolMessage = requests[requests.length - 1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).not.toBe(TOOL_DENIED_OUTPUT)
    expect(container.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe(
      'allowed',
    )
  })

  it('a read tool never asks: it runs straight through', async () => {
    const container = mountPanel()
    await flush()
    expect(requests).not.toHaveLength(0)
    const before = requests.length

    await act(async () => {
      const turn = requests.at(-1)!
      listener?.({
        requestId: turn.requestId,
        type: 'tool-call',
        toolCall: { id: 't1', name: 'read_slide', input: { slideIndex: 0 } },
      })
      listener?.({ requestId: turn.requestId, type: 'done' })
      await Promise.resolve()
    })
    await flush()

    expect(container.querySelector('.ai-confirm-card')).toBeNull()
    expect(applySlide).not.toHaveBeenCalled()
    expect(requests.length).toBeGreaterThan(before)
  })
})
