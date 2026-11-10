import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

const agentHarness = vi.hoisted(() => ({
  events: null as null | {
    onToolExecuted?: (event: {
      call: { id: string; name: string; input: unknown }
      execution: { summary: string; output?: string; isError?: boolean; mutated?: boolean }
    }) => void
    onDone?: (result: { text: string; cancelled: boolean; turnLimit: boolean }) => void
  },
  run: vi.fn(),
  cancel: vi.fn(),
  reset: vi.fn(),
  restore: vi.fn(),
}))

vi.mock('@genoffice/agent-core', async () => {
  const actual =
    await vi.importActual<typeof import('@genoffice/agent-core')>('@genoffice/agent-core')
  return {
    ...actual,
    AgentLoop: class MockAgentLoop {
      busy = false

      constructor(options: { events?: typeof agentHarness.events }) {
        agentHarness.events = options.events ?? null
      }

      run(...args: unknown[]) {
        agentHarness.run(...args)
      }

      cancel() {
        agentHarness.cancel()
      }

      reset() {
        agentHarness.reset()
      }

      restore(...args: unknown[]) {
        agentHarness.restore(...args)
      }
    },
  }
})

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
import { AI_PROVIDERS, type AiSettings, type AttachmentMeta } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((provider) => [provider.id, { apiKey: '', model: provider.defaultModel }]),
  ) as AiSettings['providers'],
}

const mountedRoots: Array<{ root: Root; container: HTMLElement }> = []

function mount(element: React.ReactElement): { root: Root; container: HTMLElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  mountedRoots.push({ root, container })
  return { root, container }
}

function panelProps(overrides: Record<string, unknown> = {}) {
  return {
    slides: [],
    current: 0,
    selectedIds: [],
    images: new Map<string, HTMLImageElement>(),
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 960,
    settings,
    open: true,
    onExpand: () => {},
    onCollapse: () => {},
    ...overrides,
  }
}

async function flushEffects(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
    await Promise.resolve()
  })
}

function installSlidesApi(): void {
  Object.defineProperty(window, 'slidesApi', {
    configurable: true,
    value: {
      aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
      beginHistoryBatch: vi.fn(async () => false),
      endHistoryBatch: vi.fn(async () => null),
      aiLogRunFailure: vi.fn(async () => undefined),
    },
  })
}

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

beforeEach(() => {
  agentHarness.events = null
  agentHarness.run.mockReset()
  agentHarness.cancel.mockReset()
  agentHarness.reset.mockReset()
  agentHarness.restore.mockReset()
  installSlidesApi()
  Object.defineProperty(window, 'projectApi', { configurable: true, value: undefined })
  Object.defineProperty(window, 'desktop', {
    configurable: true,
    value: { readAttachmentImage: vi.fn(async () => ({ ok: false })) },
  })
})

afterEach(() => {
  for (const { root, container } of mountedRoots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
  vi.restoreAllMocks()
})

describe('AiPanel agent run lifecycle (slides)', () => {
  it('persists a successful tool-only run so reopening can restore the completed turn', async () => {
    const appendChat = vi.fn(async () => undefined)
    Object.defineProperty(window, 'projectApi', {
      configurable: true,
      value: {
        resolveChat: vi.fn(async () => ({ projectId: 'default', chatId: 'deck-chat' })),
        loadChat: vi.fn(async () => []),
        appendChat,
        rebindChat: vi.fn(async () => ({ projectId: 'default', chatId: 'deck-chat' })),
      },
    })

    mount(createElement(AiPanel, panelProps({ currentFilePath: '/tmp/deck.pptx' })))
    await flushEffects()
    appendChat.mockClear()

    expect(agentHarness.events).not.toBeNull()
    act(() => {
      agentHarness.events!.onToolExecuted?.({
        call: { id: 'tool-1', name: 'apply_ops', input: { op: 'set_fill' } },
        execution: { summary: 'Changed the title fill', output: 'ok', mutated: true },
      })
      agentHarness.events!.onDone?.({ text: '', cancelled: false, turnLimit: false })
    })
    await flushEffects()

    expect(appendChat).toHaveBeenCalledWith(
      expect.objectContaining({
        projectId: 'default',
        chatId: 'deck-chat',
        role: 'assistant',
        text: '',
        tools: [
          expect.objectContaining({
            name: 'apply_ops',
            summary: 'Changed the title fill',
            output: 'ok',
          }),
        ],
      }),
    )
  })

  it('keeps readable images and continues the run when another attachment read rejects', async () => {
    const readAttachmentImage = vi.fn(async (path: string) => {
      if (path.endsWith('good.png')) return { ok: true, base64: 'AAAA', mime: 'image/png' }
      throw new Error('attachment bridge unavailable')
    })
    Object.defineProperty(window, 'desktop', {
      configurable: true,
      value: { readAttachmentImage },
    })
    const attachments: AttachmentMeta[] = [
      { path: '/tmp/good.png', name: 'good.png', ext: 'png', sizeBytes: 128 },
      { path: '/tmp/bad.png', name: 'bad.png', ext: 'png', sizeBytes: 128 },
    ]

    mount(
      createElement(
        AiPanel,
        panelProps({
          preset: {
            text: 'Polish this slide',
            nonce: 1,
            autoRun: true,
            attachments,
          },
        }),
      ),
    )
    await flushEffects()

    expect(readAttachmentImage).toHaveBeenCalledWith('/tmp/good.png')
    expect(readAttachmentImage).toHaveBeenCalledWith('/tmp/bad.png')
    expect(agentHarness.run).toHaveBeenCalledWith('Polish this slide', [
      { base64: 'AAAA', mime: 'image/png' },
    ])
  })
})
