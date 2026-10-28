// Slides: the shared AI panel prefs gate the chat log — showStreaming holds
// the reply pending until completion, showThinking renders the model's
// reasoning as a collapsible block fed by the transport's reasoning deltas,
// and showToolCalls hides the per-tool activity entries.
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

const agentHarness = vi.hoisted(() => ({
  events: null as null | {
    onText?: (text: string) => void
    onToolStart?: (call: { id: string; name: string; input: unknown }) => void
    onToolExecuted?: (event: {
      call: { id: string; name: string; input: unknown }
      execution: { summary: string; output?: string; isError?: boolean; mutated?: boolean }
    }) => void
    onDone?: (result: { text: string; cancelled: boolean; turnLimit: boolean }) => void
  },
  transport: null as null | import('@genoffice/agent-core').AgentTransport,
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

      constructor(options: {
        events?: typeof agentHarness.events
        transport?: import('@genoffice/agent-core').AgentTransport
      }) {
        agentHarness.events = options.events ?? null
        agentHarness.transport = options.transport ?? null
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
import { applyAiPanelPrefs } from '@genoffice/ui'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((provider) => [provider.id, { apiKey: '', model: provider.defaultModel }]),
  ) as AiSettings['providers'],
}

const mountedRoots: Array<{ root: Root; container: HTMLElement }> = []
let nonce = 0

function mountPanel(): { root: Root; container: HTMLElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() =>
    root.render(
      createElement(
        AiPanel,
        panelProps({
          preset: { text: `make a deck ${nonce}`, nonce: ++nonce, autoRun: true },
        }),
      ),
    ),
  )
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

/** slidesApi bridge mock: capture the stream listener so tests can push chunks */
function installSlidesApi(): { push: (chunk: Record<string, unknown>) => void } {
  let listener: ((chunk: Record<string, unknown>) => void) | null = null
  const requestIds: string[] = []
  Object.defineProperty(window, 'slidesApi', {
    configurable: true,
    value: {
      aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
      beginHistoryBatch: vi.fn(async () => false),
      endHistoryBatch: vi.fn(async () => null),
      aiLogRunFailure: vi.fn(async () => undefined),
      onAiStream: (l: (chunk: Record<string, unknown>) => void) => {
        listener = l
        return () => {
          listener = null
        }
      },
      aiStream: (request: { requestId: string }) => {
        requestIds.push(request.requestId)
        return Promise.resolve()
      },
      aiStreamCancel: vi.fn(),
    },
  })
  return {
    push: (chunk: Record<string, unknown>) =>
      listener?.({ requestId: requestIds.at(-1), ...chunk }),
  }
}

/** Stream one turn through the panel's real transport and push reasoning chunks */
function pushReasoning(api: ReturnType<typeof installSlidesApi>, text: string): void {
  agentHarness.transport!.stream(
    { system: 's', messages: [], tools: [] },
    {
      onDelta: () => {},
      onToolCall: () => {},
      onDone: () => {},
      onError: () => {},
    },
  )
  api.push({ type: 'reasoning', text })
  api.push({ type: 'done' })
}

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

beforeEach(() => {
  agentHarness.events = null
  agentHarness.transport = null
  agentHarness.run.mockReset()
  agentHarness.cancel.mockReset()
  agentHarness.reset.mockReset()
  agentHarness.restore.mockReset()
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
  // back to the all-on defaults for the next test
  applyAiPanelPrefs({})
})

describe('AiPanel pref gating (slides)', () => {
  it('renders reasoning, streamed text and tool entries while every pref is on', async () => {
    const api = installSlidesApi()
    const { container, root } = mountPanel()
    await flushEffects()

    expect(agentHarness.events).not.toBeNull()
    pushReasoning(api, 'First, consider the audience.')
    act(() => agentHarness.events!.onText?.('partial answer'))
    act(() => {
      agentHarness.events!.onToolStart?.({ id: 'tool-1', name: 'apply_ops', input: {} })
      agentHarness.events!.onToolExecuted?.({
        call: { id: 'tool-1', name: 'apply_ops', input: {} },
        execution: { summary: 'Changed the title fill', output: 'ok' },
      })
    })

    // reasoning block: attached to the streaming assistant message, live open
    const reasoning = container.querySelector<HTMLElement>('.ai-reasoning')
    expect(reasoning).not.toBeNull()
    expect(reasoning!.querySelector('.ai-reasoning-head')!.className).toContain('running')
    expect(reasoning!.querySelector('.ai-reasoning-body')!.className).toContain('open')
    expect(reasoning!.querySelector('.ai-reasoning-body-inner')!.textContent).toContain('audience')
    // live deltas and the tool activity entry render
    expect(container.querySelector('.ai-msg-assistant')!.textContent).toContain('partial answer')
    expect(container.querySelector('.ai-work-group')).not.toBeNull()

    act(() =>
      agentHarness.events!.onDone?.({ text: 'final answer', cancelled: false, turnLimit: false }),
    )
    await flushEffects()

    // completion: full text, reasoning kept on the message and now collapsed
    expect(container.querySelector('.ai-msg-assistant')!.textContent).toContain('final answer')
    const settled = container.querySelector<HTMLElement>('.ai-reasoning')!
    expect(settled.querySelector('.ai-reasoning-head')!.className).not.toContain('running')
    expect(settled.querySelector('.ai-reasoning-body')!.className).not.toContain('open')
    act(() => settled.querySelector<HTMLButtonElement>('.ai-reasoning-head')!.click())
    expect(settled.querySelector('.ai-reasoning-body')!.className).toContain('open')
  })

  it('holds the reply, hides reasoning and tool entries while every pref is off', async () => {
    const api = installSlidesApi()
    applyAiPanelPrefs({ showStreaming: false, showThinking: false, showToolCalls: false })
    const { container, root } = mountPanel()
    await flushEffects()

    pushReasoning(api, 'never shown')
    act(() => agentHarness.events!.onText?.('hidden delta'))
    act(() => {
      agentHarness.events!.onToolStart?.({ id: 'tool-1', name: 'apply_ops', input: {} })
      agentHarness.events!.onToolExecuted?.({
        call: { id: 'tool-1', name: 'apply_ops', input: {} },
        execution: { summary: 'Changed the title fill', output: 'ok' },
      })
    })

    // pending indicator instead of live deltas; no reasoning, no tool entries
    expect(container.querySelector('.ai-typing')).not.toBeNull()
    expect(container.textContent).not.toContain('hidden delta')
    expect(container.querySelector('.ai-reasoning')).toBeNull()
    expect(container.querySelector('.ai-work-group')).toBeNull()

    act(() =>
      agentHarness.events!.onDone?.({
        text: 'revealed at completion',
        cancelled: false,
        turnLimit: false,
      }),
    )
    await flushEffects()

    // the full text is revealed once the run completes
    expect(container.querySelector('.ai-typing')).toBeNull()
    expect(container.querySelector('.ai-msg-assistant')!.textContent).toContain(
      'revealed at completion',
    )
  })

  it('hides tool entries of an already-streamed reply when the pref turns off', async () => {
    const api = installSlidesApi()
    const { container, root } = mountPanel()
    await flushEffects()

    act(() => {
      agentHarness.events!.onToolStart?.({ id: 'tool-1', name: 'apply_ops', input: {} })
      agentHarness.events!.onToolExecuted?.({
        call: { id: 'tool-1', name: 'apply_ops', input: {} },
        execution: { summary: 'Changed the title fill', output: 'ok' },
      })
      agentHarness.events!.onDone?.({ text: 'done', cancelled: false, turnLimit: false })
    })
    await flushEffects()
    expect(container.querySelector('.ai-work-group')).not.toBeNull()

    // turning the pref off afterwards hides the entries live, no rerun needed
    act(() => {
      applyAiPanelPrefs({ showToolCalls: false })
    })
    expect(container.querySelector('.ai-work-group')).toBeNull()
    // the reply text itself stays visible
    expect(container.querySelector('.ai-msg-assistant')!.textContent).toContain('done')
  })
})
