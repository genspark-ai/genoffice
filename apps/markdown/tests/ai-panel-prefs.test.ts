import { afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import {
  aiPanelAutoWidth,
  aiPanelClampWidth,
  applyAiPanelPrefs,
  DEFAULT_AI_PANEL_PREFS,
} from '@genoffice/ui'
import type { AiStreamChunk, AiStreamRequest, AiSettings } from '@genoffice/ai-provider'
import { AiPanel, type MarkdownAiDeps } from '../src/renderer/ai/AiPanel'
import {
  displayPanelWidth,
  loadSavedPanelWidth,
  PANEL_WIDTH_KEY,
} from '../src/renderer/ai/panel-width'
import type { MarkdownApi } from '../src/shared/ipc'

// ── width rule: the local helpers adopt the shared @genoffice/ui math ──

describe('AI panel width', () => {
  beforeEach(() => {
    localStorage.removeItem(PANEL_WIDTH_KEY)
  })

  it('keeps the markdown localStorage key for the dragged width', () => {
    localStorage.setItem(PANEL_WIDTH_KEY, '512')
    expect(loadSavedPanelWidth()).toBe(512)
  })

  it('treats a missing or malformed preference as "never dragged"', () => {
    expect(loadSavedPanelWidth()).toBe(null)
    localStorage.setItem(PANEL_WIDTH_KEY, 'garbage')
    expect(loadSavedPanelWidth()).toBe(null)
    localStorage.setItem(PANEL_WIDTH_KEY, '-4')
    expect(loadSavedPanelWidth()).toBe(null)
  })

  it('never dragged: the display width follows the window, floored at 360', () => {
    expect(displayPanelWidth(null, 1920)).toBe(aiPanelAutoWidth(1920))
    expect(displayPanelWidth(null, 1920)).toBeGreaterThan(360)
    expect(displayPanelWidth(null, 480)).toBe(360)
    // a WebContentsView is 0×0 until the shell lays it out
    expect(displayPanelWidth(null, 0)).toBe(360)
  })

  it('dragged width: clamped into the window, legacy values below 360 migrate up', () => {
    expect(displayPanelWidth(300, 1920)).toBe(360)
    expect(displayPanelWidth(900, 1920)).toBe(720)
    // 60% cap of a narrow window, never below the 360 floor
    expect(displayPanelWidth(500, 500)).toBe(360)
    expect(displayPanelWidth(420, 1920)).toBe(aiPanelClampWidth(420, 1920))
  })
})

// ── prefs gating: showStreaming / showThinking / showToolCalls in the chat ──
// The panel is driven end-to-end through a stubbed window.markdownApi stream,
// the same IPC chunk set the shared ai:stream handler pushes.

const roots: Root[] = []

// jsdom implements neither Element.scrollTo nor the smooth-scroll options; the
// panel's stick-to-bottom effect only needs the call to exist. The act flag
// lets loop callbacks that land outside an act() wrapper (async tool
// continuations) update state without warning.
beforeAll(() => {
  ;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
  Element.prototype.scrollTo = () => {}
})

const deps: MarkdownAiDeps = {
  getEditor: () => null,
  getFrontmatter: () => '',
  setFrontmatter: () => {},
  getSnapshot: () => ({ body: '', frontmatter: '' }),
  restoreSnapshot: () => {},
  onRunDone: () => {},
}

type StreamHandler = (chunk: AiStreamChunk) => void

function installStreamStub() {
  const requests: AiStreamRequest[] = []
  let listener: StreamHandler | null = null
  const api = {
    aiGskStatus: async () => ({ loggedIn: false }),
    getAiSettings: async () => ({ providers: [] }) as unknown as AiSettings,
    setAiSettings: async () => {},
    onAiSettingsChanged: () => () => {},
    openAiModelSettings: async () => {},
    skillsList: async () => [],
    skillText: async () => '',
    setAiPanelPrefs: async (prefs: unknown) => prefs,
    openSourcePath: async () => {},
    webSearch: async () => ({
      results: [{ title: 'GenOffice', url: 'https://example.com', snippet: 's' }],
      method: 'ddg',
    }),
    aiStream: async (request: AiStreamRequest) => {
      requests.push(request)
    },
    aiStreamCancel: async () => {},
    onAiStream: (handler: StreamHandler) => {
      listener = handler
      return () => {
        listener = null
      }
    },
  }
  ;(window as unknown as { markdownApi: MarkdownApi }).markdownApi = api as unknown as MarkdownApi

  /** push one chunk into the renderer (state updates land inside act) */
  const send = (chunk: AiStreamChunk) => {
    act(() => {
      listener?.(chunk)
    })
  }
  /** pump the task queue until the loop has started another turn */
  const waitForTurn = async (): Promise<AiStreamRequest> => {
    const seen = requests.length
    for (let i = 0; i < 200 && requests.length === seen; i++) {
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 1))
      })
    }
    if (requests.length === seen) throw new Error('the agent loop never started the next turn')
    return requests[requests.length - 1]!
  }
  return { send, waitForTurn }
}

function renderPanel(prefs: Partial<typeof DEFAULT_AI_PANEL_PREFS>): HTMLDivElement {
  localStorage.removeItem(PANEL_WIDTH_KEY)
  applyAiPanelPrefs({ ...DEFAULT_AI_PANEL_PREFS, ...prefs })
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  roots.push(root)
  // the preset auto-sends once on mount, the way a ribbon chip does
  act(() => {
    root.render(
      createElement(AiPanel, {
        deps,
        filePath: null,
        preset: { text: 'Why is the sky blue?', nonce: 1 },
        onCollapse: () => {},
      }),
    )
  })
  return container
}

const flush = async (): Promise<void> => {
  for (let i = 0; i < 6; i++) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 1))
    })
  }
}

afterEach(() => {
  for (const root of roots.splice(0)) act(() => root.unmount())
  document.body.innerHTML = ''
  delete (window as unknown as { markdownApi?: MarkdownApi }).markdownApi
})

describe('AiPanel prefs gating', () => {
  it('showStreaming on: deltas land in the assistant message as they arrive', async () => {
    const stream = installStreamStub()
    const container = renderPanel({})
    const turn = await stream.waitForTurn()
    stream.send({ requestId: turn.requestId, type: 'delta', text: 'Rayleigh scattering.' })
    expect(container.textContent).toContain('Rayleigh scattering.')
    stream.send({ requestId: turn.requestId, type: 'done' })
    await flush()
    expect(container.textContent).toContain('Rayleigh scattering.')
  })

  it('showStreaming off: the bubble stays pending until the run completes', async () => {
    const stream = installStreamStub()
    const container = renderPanel({ showStreaming: false })
    const turn = await stream.waitForTurn()
    stream.send({ requestId: turn.requestId, type: 'delta', text: 'Rayleigh scattering.' })
    // held: the typing indicator is the only visible status, no text
    expect(container.querySelector('.ai-typing')).not.toBeNull()
    expect(container.textContent).not.toContain('Rayleigh scattering.')
    stream.send({ requestId: turn.requestId, type: 'done' })
    await flush()
    // revealed in full at completion
    expect(container.querySelector('.ai-typing')).toBeNull()
    expect(container.textContent).toContain('Rayleigh scattering.')
  })

  it('showThinking on: reasoning renders as a collapsible block and survives stream end', async () => {
    const stream = installStreamStub()
    const container = renderPanel({})
    const turn = await stream.waitForTurn()
    stream.send({ requestId: turn.requestId, type: 'reasoning', text: 'thinking about light' })
    stream.send({ requestId: turn.requestId, type: 'delta', text: 'Rayleigh scattering.' })
    stream.send({ requestId: turn.requestId, type: 'done' })
    await flush()
    const block = container.querySelector('.ai-reasoning-body-inner')
    expect(block?.textContent).toContain('thinking about light')
    expect(container.querySelector('.ai-reasoning-head')?.getAttribute('aria-expanded')).toBe(
      'false',
    )
  })

  it('showThinking off: reasoning never renders', async () => {
    const stream = installStreamStub()
    const container = renderPanel({ showThinking: false })
    const turn = await stream.waitForTurn()
    stream.send({ requestId: turn.requestId, type: 'reasoning', text: 'thinking about light' })
    stream.send({ requestId: turn.requestId, type: 'delta', text: 'Rayleigh scattering.' })
    stream.send({ requestId: turn.requestId, type: 'done' })
    await flush()
    expect(container.querySelector('.ai-reasoning')).toBeNull()
    expect(container.textContent).not.toContain('thinking about light')
  })

  it('showToolCalls on: executed tool calls appear in the chat log', async () => {
    const stream = installStreamStub()
    const container = renderPanel({})
    const turn = await stream.waitForTurn()
    stream.send({
      requestId: turn.requestId,
      type: 'tool-call',
      toolCall: { id: 't1', name: 'web_search', input: { query: 'sky' } },
    })
    stream.send({ requestId: turn.requestId, type: 'done' })
    // the loop runs the tool, then opens a second turn for the final reply
    const turn2 = await stream.waitForTurn()
    stream.send({ requestId: turn2.requestId, type: 'delta', text: 'Blue light scatters more.' })
    stream.send({ requestId: turn2.requestId, type: 'done' })
    await flush()
    expect(container.querySelector('.ai-work-group')).not.toBeNull()
    expect(container.querySelector('.ai-step-row')).not.toBeNull()
    expect(container.textContent).toContain('Blue light scatters more.')
  })

  it('showToolCalls off: tool entries are hidden, the typing indicator remains', async () => {
    const stream = installStreamStub()
    const container = renderPanel({ showToolCalls: false })
    const turn = await stream.waitForTurn()
    stream.send({
      requestId: turn.requestId,
      type: 'tool-call',
      toolCall: { id: 't1', name: 'web_search', input: { query: 'sky' } },
    })
    stream.send({ requestId: turn.requestId, type: 'done' })
    const turn2 = await stream.waitForTurn()
    stream.send({ requestId: turn2.requestId, type: 'delta', text: 'Blue light scatters more.' })
    stream.send({ requestId: turn2.requestId, type: 'done' })
    await flush()
    expect(container.querySelector('.ai-work-group')).toBeNull()
    expect(container.querySelector('.ai-step-row')).toBeNull()
    expect(container.textContent).toContain('Blue light scatters more.')
  })
})
