import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { AiPanel, type HtmlAiDeps } from '../src/renderer/ai/AiPanel'
import { applyAiPanelPrefs, DEFAULT_AI_PANEL_PREFS } from '@genoffice/ui'
import type { AiSettings } from '@genoffice/ai-provider'
import type { HtmlApi } from '../src/shared/ipc'
import type { ParseMap } from '../src/renderer/document/parse-map'

/**
 * The renderer side of the AI panel preferences: what each switch promises the
 * user. The stream is scripted at the IPC boundary (window.htmlApi stub), so
 * the whole chain runs for real — preload chunk shapes, the agent loop, the
 * panel's own transport wrapper and its chat-log rendering.
 */
beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  // jsdom implements no scrolling: the panel's stick-to-bottom effect needs the method to exist
  if (!Element.prototype.scrollTo) Element.prototype.scrollTo = () => {}
})
const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup())
  vi.unstubAllGlobals()
  localStorage.clear()
  applyAiPanelPrefs(DEFAULT_AI_PANEL_PREFS)
})

type Chunk = { requestId: string; type: string; text?: string; toolCall?: unknown }

let streamListener: ((chunk: Chunk) => void) | null = null
let requests: Array<{ requestId: string }> = []
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
  aiStream: (request: { requestId: string }) => {
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

function makeDeps(): HtmlAiDeps {
  return {
    access: {
      getText: () => '',
      getVersion: () => 1,
      getMap: () => emptyMap,
      getLastManualVersion: () => 1,
      getFilePath: () => null,
      getSelectedSid: () => null,
      applyOps: () => ({ ok: true, ranges: [] }),
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

/** mount the panel inside its dock (the width var lands on the dock element) */
function mountPanel(): HTMLElement {
  streamListener = null
  requests = []
  const dock = document.createElement('div')
  dock.className = 'ai-dock'
  document.body.appendChild(dock)
  const root: Root = createRoot(dock)
  act(() =>
    root.render(
      createElement(AiPanel, {
        deps: makeDeps(),
        filePath: null,
        preset: { text: 'hello', nonce: ++nonceSeq },
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

describe('showThinking', () => {
  it('renders the reasoning as a collapsible block that survives the stream end', async () => {
    const dock = mountPanel()
    await settle()
    await push({ type: 'reasoning', text: 'First ' })
    await push({ type: 'reasoning', text: 'thought.' })
    expect(dock.querySelector('.ai-reasoning')).not.toBeNull()
    expect(dock.querySelector('.ai-reasoning-text')?.textContent).toBe('First thought.')
    // expanded/live while streaming
    expect(dock.querySelector('.ai-reasoning-body')?.classList.contains('open')).toBe(true)

    await push({ type: 'delta', text: 'The reply.' })
    await push({ type: 'done' })
    expect(dock.textContent).toContain('The reply.')
    // attached to the finished message, collapsed; a manual toggle re-opens it
    expect(dock.querySelector('.ai-reasoning-body')?.classList.contains('open')).toBe(false)
    act(() => (dock.querySelector('.ai-reasoning-summary') as HTMLButtonElement).click())
    expect(dock.querySelector('.ai-reasoning-body')?.classList.contains('open')).toBe(true)
  })

  it('off hides the block completely, on brings it back (data is kept, not dropped)', async () => {
    const dock = mountPanel()
    await settle()
    await push({ type: 'reasoning', text: 'thinking hard' })
    await push({ type: 'delta', text: 'Answer.' })
    await push({ type: 'done' })

    act(() => applyAiPanelPrefs({ ...DEFAULT_AI_PANEL_PREFS, showThinking: false }))
    expect(dock.querySelector('.ai-reasoning')).toBeNull()
    expect(dock.textContent).toContain('Answer.')

    act(() => applyAiPanelPrefs(DEFAULT_AI_PANEL_PREFS))
    expect(dock.querySelector('.ai-reasoning-text')?.textContent).toBe('thinking hard')
  })
})

describe('showStreaming', () => {
  it('off holds the reply pending behind the typing indicator, revealed at completion', async () => {
    act(() => applyAiPanelPrefs({ ...DEFAULT_AI_PANEL_PREFS, showStreaming: false }))
    const dock = mountPanel()
    await settle()
    await push({ type: 'delta', text: 'Held ' })
    await push({ type: 'delta', text: 'back text.' })
    expect(dock.textContent).not.toContain('Held')
    expect(dock.querySelector('.ai-typing-row')).not.toBeNull()

    await push({ type: 'done' })
    expect(dock.textContent).toContain('Held back text.')
    expect(dock.querySelector('.ai-typing-row')).toBeNull()
  })

  it('on streams the deltas live', async () => {
    const dock = mountPanel()
    await settle()
    await push({ type: 'delta', text: 'Live text.' })
    expect(dock.textContent).toContain('Live text.')
    await push({ type: 'done' })
    expect(dock.textContent).toContain('Live text.')
  })
})

describe('showToolCalls', () => {
  it('off hides the per-tool entries, on shows them; the reply is unaffected', async () => {
    const dock = mountPanel()
    await settle()
    // a tool call whose input failed to parse: the loop reports it as an
    // executed (error) activity and goes back to the model for the answer
    await push({
      type: 'tool-call',
      toolCall: { id: 't1', name: 'web_search', input: {}, inputError: 'stubbed parse failure' },
    })
    await push({ type: 'done' })
    await push({ type: 'delta', text: 'Finished anyway.' })
    await push({ type: 'done' })
    expect(dock.querySelector('.ai-work-group')).not.toBeNull()
    expect(dock.textContent).toContain('Finished anyway.')

    act(() => applyAiPanelPrefs({ ...DEFAULT_AI_PANEL_PREFS, showToolCalls: false }))
    expect(dock.querySelector('.ai-work-group')).toBeNull()
    expect(dock.textContent).toContain('Finished anyway.')

    act(() => applyAiPanelPrefs(DEFAULT_AI_PANEL_PREFS))
    expect(dock.querySelector('.ai-work-group')).not.toBeNull()
  })
})

describe('panel width rule', () => {
  it('auto width without a saved preference, re-derived on window resize', async () => {
    const dock = mountPanel()
    await settle()
    const initial = dock.style.getPropertyValue('--ai-panel-width')
    expect(initial).toBe(`${Math.max(360, Math.round(window.innerWidth * 0.22))}px`)

    Object.defineProperty(window, 'innerWidth', { value: 1900, configurable: true })
    try {
      act(() => window.dispatchEvent(new Event('resize')))
      expect(dock.style.getPropertyValue('--ai-panel-width')).toBe(
        `${Math.max(360, Math.round(1900 * 0.22))}px`,
      )
    } finally {
      Object.defineProperty(window, 'innerWidth', { value: 1024, configurable: true })
    }
  })

  it('a legacy saved width below 360 migrates up through the clamp', async () => {
    localStorage.setItem('html-ai-panel-width', '300')
    const dock = mountPanel()
    await settle()
    expect(dock.style.getPropertyValue('--ai-panel-width')).toBe('360px')
  })
})
