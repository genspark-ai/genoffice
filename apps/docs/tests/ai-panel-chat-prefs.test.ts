// Chat display toggles from the shell's AI panel preferences, live-read through
// useAiPanelPrefs: showStreaming holds the reply until completion, showThinking
// renders the model's reasoning deltas as a collapsible block on the assistant
// message (forwarded through the docs transport tee), showToolCalls hides the
// per-tool-call activity entries.
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { applyAiPanelPrefs } from '@genoffice/ui'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: 'k', model: p.defaultModel }]),
  ) as AiSettings['providers'],
}

/** stream chunk shape the preload forwards verbatim */
type Chunk = {
  requestId: string
  type: 'delta' | 'reasoning' | 'tool-call' | 'usage' | 'done' | 'error' | 'ping'
  text?: string
  toolCall?: { id: string; name: string; input: Record<string, unknown> }
  error?: string
}
type ChunkListener = (chunk: Chunk) => void

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'EVs market research' }],
        },
      ],
    },
  })
}

/** a mock preload bridge whose stream is driven manually by the test */
function manualStreamBridge() {
  let listener: ChunkListener | null = null
  let requestId = ''
  return {
    bridge: {
      aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
      onAiStream: (l: ChunkListener) => {
        listener = l
        return () => {
          listener = null
        }
      },
      aiStream: vi.fn(async (request: { requestId: string }) => {
        requestId = request.requestId
      }),
      aiStreamCancel: vi.fn(async () => {}),
    },
    emit: (chunk: Omit<Chunk, 'requestId'>) => listener?.({ requestId, ...chunk }),
  }
}

/** a mock preload bridge that answers each model turn by itself (tool runs) */
type TurnScript = (emit: (chunk: Omit<Chunk, 'requestId'>) => void) => void

function scriptedStreamBridge(turns: TurnScript[]) {
  let listener: ChunkListener | null = null
  let turn = 0
  return {
    bridge: {
      aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
      onAiStream: (l: ChunkListener) => {
        listener = l
        return () => {
          listener = null
        }
      },
      aiStream: vi.fn(async (request: { requestId: string }) => {
        const id = request.requestId
        const emit = (chunk: Omit<Chunk, 'requestId'>) => listener?.({ requestId: id, ...chunk })
        turns[Math.min(turn++, turns.length - 1)](emit)
      }),
      aiStreamCancel: vi.fn(async () => {}),
    },
  }
}

function mount(element: React.ReactElement): {
  container: HTMLElement
  root: Root
  cleanup: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    root,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

function typeInto(textarea: HTMLTextAreaElement, text: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 0))
const settle = async (): Promise<void> => {
  for (let i = 0; i < 12; i++) await flush()
}

function panelProps(editor: Editor) {
  return {
    editor,
    blocks: [],
    settings,
    open: true,
    onExpand: () => {},
    onCollapse: () => {},
  }
}

async function send(container: HTMLElement, text: string): Promise<void> {
  const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
  typeInto(textarea, text)
  const sendBtn = container.querySelector<HTMLButtonElement>('.ai-send-btn')!
  await act(async () => {
    sendBtn.click()
    await settle()
  })
}

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  // reset the module-global prefs store to the defaults for the next test
  applyAiPanelPrefs({})
  localStorage.removeItem('docs-ai-panel-width')
})

afterAll(() => {
  applyAiPanelPrefs({})
})

describe('showStreaming', () => {
  it('streams assistant text live while on (default)', async () => {
    const editor = createEditor()
    const { bridge, emit } = manualStreamBridge()
    ;(window as unknown as Record<string, unknown>).desktop = bridge
    const mounted = mount(createElement(AiPanel, panelProps(editor)))
    try {
      await act(async () => {
        const textarea =
          mounted.container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
        typeInto(textarea, 'summarize this')
        mounted.container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click()
        await settle()
      })
      emit({ type: 'delta', text: 'Hello world' })
      await act(async () => {
        await flush()
      })
      expect(mounted.container.textContent).toContain('Hello world')
      emit({ type: 'done' })
      await act(async () => {
        await settle()
      })
      expect(mounted.container.querySelector('.ai-typing')).toBeNull()
      expect(mounted.container.textContent).toContain('Hello world')
    } finally {
      mounted.cleanup()
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })

  it('holds the reply pending while off, revealing the full text at completion', async () => {
    applyAiPanelPrefs({ showStreaming: false })
    const editor = createEditor()
    const { bridge, emit } = manualStreamBridge()
    ;(window as unknown as Record<string, unknown>).desktop = bridge
    const mounted = mount(createElement(AiPanel, panelProps(editor)))
    try {
      await act(async () => {
        const textarea =
          mounted.container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
        typeInto(textarea, 'summarize this')
        mounted.container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click()
        await settle()
      })
      emit({ type: 'delta', text: 'Hello world' })
      emit({ type: 'delta', text: 'Hello world, fuller' })
      await act(async () => {
        await flush()
      })
      // pending: typing indicator, no reply text in the log
      expect(mounted.container.querySelector('.ai-typing')).not.toBeNull()
      expect(mounted.container.textContent).not.toContain('Hello world')
      emit({ type: 'done' })
      await act(async () => {
        await settle()
      })
      expect(mounted.container.querySelector('.ai-typing')).toBeNull()
      expect(mounted.container.textContent).toContain('Hello world, fuller')
    } finally {
      mounted.cleanup()
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })
})

describe('showThinking', () => {
  it('renders reasoning deltas as a live block that collapses when the stream ends', async () => {
    const editor = createEditor()
    const { bridge, emit } = manualStreamBridge()
    ;(window as unknown as Record<string, unknown>).desktop = bridge
    const mounted = mount(createElement(AiPanel, panelProps(editor)))
    try {
      await act(async () => {
        const textarea =
          mounted.container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
        typeInto(textarea, 'summarize this')
        mounted.container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click()
        await settle()
      })
      emit({ type: 'reasoning', text: 'step one: ' })
      emit({ type: 'reasoning', text: 'read the doc' })
      await act(async () => {
        await flush()
      })
      // expanded/live while streaming, attached to the assistant message
      const block = mounted.container.querySelector<HTMLElement>('.ai-thinking')!
      expect(block).not.toBeNull()
      expect(block.querySelector('.ai-thinking-body')!.classList.contains('open')).toBe(true)
      expect(block.textContent).toContain('step one: read the doc')
      emit({ type: 'delta', text: 'Here is the summary.' })
      emit({ type: 'done' })
      await act(async () => {
        await settle()
      })
      // survives the stream end, collapsed
      const settled = mounted.container.querySelector<HTMLElement>('.ai-thinking')!
      expect(settled).not.toBeNull()
      expect(settled.querySelector('.ai-thinking-body')!.classList.contains('open')).toBe(false)
      expect(settled.textContent).toContain('step one: read the doc')
      expect(mounted.container.textContent).toContain('Here is the summary.')
    } finally {
      mounted.cleanup()
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })

  it('hides the reasoning block completely while off', async () => {
    applyAiPanelPrefs({ showThinking: false })
    const editor = createEditor()
    const { bridge, emit } = manualStreamBridge()
    ;(window as unknown as Record<string, unknown>).desktop = bridge
    const mounted = mount(createElement(AiPanel, panelProps(editor)))
    try {
      await act(async () => {
        const textarea =
          mounted.container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
        typeInto(textarea, 'summarize this')
        mounted.container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click()
        await settle()
      })
      emit({ type: 'reasoning', text: 'hidden thoughts' })
      emit({ type: 'delta', text: 'Visible reply.' })
      emit({ type: 'done' })
      await act(async () => {
        await settle()
      })
      expect(mounted.container.querySelector('.ai-thinking')).toBeNull()
      expect(mounted.container.textContent).not.toContain('hidden thoughts')
      expect(mounted.container.textContent).toContain('Visible reply.')
    } finally {
      mounted.cleanup()
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })
})

describe('showToolCalls', () => {
  const toolTurns = [
    (emit: (chunk: Omit<Chunk, 'requestId'>) => void) => {
      emit({
        type: 'tool-call',
        toolCall: {
          id: 'call-1',
          name: 'read_blocks',
          input: { startBlockIndex: 0, endBlockIndex: 0 },
        },
      })
      emit({ type: 'done' })
    },
    (emit: (chunk: Omit<Chunk, 'requestId'>) => void) => {
      emit({ type: 'delta', text: 'All done.' })
      emit({ type: 'done' })
    },
  ]

  it('shows the per-tool-call activity entries while on (default)', async () => {
    const editor = createEditor()
    const { bridge } = scriptedStreamBridge(toolTurns)
    ;(window as unknown as Record<string, unknown>).desktop = bridge
    const mounted = mount(createElement(AiPanel, panelProps(editor)))
    try {
      await send(mounted.container, 'read the first block')
      expect(mounted.container.querySelector('.ai-work-group')).not.toBeNull()
      expect(mounted.container.textContent).toContain('All done.')
    } finally {
      mounted.cleanup()
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })

  it('hides the tool activity entries while off (typing indicator stays)', async () => {
    applyAiPanelPrefs({ showToolCalls: false })
    const editor = createEditor()
    const { bridge } = scriptedStreamBridge(toolTurns)
    ;(window as unknown as Record<string, unknown>).desktop = bridge
    const mounted = mount(createElement(AiPanel, panelProps(editor)))
    try {
      await send(mounted.container, 'read the first block')
      expect(mounted.container.querySelector('.ai-work-group')).toBeNull()
      // the reply itself still lands
      expect(mounted.container.textContent).toContain('All done.')
    } finally {
      mounted.cleanup()
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })
})
