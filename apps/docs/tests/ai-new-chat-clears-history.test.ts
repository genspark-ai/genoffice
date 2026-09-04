// "New chat" must clear the restored history transcript, not just the live
// turn: previously the previous conversation stayed painted on screen with
// no way to dismiss it (#195).
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: '', model: p.defaultModel }]),
  ) as AiSettings['providers'],
}

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

function mount(element: React.ReactElement): {
  container: HTMLElement
  cleanup: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

function panelProps(editor: Editor) {
  // No onCollapse: the New chat button is then the only .ai-header-btn,
  // which keeps the selector independent of the active i18n locale.
  return {
    editor,
    blocks: [],
    settings,
    open: true,
  }
}

const HISTORY = [{ role: 'user', text: 'earlier question' }]

function mockProjectApi() {
  const win = window as unknown as { projectApi?: unknown }
  const previous = win.projectApi
  win.projectApi = {
    resolveChat: vi.fn(async () => ({ projectId: 'p', chatId: 'c' })),
    loadChat: vi.fn(async () => HISTORY),
  }
  return () => {
    win.projectApi = previous
  }
}

async function flush(): Promise<void> {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

beforeAll(() => {
  // jsdom has no scrollTo; the panel auto-scrolls its chat log
  Element.prototype.scrollTo ??= () => {}
})

describe('AiPanel new chat', () => {
  it('clears the restored history transcript along with the live turn', async () => {
    const restoreApi = mockProjectApi()
    try {
      const editor = createEditor()
      const { container, cleanup } = mount(createElement(AiPanel, panelProps(editor)))
      try {
        await flush()
        await flush()
        // the previous conversation is painted above the live turn…
        expect(container.querySelectorAll('.ai-msg-historic').length).toBe(1)
        // …and New chat is offered even though the live turn is empty
        const button = container.querySelector<HTMLButtonElement>('.ai-header-btn')
        expect(button).not.toBeNull()

        act(() => button!.click())

        // both the transcript and the button are gone
        expect(container.querySelectorAll('.ai-msg-historic').length).toBe(0)
        expect(container.querySelector('.ai-header-btn')).toBeNull()
      } finally {
        cleanup()
        editor.destroy()
      }
    } finally {
      restoreApi()
    }
  })
})
