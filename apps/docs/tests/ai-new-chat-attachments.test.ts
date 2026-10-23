import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { beforeAll, describe, expect, it, vi } from 'vitest'

import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { t } from '../src/renderer/i18n/locale'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

// New chat must drop staged composer attachments along with the transcript:
// otherwise unsent files silently join the next chat's file context
// (availableAttachments merges sent + live composer files).

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

function mount(element: React.ReactElement): { container: HTMLElement; cleanup: () => void } {
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
  return { editor, blocks: [], settings, open: true }
}

const HISTORY = [{ role: 'user', text: 'earlier question' }]

function mockApis() {
  const win = window as unknown as {
    projectApi?: unknown
    desktop?: unknown
  }
  const previousApi = win.projectApi
  const previousDesktop = win.desktop
  win.projectApi = {
    resolveChat: vi.fn(async () => ({ projectId: 'p', chatId: 'c' })),
    loadChat: vi.fn(async () => HISTORY),
  }
  win.desktop = {
    pickAttachments: vi.fn(async () => ({
      accepted: [{ path: '/tmp/plan.png', name: 'plan.png', ext: 'png', sizeBytes: 128 }],
      rejected: [],
    })),
    readAttachmentImage: vi.fn(async () => ({ ok: false })),
  }
  return () => {
    win.projectApi = previousApi
    win.desktop = previousDesktop
  }
}

async function flush(): Promise<void> {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

describe('AiPanel new chat attachments', () => {
  it('drops staged composer files along with the transcript', async () => {
    const restoreApis = mockApis()
    try {
      const editor = createEditor()
      const { container, cleanup } = mount(createElement(AiPanel, panelProps(editor)))
      try {
        await flush()
        await flush()
        // Stage a file through the "+" menu's Add file item.
        const plus = container.querySelector<HTMLButtonElement>('.ai-plus-btn')
        expect(plus).not.toBeNull()
        act(() => plus!.click())
        const addFile = container.querySelector<HTMLButtonElement>('.ai-plus-pop .ai-plus-item')
        expect(addFile).not.toBeNull()
        await act(async () => {
          addFile!.click()
          await new Promise((resolve) => setTimeout(resolve, 0))
        })
        await flush()
        expect(container.querySelector('.ai-attachments')).not.toBeNull()

        // New chat clears the composer strip (transcript goes with it).
        const button = container.querySelector<HTMLButtonElement>(
          `button[aria-label="${t('aiNewChatTitle')}"]`,
        )
        expect(button).not.toBeNull()
        act(() => button!.click())
        expect(container.querySelector('.ai-attachments')).toBeNull()
        expect(container.querySelectorAll('.ai-msg-historic').length).toBe(0)
      } finally {
        cleanup()
        editor.destroy()
      }
    } finally {
      restoreApis()
    }
  })
})
