// The docs composer's slash palette (useAiSlash): typing `/` at the draft start
// opens the shared skill listbox above the textarea, and Enter picks — the skill
// lands in the attached chips, the `/token` is struck from the draft, and the
// half-typed message is never sent (the palette's keys precede the composer's
// Enter-to-send, which lives in the shared AiComposer).
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
    AI_PROVIDERS.map((p) => [p.id, { apiKey: 'k', model: p.defaultModel }]),
  ) as AiSettings['providers'],
}

const SKILLS = [
  { name: 'tables', description: 'Header rows and totals' },
  { name: 'tone', description: 'Write plainly' },
]

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
    textarea.setSelectionRange(text.length, text.length)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function press(textarea: HTMLTextAreaElement, key: string): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  act(() => textarea.dispatchEvent(event))
  return event
}

const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 0))

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

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

describe('slash skill palette', () => {
  it('opens on a leading slash; Enter attaches the skill and strips the token without sending', async () => {
    const editor = createEditor()
    const aiStream = vi.fn(async () => {})
    ;(window as unknown as Record<string, unknown>).desktop = {
      aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
      skillsList: vi.fn(async () => SKILLS),
      onAiStream: vi.fn(() => () => {}),
      aiStream,
      aiStreamCancel: vi.fn(async () => {}),
    }
    const mounted = mount(createElement(AiPanel, panelProps(editor)))
    try {
      const textarea = mounted.container.querySelector<HTMLTextAreaElement>(
        '.ai-input-box textarea',
      )!
      typeInto(textarea, '/to')
      await act(async () => {
        await flush()
        await flush()
      })
      const pop = mounted.container.querySelector<HTMLElement>('.ai-slash-pop')!
      expect(pop).not.toBeNull()
      expect(pop.getAttribute('role')).toBe('listbox')
      expect([...pop.querySelectorAll('.ai-slash-skill-name')].map((n) => n.textContent)).toEqual(
        ['tables', 'tone'],
      )

      const event = press(textarea, 'Enter')
      await act(async () => {
        await flush()
      })
      expect(event.defaultPrevented).toBe(true)
      expect(textarea.value).toBe('')
      expect(mounted.container.querySelector('.ai-slash-pop')).toBeNull()
      expect(
        [...mounted.container.querySelectorAll('.ai-skill-chip-name')].map((n) => n.textContent),
      ).toEqual(['tables'])
      expect(aiStream).not.toHaveBeenCalled()
      expect(mounted.container.querySelector('.ai-typing')).toBeNull()
    } finally {
      mounted.cleanup()
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })
})
