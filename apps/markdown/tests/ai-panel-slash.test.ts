import { afterEach, beforeAll, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { AiPanel, type MarkdownAiDeps } from '../src/renderer/ai/AiPanel'
import type { MarkdownApi } from '../src/shared/ipc'

// Slash-palette wiring, driven through the mounted composer textarea: typing
// `/query` opens the shared listbox, Enter picks, and the pick must attach the
// skill, strike the token, and never leak the Enter into a send.

const roots: Root[] = []

beforeAll(() => {
  ;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
  Element.prototype.scrollTo ??= () => {}
})

const deps: MarkdownAiDeps = {
  getEditor: () => null,
  getFrontmatter: () => '',
  setFrontmatter: () => {},
  getSnapshot: () => ({ body: '', frontmatter: '' }),
  restoreSnapshot: () => {},
  onRunDone: () => {},
}

const SKILLS = [
  { name: 'proofread', description: 'Polish the draft' },
  { name: 'summarize', description: 'Shorten the draft' },
]

/** the panel's own IPC surface, minus everything a mounted idle composer needs */
function installApi(): { streams: unknown[] } {
  const streams: unknown[] = []
  const api = {
    aiGskStatus: async () => ({ loggedIn: false }),
    getAiSettings: async () => ({ providers: [] }),
    setAiSettings: async () => {},
    onAiSettingsChanged: () => () => {},
    openAiModelSettings: async () => {},
    skillsList: async () => SKILLS,
    skillText: async () => 'skill body',
    setAiPanelPrefs: async (prefs: unknown) => prefs,
    openSourcePath: async () => {},
    aiStream: async (request: unknown) => {
      streams.push(request)
    },
    aiStreamCancel: async () => {},
    onAiStream: () => () => {},
  }
  ;(window as unknown as { markdownApi: MarkdownApi }).markdownApi = api as unknown as MarkdownApi
  return { streams }
}

function mountPanel(): HTMLDivElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  roots.push(root)
  act(() => {
    root.render(createElement(AiPanel, { deps, filePath: null, onCollapse: () => {} }))
  })
  return container
}

function typeInto(textarea: HTMLTextAreaElement, text: string): void {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

/** pump timers so the skills promise resolves and the caret's rAF lands */
const flush = async (): Promise<void> => {
  for (let i = 0; i < 3; i++) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 12))
    })
  }
}

afterEach(() => {
  for (const root of roots.splice(0)) act(() => root.unmount())
  document.body.innerHTML = ''
  delete (window as unknown as { markdownApi?: MarkdownApi }).markdownApi
})

describe('AiPanel slash palette', () => {
  it('typing / opens the listbox; picking attaches the skill and strips the token', async () => {
    const { streams } = installApi()
    const container = mountPanel()
    const textarea = container.querySelector<HTMLTextAreaElement>('textarea')!

    typeInto(textarea, '/proof')
    const listbox = container.querySelector('[role="listbox"]')
    expect(listbox).not.toBeNull()
    await flush()
    const options = listbox!.querySelectorAll('[role="option"]')
    expect(options.length).toBe(1)
    expect(options[0]!.textContent).toContain('proofread')

    act(() => {
      textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })
    await flush()

    expect(textarea.value).toBe('')
    expect(container.querySelector('[role="listbox"]')).toBeNull()
    const chips = container.querySelectorAll('.ai-skill-chip')
    expect(chips.length).toBe(1)
    expect(chips[0]!.textContent).toContain('proofread')
    // the palette consumed Enter: the half-typed token was never sent
    expect(streams.length).toBe(0)
  })

  it('picking an already-attached skill keeps it attached', async () => {
    installApi()
    const container = mountPanel()
    const textarea = container.querySelector<HTMLTextAreaElement>('textarea')!

    for (let i = 0; i < 2; i++) {
      typeInto(textarea, '/proof')
      await flush()
      act(() => {
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
      })
      await flush()
    }

    expect(textarea.value).toBe('')
    expect(container.querySelectorAll('.ai-skill-chip').length).toBe(1)
  })
})
