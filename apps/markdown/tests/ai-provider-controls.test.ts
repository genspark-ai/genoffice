import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings, type AiSettings } from '@genoffice/ai-provider'
import { AiPanel } from '../src/renderer/ai/AiPanel'

const activeCleanups = new Set<() => void>()
const activeApiRestorers = new Set<() => void>()

function mount(element: React.ReactElement): {
  container: HTMLElement
  root: Root
  cleanup: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  let mounted = true
  const cleanup = () => {
    if (!mounted) return
    mounted = false
    activeCleanups.delete(cleanup)
    act(() => root.unmount())
    container.remove()
  }
  activeCleanups.add(cleanup)
  return {
    container,
    root,
    cleanup,
  }
}

function settings(provider: AiSettings['provider']): AiSettings {
  return { ...defaultAiSettings(), provider }
}

function installApi(initial: AiSettings, loggedIn = false) {
  let settingsChanged: ((next: AiSettings) => void) | undefined
  const api = {
    getAiSettings: vi.fn(async () => initial),
    setAiSettings: vi.fn(async () => {}),
    onAiSettingsChanged: vi.fn((handler: (next: AiSettings) => void) => {
      settingsChanged = handler
      return () => {
        settingsChanged = undefined
      }
    }),
    aiCodexStatus: vi.fn(async () => ({ loggedIn })),
    aiCodexLogin: vi.fn(async () => ({ loggedIn: true })),
    aiCodexCancelLogin: vi.fn(async () => {}),
    aiCodexLogout: vi.fn(async () => ({ loggedIn: false })),
    aiCodexCapabilities: vi.fn(async () => ({
      models: [
        {
          id: 'gpt-5.5',
          name: 'GPT-5.5',
          reasoningEfforts: ['none', 'high'],
          serviceTiers: [{ id: 'default', name: 'Standard' }],
        },
      ],
    })),
    onAiStream: vi.fn(() => () => {}),
    aiStream: vi.fn(async () => {}),
    aiStreamCancel: vi.fn(async () => {}),
    webSearch: vi.fn(async () => ({ results: [] })),
  }
  const previous = Object.getOwnPropertyDescriptor(window, 'markdownApi')
  Object.defineProperty(window, 'markdownApi', { configurable: true, value: api })
  const restore = () => {
    if (!activeApiRestorers.delete(restore)) return
    if (previous) Object.defineProperty(window, 'markdownApi', previous)
    else Reflect.deleteProperty(window, 'markdownApi')
  }
  activeApiRestorers.add(restore)
  return {
    api,
    emit(next: AiSettings) {
      settingsChanged?.(next)
    },
  }
}

async function waitFor(assertion: () => void): Promise<void> {
  await vi.waitFor(async () => {
    await act(async () => {
      await new Promise<void>((resolve) => queueMicrotask(resolve))
    })
    assertion()
  })
}

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  for (const cleanup of [...activeCleanups]) cleanup()
  for (const restore of [...activeApiRestorers]) restore()
  document.body.replaceChildren()
})

describe('Markdown Codex provider wiring', () => {
  it('keeps drafts editable while signed-out Codex disables Send', async () => {
    const installed = installApi(settings('openai-codex'))
    const { container, cleanup } = mount(
      createElement(AiPanel, {
        deps: {
          getEditor: () => null,
          getSnapshot: () => '',
          restoreSnapshot: () => {},
          onRunDone: () => {},
        },
        filePath: null,
        onCollapse: () => {},
      }),
    )
    await waitFor(() => expect(container.querySelector('.ai-provider-select-input')).not.toBeNull())
    expect(container.querySelector<HTMLButtonElement>('.ai-send-btn')?.disabled).toBe(true)
    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    act(() => {
      setter.call(textarea, 'draft')
      textarea.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect(textarea.value).toBe('draft')
    expect(installed.api.setAiSettings).not.toHaveBeenCalled()
    cleanup()
  })

  it('renders capability-backed Codex controls and applies external settings one-way', async () => {
    const installed = installApi(settings('openai-codex'), true)
    const { container, cleanup } = mount(
      createElement(AiPanel, {
        deps: {
          getEditor: () => null,
          getSnapshot: () => '',
          restoreSnapshot: () => {},
          onRunDone: () => {},
        },
        filePath: null,
        onCollapse: () => {},
      }),
    )
    await waitFor(() => expect(container.querySelector('.ai-codex-model-trigger')).not.toBeNull())
    installed.emit(settings('genspark'))
    await waitFor(() =>
      expect(container.querySelector<HTMLSelectElement>('.ai-provider-select-input')?.value).toBe(
        'genspark',
      ),
    )
    expect(installed.api.setAiSettings).not.toHaveBeenCalled()
    cleanup()
  })
})
