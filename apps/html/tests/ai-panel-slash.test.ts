import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { AiPanel, type HtmlAiDeps } from '../src/renderer/ai/AiPanel'
import { applyAiPanelPrefs, DEFAULT_AI_PANEL_PREFS } from '@genoffice/ui'
import type { AiSettings } from '@genoffice/ai-provider'
import type { HtmlApi } from '../src/shared/ipc'
import type { ParseMap } from '../src/renderer/document/parse-map'

/**
 * The composer's slash palette wiring: typing a `/token` opens the skill
 * listbox above the textarea, and picking one attaches the skill, strikes the
 * token out of the draft, and never lets the pick's Enter also send the
 * message. The IPC boundary is stubbed (window.htmlApi), so the panel and the
 * shared composer run for real.
 */
beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  // jsdom implements neither scrolling nor animation frames
  if (!Element.prototype.scrollTo) Element.prototype.scrollTo = () => {}
  if (!window.requestAnimationFrame) {
    vi.stubGlobal('requestAnimationFrame', (cb: FrameRequestCallback) => setTimeout(cb, 0))
  }
})
const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup())
  vi.unstubAllGlobals()
  localStorage.clear()
  applyAiPanelPrefs(DEFAULT_AI_PANEL_PREFS)
})

const SKILLS = [
  { name: 'tables', description: 'Header rows and totals' },
  { name: 'tone', description: 'Write plainly' },
]

let requests: Array<{ requestId: string }> = []

const htmlApi = {
  getAiSettings: () => Promise.resolve({} as AiSettings),
  setAiSettings: () => Promise.resolve(),
  onAiSettingsChanged: () => () => {},
  aiGskStatus: () => Promise.resolve({ loggedIn: false }),
  openAiModelSettings: () => Promise.resolve(),
  openSkillsSettings: () => Promise.resolve(),
  skillsList: () => Promise.resolve(SKILLS),
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
  onAiStream: () => () => {},
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

/** mount without a preset: nothing auto-sends, so a leaked Enter shows up in `requests` */
function mountPanel(): HTMLElement {
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
        preset: null,
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

/** flush the listSkills promise, the pick's caret frame, and any pending render */
async function settle(): Promise<void> {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

/** type into the composer the way a user does: value, caret, input event */
function type(dock: HTMLElement, text: string): void {
  const ta = dock.querySelector('textarea')!
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    setter.call(ta, text)
    ta.setSelectionRange(text.length, text.length)
    ta.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function press(dock: HTMLElement, key: string): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  act(() => dock.querySelector('textarea')!.dispatchEvent(event))
  return event
}

describe('slash palette', () => {
  it('opens the skill listbox on a leading slash and picks with Enter without sending', async () => {
    const dock = mountPanel()
    await settle()
    expect(dock.querySelector('.ai-slash-pop')).toBeNull()

    type(dock, '/to')
    await settle()
    const pop = dock.querySelector<HTMLElement>('.ai-slash-pop')!
    expect(pop.getAttribute('role')).toBe('listbox')
    // the same substring match the "+" menu uses: 'to' hits the tone name and the tables description
    expect([...pop.querySelectorAll('.ai-slash-skill-name')].map((n) => n.textContent)).toEqual([
      'tables',
      'tone',
    ])
    const active = () =>
      pop.querySelector('[aria-selected="true"] .ai-slash-skill-name')?.textContent
    expect(active()).toBe('tables')

    press(dock, 'ArrowDown')
    expect(active()).toBe('tone')
    const event = press(dock, 'Enter')
    expect(event.defaultPrevented).toBe(true)
    // the pick's Enter must not also send the half-typed token
    expect(requests).toHaveLength(0)
    expect(dock.querySelector('textarea')!.value).toBe('')
    expect(dock.querySelector('.ai-slash-pop')).toBeNull()
    expect(dock.querySelector('.ai-skill-chip-name')?.textContent).toBe('tone')

    await settle()
    const ta = dock.querySelector('textarea')!
    expect(ta.selectionStart).toBe(0)
    // focus returns to the composer after the pick, but the timing is
    // environment-dependent in jsdom — wait for it instead of asserting a fixed tick
    await vi.waitFor(() => expect(document.activeElement).toBe(ta))
  })

  it('picking an already-attached skill keeps it attached and strips a mid-draft token', async () => {
    const dock = mountPanel()
    await settle()
    type(dock, '/to')
    await settle()
    press(dock, 'Enter')
    expect([...dock.querySelectorAll('.ai-skill-chip-name')].map((n) => n.textContent)).toEqual([
      'tables',
    ])

    type(dock, 'please /to')
    await settle()
    press(dock, 'Enter')
    // the `/to` token is struck, the draft text around it stays
    expect(dock.querySelector('textarea')!.value).toBe('please ')
    // attaching is add-only: re-picking an attached skill never detaches it
    expect([...dock.querySelectorAll('.ai-skill-chip-name')].map((n) => n.textContent)).toEqual([
      'tables',
    ])
    expect(requests).toHaveLength(0)
  })
})
