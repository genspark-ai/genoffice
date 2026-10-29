import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { AiPanel, type HtmlAiDeps } from '../src/renderer/ai/AiPanel'
import { applyAiPanelPrefs, DEFAULT_AI_PANEL_PREFS } from '@genoffice/ui'
import type { AiSettings } from '@genoffice/ai-provider'
import type { HtmlApi } from '../src/shared/ipc'
import { buildParseMap } from '../src/renderer/document/parse-map'

/**
 * The context variables an attached skill's body expands. The composer's send
 * is the only path that attaches skills, so the vars must describe the document
 * the user was looking at when they hit send: the file's name, and the selected
 * element as the redaction projection sees it (a span withheld from the AI
 * cannot ride along in ${SELECTION}). The IPC boundary is stubbed
 * (window.htmlApi), so the panel and the agent loop run for real; the expanded
 * body is asserted on the user message of the streamed request.
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

const SKILL = { name: 'vars', description: 'Expand every context variable' }
const SKILL_BODY =
  'In ${APP} (${LANGUAGE}) working on ${FILE_NAME}; selected: ${SELECTION}; today is ${TODAY}. Request: $ARGUMENTS'

/** the redaction mark withholds the span's content from everything the model sees */
const PAGE = '<main><p>Call <span data-gx-redact="phone">555-0100</span> now</p></main>'

type StreamedRequest = { requestId: string; messages: Array<{ role: string; text: string }> }

let requests: StreamedRequest[] = []
/** settings reads take one timer tick: a send whose selection moves mid-pre-run is observable */
let settingsDelayMs = 0

const htmlApi = {
  getAiSettings: () =>
    new Promise<AiSettings>((resolve) =>
      setTimeout(() => resolve({} as AiSettings), settingsDelayMs),
    ),
  setAiSettings: () => Promise.resolve(),
  onAiSettingsChanged: () => () => {},
  aiGskStatus: () => Promise.resolve({ loggedIn: true }),
  openAiModelSettings: () => Promise.resolve(),
  openSkillsSettings: () => Promise.resolve(),
  skillsList: () => Promise.resolve([SKILL]),
  skillText: (name: string) => Promise.resolve(name === SKILL.name ? SKILL_BODY : ''),
  setAiPanelPrefs: () => Promise.resolve(),
  onAiPanelPrefsChanged: () => () => {},
  getPathForFile: () => null,
  readAttachmentImage: () => Promise.resolve({ ok: false, error: 'stub' }),
  aiStream: (request: StreamedRequest) => {
    requests.push(request)
    return Promise.resolve()
  },
  aiStreamCancel: () => {},
  onAiStream: () => () => {},
} as unknown as HtmlApi

window.htmlApi = htmlApi

// the document the panel's deps expose; the tests move each piece between sends
let docText = PAGE
let filePath: string | null = null
let selectedSid: number | null = null

const mapOf = () => buildParseMap(docText, 1, null)
const sidOf = (tag: string) => mapOf().elements.find((e) => e.tag === tag)!.sid

function makeDeps(): HtmlAiDeps {
  return {
    access: {
      getText: () => docText,
      getVersion: () => 1,
      getMap: () => mapOf(),
      getLastManualVersion: () => 1,
      getFilePath: () => filePath,
      getSelectedSid: () => selectedSid,
      applyOps: () => ({ ok: true, ranges: [] }),
      replaceAll: () => {},
    },
    getSnapshot: () => ({ text: docText }),
    restoreSnapshot: () => {},
    onPrompt: () => {},
    onRunDone: () => {},
    clearHighlights: () => {},
    navigateTo: () => {},
    onBriefConfirmed: () => {},
    previewDraft: () => {},
  }
}

/** mount without a preset: nothing auto-sends, so the request count stays honest */
function mountPanel(): HTMLElement {
  requests = []
  settingsDelayMs = 0
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

/** flush the skillsList / settings promises and any pending render */
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

function press(dock: HTMLElement, key: string): void {
  act(() =>
    dock
      .querySelector('textarea')!
      .dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })),
  )
}

/** attach the skill through the composer's slash palette, then send the request */
async function sendWithSkill(dock: HTMLElement, request: string): Promise<void> {
  type(dock, '/vars')
  await settle()
  press(dock, 'Enter')
  type(dock, request)
  press(dock, 'Enter')
  await settle()
  await settle()
}

/** the model-facing user message of the last streamed request */
function sentUserText(): string {
  expect(requests).toHaveLength(1)
  const msgs = requests[0]!.messages
  return msgs.filter((m) => m.role === 'user').at(-1)!.text
}

describe('attached-skill context variables', () => {
  it('expands APP, LANGUAGE, FILE_NAME, SELECTION and $ARGUMENTS from the live document', async () => {
    filePath = 'D:\\Documents\\Q3 page.html'
    selectedSid = sidOf('p')
    const dock = mountPanel()
    await settle()

    await sendWithSkill(dock, 'make it pop')
    const text = sentUserText()
    expect(text).toContain('<skill name="vars">')
    expect(text).toContain('In HTML 编辑器 (中文) working on Q3 page.html')
    // the selected element's text through the projection: the withheld number never rides along
    expect(text).toContain('selected: Call {{phone}} now')
    expect(text).not.toContain('555-0100')
    expect(text).toContain('Request: make it pop')
    // every placeholder the panel can supply is expanded; none is left behind
    expect(text).not.toContain('${')
  })

  it('keeps FILE_NAME and SELECTION literal when the document has neither', async () => {
    filePath = null
    selectedSid = null
    const dock = mountPanel()
    await settle()

    await sendWithSkill(dock, 'start over')
    const text = sentUserText()
    expect(text).toContain('In HTML 编辑器 (中文) working on ${FILE_NAME}')
    expect(text).toContain('selected: ${SELECTION}')
    expect(text).toContain('Request: start over')
  })

  it('caps the selection at 500 characters', async () => {
    docText = `<p>${'word '.repeat(150)}</p>`
    filePath = null
    selectedSid = sidOf('p')
    const dock = mountPanel()
    await settle()

    await sendWithSkill(dock, 'trim it')
    const text = sentUserText()
    const from = text.indexOf('selected: ') + 'selected: '.length
    const selection = text.slice(from, text.indexOf('; today is', from))
    expect(selection).toHaveLength(501)
    expect(selection.endsWith('…')).toBe(true)
    expect(selection.startsWith('word word')).toBe(true)
  })

  it('freezes the selection at send time, not at run time', async () => {
    docText = PAGE
    filePath = null
    selectedSid = sidOf('p')
    const dock = mountPanel()
    await settle()

    // the settings read stalls the pre-run: the selection moves after send but before the run
    settingsDelayMs = 20
    type(dock, '/vars')
    await settle()
    press(dock, 'Enter')
    type(dock, 'rework it')
    press(dock, 'Enter')
    docText = '<main><p>A different paragraph</p></main>'
    selectedSid = sidOf('p')
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 40))
    })

    const text = sentUserText()
    // the skill quotes the selection the user saw when hitting send...
    expect(text).toContain('selected: Call {{phone}} now')
    // ...while the document context the loop reads live describes the new one
    expect(text).toContain('A different paragraph')
  })
})
