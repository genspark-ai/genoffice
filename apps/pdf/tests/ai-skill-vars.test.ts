/**
 * Attached-skill template variables: the panel expands a skill body's
 * $ARGUMENTS, positionals and ${...} context vars at send time from its own
 * state — the request the composer sends, the file name the app reports, the
 * selection cached at mouseup, the UI language — while absent context stays
 * literal instead of substituting an empty string. The expanded instruction is
 * asserted on the stream request the transport forwards to the main process.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { AiPanel } from '../src/renderer/ai/AiPanel'
import type { PdfAppDeps } from '../src/renderer/ai/tools'
import type { AgentMessage } from '@genoffice/agent-core'

// single-quoted on purpose: ${...} must reach the skill body as template text
const SKILL_BODY =
  'File: ${FILE_NAME}\n' +
  'App: ${APP}\n' +
  'Lang: ${LANGUAGE}\n' +
  'Date: ${TODAY}\n' +
  'Sel: ${SELECTION}\n' +
  'Ask: $ARGUMENTS\n' +
  'First: $0\n' +
  'Third: $2'

const SKILLS = [{ name: 'vars', description: 'Template variable probe' }]

type StreamChunk = { requestId: string; type: 'done'; stopReason?: string }
type StreamListener = (chunk: StreamChunk) => void

const streamListeners = new Set<StreamListener>()
let startedRequestId = ''

const aiStream = vi.fn((request: { requestId: string; messages: AgentMessage[] }) => {
  startedRequestId = request.requestId
  return Promise.resolve()
})

const pdfApi = {
  // a configured custom endpoint keeps the composer out of the setup gate
  getAiSettings: () => Promise.resolve({ customEndpoints: [{ baseUrl: 'http://x', model: 'm' }] }),
  setAiSettings: () => Promise.resolve(),
  onAiSettingsChanged: () => () => {},
  openAiModelSettings: () => Promise.resolve(),
  openSkillsSettings: () => Promise.resolve(),
  skillsList: () => Promise.resolve(SKILLS),
  skillText: (name: string) => Promise.resolve(name === 'vars' ? SKILL_BODY : ''),
  gskStatus: () => Promise.resolve({ loggedIn: false }),
  aiStream,
  aiStreamCancel: () => Promise.resolve(),
  onAiStream: (handler: StreamListener) => {
    streamListeners.add(handler)
    return () => streamListeners.delete(handler)
  },
  setAiPanelPrefs: () => Promise.resolve(),
  listPageImages: () => Promise.resolve([]),
  pageImageBytes: () => Promise.resolve(null),
  openSourcePath: () => Promise.resolve(),
}

const projectApi = {
  resolveChat: () => Promise.resolve({ projectId: 'p', chatId: 'c' }),
  appendChat: () => Promise.resolve(),
  loadChat: () => Promise.resolve([]),
  rebindChat: () => Promise.resolve(null),
}

/** the App capability surface, parameterized on the two sources the vars read */
const makeApi = ({
  fileName = () => 'Q3 report.pdf',
  selection = () => ({ page: 2, lastPage: 2, text: '  revenue \n climbed\tsharply  ' }),
}: {
  fileName?: () => string
  selection?: () => { page: number; lastPage: number; text: string } | null
}): PdfAppDeps =>
  ({
    doc: () => null,
    fileName,
    pageCount: () => 1,
    currentPage: () => 1,
    readOnly: () => false,
    ocrText: () => null,
    selection,
    pendingSummary: () => '',
    outline: () => null,
    searchIndex: () => null,
    isDeleted: () => false,
    gotoPage: () => true,
    addMarkup: () => {},
    annotationsOn: () => Promise.resolve({ threads: [], markups: [] }),
    annotationSummary: () => '',
    metadata: () => ({}),
    addNote: () => Promise.resolve(null),
    deletePages: () => Promise.resolve({ ok: true, pageCount: 1 }),
    rotatePages: () => Promise.resolve({ ok: true, pageCount: 1 }),
    setPageSize: () => Promise.resolve({ ok: true, pageCount: 1 }),
    cropPages: () => Promise.resolve({ ok: true, pageCount: 1 }),
    replacePages: () => Promise.resolve({ ok: true, pageCount: 1 }),
    extractPages: () => Promise.resolve({ ok: true, savedPath: '' }),
    splitPdf: () => Promise.resolve({ ok: true, savedDir: '', count: 0 }),
    splitPages: () => Promise.resolve({ ok: true, savedPath: '' }),
    mergePages: () => Promise.resolve({ ok: true, savedPath: '' }),
  }) as unknown as PdfAppDeps

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  Object.defineProperty(window, 'projectApi', { value: projectApi, configurable: true })
  Object.defineProperty(window, 'pdfApi', { value: pdfApi, configurable: true })
  Element.prototype.scrollTo = () => {}
  // jsdom's crypto may lack randomUUID (the IPC transport mints stream ids)
  const cryptoObj = globalThis.crypto as Crypto & { randomUUID?: () => string }
  if (typeof cryptoObj.randomUUID !== 'function') {
    Object.defineProperty(cryptoObj, 'randomUUID', {
      value: () => `test-${Math.random().toString(36).slice(2)}`,
      configurable: true,
    })
  }
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
  aiStream.mockClear()
})

const flush = async (): Promise<void> => {
  for (let i = 0; i < 12; i++) await Promise.resolve()
}

async function renderPanel(api: PdfAppDeps): Promise<HTMLDivElement> {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(
      createElement(AiPanel, {
        api,
        filePath: 'sample.pdf',
        onCollapse: () => {},
        preset: null,
      }),
    )
    await flush()
  })
  return container
}

const textarea = (): HTMLTextAreaElement =>
  container!.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!

/** replace the draft the way typing does (native setter + input event) */
const type = (text: string): void => {
  act(() => {
    const ta = textarea()
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    setter.call(ta, text)
    ta.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

const press = (key: string): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  act(() => textarea().dispatchEvent(event))
  return event
}

/** attach the probe skill through the slash palette, then send the draft */
const sendWithSkill = async (draft: string): Promise<void> => {
  type('/vars')
  await act(async () => {
    await flush()
  })
  press('Enter')
  type(draft)
  press('Enter')
  await act(async () => {
    await flush()
  })
}

/** settle the started run so the transport's silence watchdog timer clears */
const finishRun = async (): Promise<void> => {
  await act(async () => {
    for (const listener of [...streamListeners])
      listener({ requestId: startedRequestId, type: 'done' })
    await flush()
  })
}

const isUser = (m: AgentMessage): m is Extract<AgentMessage, { role: 'user' }> => m.role === 'user'

/** the last user message on the stream request: the model-facing instruction */
const sentUserText = (): string => {
  const req = aiStream.mock.calls.at(-1)?.[0]
  expect(req, 'the transport started a stream').toBeDefined()
  const user = [...req!.messages].reverse().find(isUser)
  expect(user, 'a user message carrying the instruction').toBeDefined()
  return user!.text
}

describe('AiPanel skill template variables', () => {
  it('expands $ARGUMENTS, positionals and the standard context vars at send time', async () => {
    await renderPanel(makeApi({}))
    await sendWithSkill('summarize "page two"')

    const text = sentUserText()
    expect(text).toContain('<skill name="vars">')
    expect(text).toContain('File: Q3 report.pdf')
    expect(text).toContain('App: PDF')
    expect(text).toContain('Lang: 中文')
    expect(text).toContain(`Date: ${new Date().toLocaleDateString('zh')}`)
    // whitespace-collapsed selection text, not the scope chip's display excerpt
    expect(text).toContain('Sel: revenue climbed sharply')
    expect(text).toContain('Ask: summarize "page two"')
    expect(text).toContain('First: summarize')
    // a position with no argument stays literal
    expect(text).toContain('Third: $2')
    // the user's own request closes the skill block (the loop's document
    // context follows it, so the message itself does not end here)
    expect(text).toContain('User request:\nsummarize "page two"')
    await finishRun()
  })

  it('keeps FILE_NAME and SELECTION literal when the app has none, instead of substituting empty', async () => {
    await renderPanel(makeApi({ fileName: () => '', selection: () => null }))
    await sendWithSkill('anything goes')

    const text = sentUserText()
    expect(text).toContain('File: ${FILE_NAME}')
    expect(text).toContain('Sel: ${SELECTION}')
    // the always-present context still expands
    expect(text).toContain('App: PDF')
    expect(text).toContain('Lang: 中文')
    expect(text).toContain('Ask: anything goes')
    await finishRun()
  })
})
