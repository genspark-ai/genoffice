/**
 * Slash palette wiring: typing `/` in the composer opens the shared skill
 * listbox above the textarea; Enter picks — the skill joins the attached
 * chips, the `/query` token is struck from the draft, and the consumed key
 * never falls through to the composer's send.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { AiPanel } from '../src/renderer/ai/AiPanel'
import type { PdfAppDeps } from '../src/renderer/ai/tools'

const SKILLS = [
  { name: 'tables', description: 'Header rows and totals' },
  { name: 'tone', description: 'Write plainly' },
]

const pdfApi = {
  getAiSettings: () => Promise.resolve({}),
  setAiSettings: () => Promise.resolve(),
  onAiSettingsChanged: () => () => {},
  openAiModelSettings: () => Promise.resolve(),
  openSkillsSettings: () => Promise.resolve(),
  skillsList: () => Promise.resolve(SKILLS),
  skillText: () => Promise.resolve(''),
  gskStatus: () => Promise.resolve({ loggedIn: false }),
  aiStream: vi.fn((): Promise<void> => Promise.resolve()),
  aiStreamCancel: () => Promise.resolve(),
  onAiStream: () => () => {},
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

const api = {
  doc: () => null,
  fileName: () => 'sample.pdf',
  pageCount: () => 1,
  currentPage: () => 1,
  readOnly: () => false,
  ocrText: () => null,
  selection: () => null,
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
} as unknown as PdfAppDeps

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  Object.defineProperty(window, 'projectApi', { value: projectApi, configurable: true })
  Object.defineProperty(window, 'pdfApi', { value: pdfApi, configurable: true })
  Element.prototype.scrollTo = () => {}
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
  pdfApi.aiStream.mockClear()
})

const flush = async (): Promise<void> => {
  for (let i = 0; i < 6; i++) await Promise.resolve()
}

async function renderPanel(): Promise<HTMLDivElement> {
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
  await act(async () => {
    await flush()
  })
  return container
}

const textarea = (): HTMLTextAreaElement =>
  container!.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!

/** replace the draft the way typing does (native setter + input event) */
const type = (text: string, caret = text.length): void => {
  act(() => {
    const ta = textarea()
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    setter.call(ta, text)
    ta.setSelectionRange(caret, caret)
    ta.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

const press = (key: string): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  act(() => textarea().dispatchEvent(event))
  return event
}

const chipNames = (): string[] =>
  [...container!.querySelectorAll('.ai-skill-chip-name')].map((n) => n.textContent)

describe('AiPanel slash palette', () => {
  it('opens the skill listbox on a leading slash; Enter picks, attaches and strips the token without sending', async () => {
    await renderPanel()
    type('/t')
    await act(async () => {
      await flush()
    })

    const pop = container!.querySelector<HTMLElement>('.ai-slash-pop')!
    expect(pop, 'the palette above the composer').not.toBeNull()
    expect(pop.getAttribute('role')).toBe('listbox')
    expect([...pop.querySelectorAll('.ai-slash-skill-name')].map((n) => n.textContent)).toEqual([
      'tables',
      'tone',
    ])

    const event = press('Enter')
    expect(event.defaultPrevented).toBe(true)
    expect(textarea().value, 'the /query token is struck').toBe('')
    expect(chipNames()).toEqual(['tables'])
    // the palette ate the Enter: the composer never sent the half-typed token
    expect(container!.querySelector('.ai-msg-user')).toBeNull()
    expect(container!.querySelector('.ai-typing')).toBeNull()
    expect(pdfApi.aiStream).not.toHaveBeenCalled()
    expect(container!.querySelector('.ai-slash-pop')).toBeNull()
  })

  it('keeps an already-attached skill attached when the same one is picked again', async () => {
    await renderPanel()
    type('/tables')
    await act(async () => {
      await flush()
    })
    press('Enter')
    type('/tables')
    await act(async () => {
      await flush()
    })
    press('Enter')
    expect(chipNames()).toEqual(['tables'])
    expect(textarea().value).toBe('')
  })

  it('strikes a token mid-draft and restores the caret where it stood', async () => {
    await renderPanel()
    type('please /tone')
    await act(async () => {
      await flush()
    })
    press('Enter')
    expect(textarea().value).toBe('please ')
    // the pick refocuses the textarea and parks the caret at the token start
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 50))
    })
    expect(document.activeElement).toBe(textarea())
    expect(textarea().selectionStart).toBe(7)
    expect(textarea().selectionEnd).toBe(7)
  })
})
