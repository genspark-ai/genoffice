/**
 * AiPanel reply-copy toolbar: completed assistant messages render a copy button
 * that routes the reply through copyMarkdownRich (rich clipboard), and the
 * checkmark flip is keyed to the message index.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
// copyMarkdownRich is the interaction target; vi.mock hoists above imports, so
// the factory must reference the fn through the hoisted holder.
vi.mock('@genoffice/ui', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@genoffice/ui')>()
  return { ...actual, copyMarkdownRich: (text: string) => copyMock(text) }
})

import { AiPanel } from '../src/renderer/ai/AiPanel'
import type { PdfAppDeps } from '../src/renderer/ai/tools'

const copyMock = vi.fn(() => Promise.resolve())

/** minimal App capability surface: the render path only touches selection/gotoPage */
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

const projectApi = {
  resolveChat: () => Promise.resolve({ projectId: 'p', chatId: 'c' }),
  appendChat: () => Promise.resolve(),
  // one completed (non-streaming) assistant reply: exactly what the toolbar keys on
  loadChat: () => Promise.resolve([{ role: 'assistant', text: '**Answer**: 42 rows matched' }]),
}

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  Object.defineProperty(window, 'projectApi', { value: projectApi, configurable: true })
})
// jsdom implements neither scrollTo nor panel layout; the panel auto-scrolls
// the chat on every chat/busy change
beforeAll(() => {
  Element.prototype.scrollTo = () => {}
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
  vi.useRealTimers()
  copyMock.mockClear()
})

async function renderPanel() {
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
    await Promise.resolve()
  })
  return container
}

describe('AiPanel copy toolbar', () => {
  it('copy button routes the reply text through copyMarkdownRich', async () => {
    const panel = await renderPanel()
    // a completed assistant reply with text carries the toolbar
    const btn = panel.querySelector<HTMLButtonElement>('.ai-msg-assistant .ai-msg-tool-btn')
    expect(btn, 'toolbar copy button on an assistant message').not.toBeNull()
    // the panel reads the shipped zh strings table (default locale)
    expect(btn!.getAttribute('aria-label')).toBe('复制回复')

    await act(async () => {
      btn!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })
    expect(copyMock).toHaveBeenCalledTimes(1)
    const arg = copyMock.mock.calls[0]![0]
    expect(typeof arg).toBe('string')
    expect(arg.length).toBeGreaterThan(0)
  })

  it('flips to the checkmark after copy and back after 1.2s', async () => {
    const panel = await renderPanel()
    const btn = panel.querySelector<HTMLButtonElement>('.ai-msg-assistant .ai-msg-tool-btn')!
    vi.useFakeTimers()
    await act(async () => {
      btn.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await vi.advanceTimersByTimeAsync(0)
    })
    // flipped: the checkmark polyline svg replaced the copy-plate svg
    expect(btn.querySelector('polyline')).not.toBeNull()
    await act(async () => {
      await vi.advanceTimersByTimeAsync(1200)
    })
    expect(btn.querySelector('polyline')).toBeNull()
  })
})
