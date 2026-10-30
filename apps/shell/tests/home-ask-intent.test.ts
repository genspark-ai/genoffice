// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { Home } from '../src/renderer/src/Home'
import { LocaleProvider } from '../src/renderer/src/locale'
import type { RecentEntry } from '../src/shared/home-api'

/**
 * The ask intent's two entry routes into Home: the command palette's Ask AI
 * fallback row, and the `home:deep-link` channel (genoffice://ask?q=…). Both
 * must land in the ask view, but only the palette's user-typed text submits
 * itself — a deep link's web-supplied query parks in the composer and waits
 * for an explicit Enter, so a web page can never inject a prompt that runs.
 * The corpus mock is empty on purpose: the submission then resolves to the
 * panel's guidance answer synchronously after the corpus check, so no AI
 * transport takes part in these tests.
 */

let root: Root
let container: HTMLDivElement
let deepLinkHandler: ((payload: unknown) => void) | null
let paletteHandler: (() => void) | null

const files: RecentEntry[] = [
  {
    path: '/fixtures/brief.docx',
    name: 'brief.docx',
    ext: 'docx',
    mtimeMs: 1,
    sizeBytes: 10,
    starred: false,
  },
]

beforeEach(async () => {
  HTMLElement.prototype.scrollTo = vi.fn()
  // this suite starts from the list view: the intent routes must switch away from it
  localStorage.setItem('home.view', 'recent')
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  vi.stubGlobal(
    'fetch',
    vi.fn(() => Promise.reject(new Error('No network in ask-intent tests'))),
  )
  deepLinkHandler = null
  paletteHandler = null
  window.aiOffice = {
    folderRoots: async () => [],
    onFolderChanged: () => () => {},
    recents: async () => ({ entries: files, total: files.length, totalAll: files.length }),
    starred: async () => ({ entries: [], total: 0, totalAll: 0 }),
    starredGroups: async () => [],
    accountStatus: async () => ({ loggedIn: false }),
    onOpenSettings: () => () => {},
    onAccountLogin: () => () => {},
    kbList: async () => [],
    aiGetSettings: async () => defaultAiSettings(),
    // logged in so the composer's setup gate stays open: an explicit Enter
    // must be able to send the parked deep-link draft
    aiGskStatus: async () => ({ loggedIn: true }),
    onAiSettingsChanged: () => () => {},
    onDeepLink: (handler: (payload: unknown) => void) => {
      deepLinkHandler = handler
      return () => {}
    },
    onPaletteOpen: (handler: () => void) => {
      paletteHandler = handler
      return () => {}
    },
    openPath: async () => {},
  } as unknown as Window['aiOffice']
  container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
  await act(async () =>
    root.render(createElement(LocaleProvider, { initial: 'en', children: createElement(Home) })),
  )
})

afterEach(async () => {
  await act(async () => root.unmount())
  container.remove()
  vi.unstubAllGlobals()
  localStorage.clear()
})

/** flush the mounted component plus the async submission chain (kbList, state) */
async function flush(times = 6): Promise<void> {
  await act(async () => {
    for (let i = 0; i < times; i++) await Promise.resolve()
  })
}

/** the questions the ask panel shows, in thread order */
const askedMessages = (): string[] =>
  [...container.querySelectorAll('.kb-ask-msg.kb-ask-user')].map((el) => el.textContent ?? '')

/** the ask panel's composer draft */
const composerDraft = (): string =>
  container.querySelector<HTMLTextAreaElement>('.kb-ask textarea')?.value ?? ''

describe('Home ask intent — deep link', () => {
  it('home:deep-link {view:ask, query} opens the ask view and parks the query in the composer', async () => {
    expect(deepLinkHandler).not.toBeNull()
    await act(async () => deepLinkHandler!({ view: 'ask', query: 'hello from a link' }))
    await flush()
    expect(container.querySelector('.kb-ask')).not.toBeNull()
    // web-supplied text never submits itself: it waits in the draft for Enter
    expect(askedMessages()).toEqual([])
    expect(composerDraft()).toBe('hello from a link')
  })

  it('the parked deep-link query is sent only by an explicit Enter', async () => {
    await act(async () => deepLinkHandler!({ view: 'ask', query: 'hello from a link' }))
    await flush()
    expect(askedMessages()).toEqual([])
    const composer = container.querySelector<HTMLTextAreaElement>('.kb-ask textarea')!
    await act(async () => {
      composer.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
      await Promise.resolve()
    })
    await flush()
    expect(askedMessages()).toEqual(['hello from a link'])
    expect(composerDraft()).toBe('')
  })

  it('a deep link without a query only opens the ask view', async () => {
    await act(async () => deepLinkHandler!({ view: 'ask' }))
    await flush()
    expect(container.querySelector('.kb-ask')).not.toBeNull()
    expect(askedMessages()).toEqual([])
    expect(composerDraft()).toBe('')
  })

  it('a deep link for another view is ignored', async () => {
    await act(async () => deepLinkHandler!({ view: 'recent' }))
    await flush()
    expect(container.querySelector('.kb-ask')).toBeNull()
    expect(container.querySelector('.content')?.textContent).toContain('brief.docx')
  })
})

describe('Home ask intent — palette fallback', () => {
  it('the palette Ask AI row routes its query into the ask panel', async () => {
    expect(paletteHandler).not.toBeNull()
    await act(async () => paletteHandler!())
    const input = container.querySelector<HTMLInputElement>('.cmd-input')
    expect(input).not.toBeNull()
    await act(async () => {
      const proto = window.HTMLInputElement.prototype
      Object.getOwnPropertyDescriptor(proto, 'value')?.set?.call(input, 'zzz palette question')
      input!.dispatchEvent(new Event('input', { bubbles: true }))
      await Promise.resolve()
      await Promise.resolve()
    })
    const row = [...container.querySelectorAll<HTMLButtonElement>('.cmd-row')].find((el) =>
      el.textContent?.includes('Ask AI'),
    )
    expect(row).toBeTruthy()
    await act(async () => row!.click())
    await flush()
    // the palette closed and the question is a message in the ask thread
    expect(container.querySelector('.cmd-overlay')).toBeNull()
    expect(container.querySelector('.kb-ask')).not.toBeNull()
    expect(askedMessages()).toEqual(['zzz palette question'])
  })
})
