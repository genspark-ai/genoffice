// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import type { ZoteroLocalItem } from '../src/shared/ipc'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { ZoteroQuickCite } from '../src/renderer/components/ZoteroQuickCite'
import { ZOTERO_SEARCH_DEBOUNCE_MS } from '../src/renderer/components/zotero-items'
import { runMenuCommand, type CommandHost } from '../src/renderer/commands/menu-commands'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

const item = (key: string, title: string): ZoteroLocalItem => ({
  key,
  libraryId: 7,
  title,
  creators: 'Smith',
  year: '2020',
  itemType: 'journalArticle',
})
const ITEMS = [item('ABCD2345', 'First paper'), item('WXYZ6789', 'Second paper')]

let root: Root
let container: HTMLElement

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
})

afterEach(() => {
  act(() => root.unmount())
  container.remove()
  vi.useRealTimers()
  vi.unstubAllGlobals()
})

function type(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
  act(() => {
    setter.call(input, value)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function key(input: HTMLInputElement, keyName: string, init: KeyboardEventInit = {}) {
  act(() => {
    input.dispatchEvent(new KeyboardEvent('keydown', { key: keyName, bubbles: true, ...init }))
  })
}

async function settle() {
  await act(async () => {
    vi.advanceTimersByTime(ZOTERO_SEARCH_DEBOUNCE_MS + 1)
  })
  await act(async () => {})
}

describe('Quick Cite bar', () => {
  function mount(onCite = vi.fn().mockResolvedValue(null), onClose = vi.fn()) {
    vi.useFakeTimers()
    const zoteroSearch = vi.fn().mockResolvedValue({ ok: true, value: ITEMS })
    vi.stubGlobal('desktop', { zoteroSearch })
    act(() => root.render(createElement(ZoteroQuickCite, { onCite, onClose })))
    const input = container.querySelector<HTMLInputElement>(
      '[data-testid="zotero-quick-cite-input"]',
    )!
    return { input, onCite, onClose, zoteroSearch }
  }

  it('searches as the user types and cites the highlighted item on Enter', async () => {
    const { input, onCite, onClose, zoteroSearch } = mount()
    expect(document.activeElement).toBe(input)
    type(input, 'smith')
    await settle()
    expect(zoteroSearch).toHaveBeenLastCalledWith('smith')
    expect(container.querySelectorAll('[role="option"]')).toHaveLength(2)
    key(input, 'ArrowDown')
    await act(async () => key(input, 'Enter'))
    expect(onCite).toHaveBeenCalledWith(['WXYZ6789'])
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('builds one citation of several items with the command key', async () => {
    const { input, onCite } = mount()
    type(input, 'paper')
    await settle()
    key(input, 'Enter', { metaKey: true })
    expect(container.querySelectorAll('.zotero-quick-cite-chip')).toHaveLength(1)
    key(input, 'ArrowDown')
    await act(async () => key(input, 'Enter'))
    expect(onCite).toHaveBeenCalledWith(['ABCD2345', 'WXYZ6789'])
  })

  it('keeps the bar open with the message when the citation could not go in', async () => {
    const { input, onClose } = mount(vi.fn().mockResolvedValue('zoteroLocalApiOff'))
    type(input, 'smith')
    await settle()
    await act(async () => key(input, 'Enter'))
    expect(onClose).not.toHaveBeenCalled()
    expect(container.textContent).toContain(t('zoteroLocalApiOff'))
  })

  it('ignores Enter while the list still answers the previous search', async () => {
    const { input, onCite, zoteroSearch } = mount()
    type(input, 'smith')
    await settle()
    zoteroSearch.mockReturnValue(new Promise(() => {}))
    type(input, 'lee 2021')
    await act(async () => key(input, 'Enter'))
    expect(onCite).not.toHaveBeenCalled()
  })

  it('gives the search box back after a citation that could not go in', async () => {
    const { input } = mount(vi.fn().mockResolvedValue('zoteroConnectionError'))
    type(input, 'smith')
    await settle()
    await act(async () => key(input, 'Enter'))
    expect(input.disabled).toBe(false)
    expect(document.activeElement).toBe(input)
  })

  it('closes on Escape without citing', () => {
    const { input, onCite, onClose } = mount()
    key(input, 'Escape')
    expect(onClose).toHaveBeenCalledTimes(1)
    expect(onCite).not.toHaveBeenCalled()
  })
})

describe('Zotero tab', () => {
  it('runs the Quick Cite, Library and Refresh on Open commands', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] },
    })
    const onCommand = vi.fn()
    act(() =>
      root.render(
        createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
          onCommand,
          showZotero: true,
          zoteroAutoRefresh: true,
        }),
      ),
    )
    const zoteroTab = [...container.querySelectorAll<HTMLButtonElement>('.ribbon-tab')].find(
      (button) => button.textContent === 'Zotero',
    )!
    act(() => zoteroTab.click())
    const click = (id: string) =>
      act(() => container.querySelector<HTMLButtonElement>(`[data-testid="${id}"]`)!.click())
    click('zotero-quick-cite')
    click('zotero-library')
    click('zotero-auto-refresh')
    expect(onCommand.mock.calls.map(([id]) => id)).toEqual([
      'zotero-quick-cite',
      'zotero-library',
      'zotero-auto-refresh',
    ])
    expect(
      container.querySelector('[data-testid="zotero-auto-refresh"]')!.getAttribute('aria-pressed'),
    ).toBe('true')
    editor.destroy()
  })

  it('opens Quick Cite and the library only with a document', () => {
    const host = {
      editor: {} as CommandHost['editor'],
      hasDoc: true,
      openZoteroQuickCite: vi.fn(),
      toggleZoteroLibrary: vi.fn(),
      toggleZoteroAutoRefresh: vi.fn(),
    } as unknown as CommandHost
    runMenuCommand(host, 'zotero-quick-cite')
    runMenuCommand(host, 'zotero-library')
    runMenuCommand(host, 'zotero-auto-refresh')
    expect(host.openZoteroQuickCite).toHaveBeenCalledTimes(1)
    expect(host.toggleZoteroLibrary).toHaveBeenCalledTimes(1)
    expect(host.toggleZoteroAutoRefresh).toHaveBeenCalledTimes(1)
    const closed = { ...host, hasDoc: false, openZoteroQuickCite: vi.fn() } as CommandHost
    runMenuCommand(closed, 'zotero-quick-cite')
    expect(closed.openZoteroQuickCite).not.toHaveBeenCalled()
  })
})
