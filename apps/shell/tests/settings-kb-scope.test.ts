/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { HomeApi } from '../src/shared/home-api'
import { LocaleProvider } from '../src/renderer/src/locale'
import { SettingsModal } from '../src/renderer/src/SettingsModal'

/**
 * The settings center's Ask AI corpus-scope group (General section): the
 * two-option radio reads the stored kbScope through the bridge and writes a
 * pick through setKbScope, taking the returned value as the truth. A bridge
 * without the pair (an older preload) hides the group rather than showing a
 * toggle that cannot work.
 */

const actEnvironment = globalThis as typeof globalThis & {
  IS_REACT_ACT_ENVIRONMENT?: boolean
}
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  window.aiOffice = undefined as unknown as HomeApi
})

function stubApi(overrides: Record<string, unknown> = {}): HomeApi {
  return {
    getTheme: async () => 'system',
    getDefaultSaveDir: async () => '',
    getAnalyticsEnabled: async () => true,
    setAnalyticsEnabled: async () => true,
    getUpdateChannel: async () => 'stable',
    getAppVersion: async () => '1.0.0',
    githubStars: async () => null,
    ...overrides,
  } as unknown as HomeApi
}

async function openGeneral(api: HomeApi): Promise<void> {
  window.aiOffice = api
  await act(async () => {
    root.render(
      createElement(
        LocaleProvider,
        { initial: 'en' },
        createElement(SettingsModal, {
          status: null,
          loggingOut: false,
          loginWaiting: false,
          loginUrl: null,
          urlCopied: false,
          onOpenLoginUrl: vi.fn(),
          onCopyLoginUrl: vi.fn(),
          onClose: vi.fn(),
          onLogin: vi.fn(),
          onLogout: vi.fn(),
        }),
      ),
    )
    await Promise.resolve()
  })
  const general = Array.from(host.querySelectorAll<HTMLButtonElement>('.set-nav-item')).find(
    (button) => button.textContent?.includes('General'),
  )
  expect(general).toBeDefined()
  await act(async () => {
    general!.click()
    await Promise.resolve()
  })
  // let the mount-time reads land and re-render
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

/** the two radio buttons of the scope group, in display order */
const scopeButtons = () =>
  Array.from(
    host.querySelectorAll<HTMLButtonElement>(
      '[role="radiogroup"][aria-label="Ask AI corpus scope"] [role="radio"]',
    ),
  )

describe('Settings — Ask AI corpus scope', () => {
  it('renders the two options from the stored scope and writes a pick', async () => {
    const setKbScope = vi
      .fn<(scope: 'indexed' | 'starred') => Promise<'indexed' | 'starred'>>()
      .mockResolvedValue('starred')
    await openGeneral({
      ...stubApi(),
      getKbScope: async () => 'indexed',
      setKbScope,
    } as unknown as HomeApi)
    const buttons = scopeButtons()
    expect(buttons.map((b) => b.textContent)).toEqual(['All indexed files', 'Starred only'])
    expect(buttons[0]!.getAttribute('aria-checked')).toBe('true')
    expect(buttons[1]!.getAttribute('aria-checked')).toBe('false')

    await act(async () => {
      buttons[1]!.click()
      await Promise.resolve()
    })
    expect(setKbScope).toHaveBeenCalledWith('starred')
    // the row re-renders from the returned value, not the optimistic pick
    expect(scopeButtons()[1]!.getAttribute('aria-checked')).toBe('true')
  })

  it('a refused write keeps the persisted scope', async () => {
    const setKbScope = vi
      .fn<(scope: 'indexed' | 'starred') => Promise<'indexed' | 'starred'>>()
      .mockRejectedValue(new Error('settings file unwritable'))
    await openGeneral({
      ...stubApi(),
      getKbScope: async () => 'indexed',
      setKbScope,
    } as unknown as HomeApi)
    await act(async () => {
      scopeButtons()[1]!.click()
      await Promise.resolve()
    })
    expect(setKbScope).toHaveBeenCalledWith('starred')
    // nothing moved: the group still shows the scope that is on disk
    expect(scopeButtons()[0]!.getAttribute('aria-checked')).toBe('true')
    expect(scopeButtons()[1]!.getAttribute('aria-checked')).toBe('false')
  })

  it('hides the group when the bridge cannot serve the scope', async () => {
    await openGeneral(stubApi())
    expect(host.querySelector('[aria-label="Ask AI corpus scope"]')).toBeNull()
    expect(host.textContent).not.toContain('All indexed files')
  })
})
