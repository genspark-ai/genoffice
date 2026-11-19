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
})

async function click(button: HTMLButtonElement): Promise<void> {
  await act(async () => {
    button.click()
    await Promise.resolve()
  })
}

describe('Settings analytics consent', () => {
  it('starts on and changes only after persistence succeeds', async () => {
    const persist = vi
      .fn<(enabled: boolean) => Promise<boolean>>()
      .mockResolvedValueOnce(false)
      .mockResolvedValueOnce(true)
    window.aiOffice = {
      getTheme: async () => 'system',
      getDefaultSaveDir: async () => '',
      getAnalyticsEnabled: async () => true,
      setAnalyticsEnabled: persist,
      getAiPanelPrefs: async () => ({ fontSize: 'default', spellcheck: true }),
      setAiPanelPrefs: async (patch) => ({ fontSize: 'default', spellcheck: true, ...patch }),
      getUpdateChannel: async () => 'stable',
      getAppVersion: async () => '1.0.0',
      githubStars: async () => null,
    } as unknown as HomeApi

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
    await click(general!)

    const consent = host.querySelector<HTMLButtonElement>(
      '.set-switch[aria-label="Send anonymous usage statistics"]',
    )
    expect(consent?.getAttribute('aria-checked')).toBe('true')

    await click(consent!)
    expect(persist).toHaveBeenLastCalledWith(false)
    expect(consent?.getAttribute('aria-checked')).toBe('true')

    await click(consent!)
    expect(consent?.getAttribute('aria-checked')).toBe('false')
  })
})

describe('Settings AutoSave default', () => {
  it('reads the shell default and persists a flip', async () => {
    const persist = vi.fn<(on: boolean) => Promise<void>>().mockResolvedValue(undefined)
    window.aiOffice = {
      getTheme: async () => 'system',
      getDefaultSaveDir: async () => '',
      getAnalyticsEnabled: async () => true,
      setAnalyticsEnabled: async () => true,
      getAutoSaveDefault: async () => ({ on: true, updatedAt: 1 }),
      setAutoSaveDefault: persist,
      getUpdateChannel: async () => 'stable',
      getAppVersion: async () => '1.0.0',
      githubStars: async () => null,
    } as unknown as HomeApi

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
    await click(general!)

    const toggle = host.querySelector<HTMLButtonElement>(
      '.set-switch[aria-label="Auto-save all documents"]',
    )
    expect(toggle?.getAttribute('aria-checked')).toBe('true')

    await click(toggle!)
    expect(persist).toHaveBeenLastCalledWith(false)
    expect(toggle?.getAttribute('aria-checked')).toBe('false')
  })
})
