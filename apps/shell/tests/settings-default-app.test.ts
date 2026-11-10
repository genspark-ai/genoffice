/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { DefaultAppStatus, HomeApi } from '../src/shared/home-api'
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

async function openGeneral(api: Partial<HomeApi>): Promise<void> {
  window.aiOffice = {
    getTheme: async () => 'system',
    getDefaultSaveDir: async () => '',
    getAnalyticsEnabled: async () => true,
    setAnalyticsEnabled: async () => true,
    getAiPanelPrefs: async () => ({ fontSize: 'default', spellcheck: true }),
    setAiPanelPrefs: async (patch) => ({ fontSize: 'default', spellcheck: true, ...patch }),
    getUpdateChannel: async () => 'stable',
    getAppVersion: async () => '1.0.0',
    githubStars: async () => null,
    ...api,
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
}

function row(): HTMLElement | null {
  return (
    Array.from(host.querySelectorAll<HTMLElement>('.set-field')).find((el) =>
      el.textContent?.includes('Default app for Office documents'),
    ) ?? null
  )
}

describe('Settings default-app row', () => {
  it('is hidden when the platform/build cannot manage defaults', async () => {
    await openGeneral({
      getDefaultAppStatus: async () => ({ state: 'unsupported', others: [], manualOnly: false }),
    })
    expect(row()).toBeNull()
  })

  it('names the current owner and flips to "already default" after claiming', async () => {
    const other: DefaultAppStatus = { state: 'other', others: ['WPS Office'], manualOnly: false }
    const claimed: DefaultAppStatus = { state: 'default', others: [], manualOnly: false }
    const set = vi.fn(async () => claimed)
    await openGeneral({ getDefaultAppStatus: async () => other, setDefaultApp: set })

    const field = row()
    expect(field?.textContent).toContain('Current default: WPS Office')
    const button = field!.querySelector<HTMLButtonElement>('button')!
    expect(button.textContent).toBe('Make default')
    expect(button.disabled).toBe(false)

    await click(button)
    expect(set).toHaveBeenCalledTimes(1)
    expect(row()?.textContent).toContain('GenOffice is already the default.')
    expect(row()!.querySelector<HTMLButtonElement>('button')!.disabled).toBe(true)
  })

  it('keeps the claim button live when a type has no handler yet', async () => {
    await openGeneral({
      getDefaultAppStatus: async () => ({ state: 'other', others: [], manualOnly: false }),
    })
    const field = row()!
    expect(field.textContent).toContain('Open .docx, .xlsx and .pptx files in GenOffice')
    expect(field.querySelector<HTMLButtonElement>('button')!.disabled).toBe(false)
  })

  it('shows a failure hint when the claim did not stick', async () => {
    const other: DefaultAppStatus = { state: 'other', others: ['WPS Office'], manualOnly: false }
    await openGeneral({ getDefaultAppStatus: async () => other, setDefaultApp: async () => other })
    await click(row()!.querySelector<HTMLButtonElement>('button')!)
    expect(row()?.textContent).toContain('Could not change it.')
  })

  it('offers the system settings page on Windows', async () => {
    const set = vi.fn(
      async () => ({ state: 'other', others: ['Word'], manualOnly: true }) as DefaultAppStatus,
    )
    await openGeneral({
      getDefaultAppStatus: async () => ({ state: 'other', others: ['Word'], manualOnly: true }),
      setDefaultApp: set,
    })
    const button = row()!.querySelector<HTMLButtonElement>('button')!
    expect(button.textContent).toBe('Open system settings')
    await click(button)
    expect(set).toHaveBeenCalledTimes(1)
    expect(row()?.textContent).not.toContain('Could not change it.')
  })
})
