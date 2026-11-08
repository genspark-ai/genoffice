/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { Root } from 'react-dom/client'
import type { HomeApi } from '../src/shared/home-api'
import { LocaleProvider } from '../src/renderer/src/locale'
import { SettingsModal } from '../src/renderer/src/SettingsModal'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
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

async function renderSettings(value = 'last-used', save = vi.fn(async (next: string) => next)) {
  window.aiOffice = {
    getPdfDefaultZoom: async () => value,
    setPdfDefaultZoom: save,
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
          target: { section: 'general' },
        }),
      ),
    )
  })
  return save
}

const trigger = () => host.querySelector<HTMLButtonElement>('[aria-label="Default PDF zoom"]')
async function pick(label: string) {
  await act(async () => trigger()!.click())
  const option = [...document.querySelectorAll<HTMLButtonElement>('[role="option"]')].find(
    (button) => button.textContent === label,
  )
  expect(option).toBeDefined()
  await act(async () => option!.click())
}

describe('PDF opening zoom in General settings', () => {
  it('loads the saved preference and persists a new choice', async () => {
    const save = await renderSettings('width')
    expect(trigger()).not.toBeNull()
    expect(trigger()?.textContent).toContain('Fit width')
    await pick('Fit page')
    expect(save).toHaveBeenCalledWith('page')
    expect(trigger()?.textContent).toContain('Fit page')
  })

  it('keeps the saved selection and reports a failed write', async () => {
    const save = vi.fn(async () => {
      throw new Error('disk full')
    })
    await renderSettings('width', save)
    await pick('Fit page')
    expect(trigger()?.textContent).toContain('Fit width')
    expect(host.querySelector('[role="alert"]')?.textContent).toContain('Could not load or save')
  })

  it('keeps the selection until persistence succeeds and prevents overlapping writes', async () => {
    let finish!: (value: string) => void
    const save = vi.fn(
      () =>
        new Promise<string>((resolve) => {
          finish = resolve
        }),
    )
    await renderSettings('last-used', save)
    await pick('Fit page')
    expect(trigger()?.textContent).toContain('Last used zoom')
    expect(trigger()?.disabled).toBe(true)
    await act(async () => finish('page'))
    expect(trigger()?.textContent).toContain('Fit page')
    expect(trigger()?.disabled).toBe(false)
  })
})
