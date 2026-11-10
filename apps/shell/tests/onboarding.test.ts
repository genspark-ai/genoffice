/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Onboarding } from '../src/renderer/src/Onboarding'
import { LocaleProvider } from '../src/renderer/src/locale'

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

function renderOnboarding(onDone: () => Promise<boolean>): void {
  act(() => {
    root.render(
      createElement(LocaleProvider, { initial: 'en' }, createElement(Onboarding, { onDone })),
    )
  })
}

async function click(selector: string): Promise<void> {
  const button = host.querySelector<HTMLButtonElement>(selector)
  expect(button).not.toBeNull()
  await act(async () => {
    button!.click()
    await Promise.resolve()
  })
}

describe('analytics consent in onboarding', () => {
  it('explains default-on analytics without an onboarding control', () => {
    renderOnboarding(vi.fn(async () => true))
    expect(host.querySelector('[role="switch"]')).toBeNull()
    expect(host.textContent).toContain('Enabled by default')
    expect(host.textContent).toContain('Settings → General')
  })

  it('lets Skip and Escape finish without changing analytics', async () => {
    const onDone = vi.fn(async () => true)
    renderOnboarding(onDone)

    await click('.onb-skip')
    expect(onDone).toHaveBeenLastCalledWith()

    onDone.mockClear()
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
      await Promise.resolve()
    })
    expect(onDone).toHaveBeenLastCalledWith()
  })

  it('finishes the final slide without an analytics choice', async () => {
    const onDone = vi.fn(async () => true)
    renderOnboarding(onDone)

    await click('.onb-next')
    await click('.onb-next')
    await click('.onb-next')
    expect(onDone).toHaveBeenCalledWith()
  })
})
