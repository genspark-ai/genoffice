/**
 * StampDialog / SignatureDialog accessibility: dialog semantics, Escape/backdrop
 * close, and focus-first-field with return-focus on unmount.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { StampDialog } from '../src/renderer/StampDialog'
import { SignatureDialog } from '../src/renderer/SignatureDialog'
import type { TFunc } from '../src/renderer/i18n/locale'

// Identity translator: keeps the test independent of locale data
const t = ((key: string) => key) as unknown as TFunc

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
})

async function renderDialog(onCancel: () => void) {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(createElement(StampDialog, { t, onCancel, onApply: () => {} }))
    await Promise.resolve()
  })
  return container
}

describe('StampDialog a11y', () => {
  it('exposes dialog semantics labelled by the existing title key', async () => {
    const dialog = await renderDialog(() => {})
    const node = dialog.querySelector('[role="dialog"]')
    expect(node).not.toBeNull()
    expect(node!.getAttribute('aria-modal')).toBe('true')
    expect(node!.getAttribute('aria-label')).toBe('stampTitle')
  })

  it('closes on Escape and on backdrop click', async () => {
    const onCancel = vi.fn()
    const dialog = await renderDialog(onCancel)
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
      await Promise.resolve()
    })
    expect(onCancel).toHaveBeenCalledTimes(1)
    const mask = dialog.querySelector<HTMLElement>('.pdf-modal-mask')!
    await act(async () => {
      mask.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await Promise.resolve()
    })
    expect(onCancel).toHaveBeenCalledTimes(2)
  })

  it('focuses the first field on mount and returns focus on unmount', async () => {
    const opener = document.createElement('button')
    opener.textContent = 'opener'
    document.body.appendChild(opener)
    opener.focus()
    expect(document.activeElement).toBe(opener)
    try {
      const dialog = await renderDialog(() => {})
      const node = dialog.querySelector('[role="dialog"]')!
      const active = document.activeElement
      expect(active).not.toBeNull()
      expect(node.contains(active)).toBe(true)
      await act(async () => root!.unmount())
      root = null
      expect(document.activeElement).toBe(opener)
    } finally {
      opener.remove()
    }
  })
})

describe('SignatureDialog a11y', () => {
  beforeAll(() => {
    ;(window as unknown as { pdfApi: unknown }).pdfApi = {
      listSavedSignatures: () => Promise.resolve([]),
      removeSavedSignature: () => Promise.resolve(),
      addSavedSignature: () => Promise.resolve(),
    }
  })

  async function renderSign(onCancel: () => void) {
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
    await act(async () => {
      root!.render(
        createElement(SignatureDialog, { t, color: [0, 0, 0], onCancel, onConfirm: () => {} }),
      )
      await Promise.resolve()
    })
    return container
  }

  it('is a labelled modal that takes focus and closes on Escape', async () => {
    const opener = document.createElement('button')
    document.body.appendChild(opener)
    opener.focus()
    const onCancel = vi.fn()
    try {
      const host = await renderSign(onCancel)
      const node = host.querySelector('[role="dialog"]')!
      expect(node.getAttribute('aria-modal')).toBe('true')
      expect(node.getAttribute('aria-label')).toBe('signTitle')
      expect(node.contains(document.activeElement)).toBe(true)
      await act(async () => {
        window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
      })
      expect(onCancel).toHaveBeenCalledTimes(1)
      await act(async () => root!.unmount())
      root = null
      expect(document.activeElement).toBe(opener)
    } finally {
      opener.remove()
    }
  })

  it('exposes the image upload affordance as a button in image mode', async () => {
    const host = await renderSign(() => {})
    const tabs = host.querySelectorAll<HTMLButtonElement>('.pdf-sign-tab')
    await act(async () => {
      tabs[2]!.click()
    })
    const box = host.querySelector('.pdf-sign-imgbox')
    expect(box?.tagName).toBe('BUTTON')
  })
})
