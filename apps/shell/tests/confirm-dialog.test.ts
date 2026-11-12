/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ConfirmDialog } from '../src/renderer/src/ConfirmDialog'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

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

function mount(onConfirm: () => void, onClose: () => void, danger = false): void {
  act(() => {
    root.render(
      createElement(ConfirmDialog, {
        request: {
          title: 'Sign out?',
          body: 'Really?',
          confirmLabel: 'Sign out',
          danger,
          onConfirm,
        },
        cancelLabel: 'Cancel',
        onClose,
      }),
    )
  })
}

const button = (text: string) =>
  [...document.body.querySelectorAll<HTMLButtonElement>('.cd-card button')].find(
    (b) => b.textContent === text,
  )!

describe('ConfirmDialog', () => {
  it('runs the action only on confirm, and closes either way', () => {
    const onConfirm = vi.fn()
    const onClose = vi.fn()
    mount(onConfirm, onClose)
    expect(document.body.querySelector('[role="alertdialog"]')?.textContent).toContain('Really?')

    act(() => button('Cancel').click())
    expect(onConfirm).not.toHaveBeenCalled()
    expect(onClose).toHaveBeenCalledTimes(1)

    act(() => button('Sign out').click())
    expect(onConfirm).toHaveBeenCalledTimes(1)
    expect(onClose).toHaveBeenCalledTimes(2)
  })

  it('focuses the confirm button and marks a destructive action', () => {
    mount(vi.fn(), vi.fn(), true)
    expect(document.activeElement).toBe(button('Sign out'))
    expect(button('Sign out').classList.contains('danger')).toBe(true)
  })

  it('takes Escape for itself so the dialog underneath stays open', () => {
    const onClose = vi.fn()
    const underneath = vi.fn()
    window.addEventListener('keydown', underneath)
    mount(vi.fn(), onClose)
    act(() => {
      document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    window.removeEventListener('keydown', underneath)
    expect(onClose).toHaveBeenCalledTimes(1)
    expect(underneath).not.toHaveBeenCalled()
  })
})
