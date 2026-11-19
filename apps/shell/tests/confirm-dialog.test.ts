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
let opener: HTMLButtonElement
let behind: HTMLButtonElement

beforeEach(() => {
  // a focused opener plus another control "underneath", standing in for Settings
  opener = document.createElement('button')
  opener.textContent = 'Sign out (opener)'
  behind = document.createElement('button')
  behind.textContent = 'Underneath'
  document.body.append(opener, behind)
  opener.focus()
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  opener.remove()
  behind.remove()
})

function mount(props: { danger?: boolean; onConfirm?: () => void; onCancel?: () => void } = {}) {
  const onConfirm = props.onConfirm ?? vi.fn()
  const onCancel = props.onCancel ?? vi.fn()
  act(() => {
    root.render(
      createElement(ConfirmDialog, {
        title: 'Sign out?',
        body: 'Really?',
        confirmLabel: 'Sign out',
        cancelLabel: 'Cancel',
        danger: props.danger,
        onConfirm,
        onCancel,
      }),
    )
  })
  return { onConfirm, onCancel }
}

const dialog = () => document.body.querySelector<HTMLElement>('[role="alertdialog"]')!
const button = (text: string) =>
  [...dialog().querySelectorAll('button')].find((b) => b.textContent === text)!

function key(k: string, shiftKey = false): void {
  act(() => {
    ;(document.activeElement ?? document.body).dispatchEvent(
      new KeyboardEvent('keydown', { key: k, shiftKey, bubbles: true, cancelable: true }),
    )
  })
}

describe('ConfirmDialog', () => {
  it('names itself from its title and body', () => {
    mount()
    const d = dialog()
    expect(d.getAttribute('aria-modal')).toBe('true')
    expect(document.getElementById(d.getAttribute('aria-labelledby')!)?.textContent).toBe(
      'Sign out?',
    )
    expect(document.getElementById(d.getAttribute('aria-describedby')!)?.textContent).toBe(
      'Really?',
    )
  })

  it('opens a destructive prompt on Cancel, and a plain one on its confirm button', () => {
    mount({ danger: true })
    expect(document.activeElement).toBe(button('Cancel'))
    expect(button('Sign out').classList.contains('btn-danger')).toBe(true)
    act(() => root.render(null))
    opener.focus()
    mount()
    expect(document.activeElement).toBe(button('Sign out'))
    expect(button('Sign out').classList.contains('btn-primary')).toBe(true)
  })

  it('keeps Tab and Shift+Tab inside the dialog', () => {
    mount({ danger: true })
    const cancel = button('Cancel')
    const confirm = button('Sign out')
    // jsdom does not move focus on Tab, so the trap is checked at the wrap points
    confirm.focus()
    key('Tab')
    expect(document.activeElement).toBe(cancel)
    key('Tab', true)
    expect(document.activeElement).toBe(confirm)
    // focus that escaped (e.g. a click underneath) is pulled back in
    behind.focus()
    key('Tab')
    expect(dialog().contains(document.activeElement)).toBe(true)
  })

  it('returns focus to the opener when it closes', () => {
    mount({ danger: true })
    expect(document.activeElement).not.toBe(opener)
    act(() => root.render(null))
    expect(document.activeElement).toBe(opener)
  })

  it('runs the action only on confirm', () => {
    const { onConfirm, onCancel } = mount({ danger: true })
    act(() => button('Cancel').click())
    expect(onCancel).toHaveBeenCalledTimes(1)
    expect(onConfirm).not.toHaveBeenCalled()
    act(() => button('Sign out').click())
    expect(onConfirm).toHaveBeenCalledTimes(1)
  })

  it('cancels on a click outside the card, not inside it', () => {
    const { onCancel } = mount()
    act(() => dialog().click())
    expect(onCancel).not.toHaveBeenCalled()
    act(() => document.body.querySelector<HTMLElement>('.modal-overlay')!.click())
    expect(onCancel).toHaveBeenCalledTimes(1)
  })

  it('takes Escape for itself, so the dialog underneath stays open', () => {
    const underneath = vi.fn()
    window.addEventListener('keydown', underneath)
    try {
      const { onCancel } = mount()
      key('Escape')
      expect(onCancel).toHaveBeenCalledTimes(1)
      expect(underneath).not.toHaveBeenCalled()
    } finally {
      window.removeEventListener('keydown', underneath)
    }
  })
})
