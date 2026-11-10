// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { AllowEditRangesDialog } from '../src/renderer/AllowEditRangesDialog'
import { GoToDialog } from '../src/renderer/GoToDialog'
import { SubtotalDialog } from '../src/renderer/SubtotalDialog'
import { trapTab } from '../src/renderer/modal-dialog'

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
  document.body.innerHTML = ''
})

async function mount(element: React.ReactElement): Promise<HTMLDivElement> {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(element)
    await Promise.resolve()
  })
  return container
}

const keydown = (target: EventTarget, init: KeyboardEventInit): boolean =>
  target.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }))

describe('Sheets modal dialogs', () => {
  it('Go To: aria-modal, Escape closes from the input and from body, Tab stays inside', async () => {
    const opener = document.createElement('button')
    document.body.appendChild(opener)
    opener.focus()
    const onClose = vi.fn()
    const host = await mount(createElement(GoToDialog, { names: [], onGo: () => null, onClose }))
    const dialog = host.querySelector<HTMLElement>('[role="dialog"]')!
    expect(dialog.getAttribute('aria-modal')).toBe('true')

    const input = dialog.querySelector('input')!
    expect(document.activeElement).toBe(input)
    await act(async () => void keydown(input, { key: 'Escape' }))
    expect(onClose).toHaveBeenCalledTimes(1)

    ;(document.activeElement as HTMLElement).blur()
    await act(async () => void keydown(window, { key: 'Escape' }))
    expect(onClose).toHaveBeenCalledTimes(2)

    const last = dialog.querySelector<HTMLButtonElement>('button.secondary')!
    last.focus()
    await act(async () => void keydown(last, { key: 'Tab' }))
    expect(document.activeElement).toBe(input)
    await act(async () => void keydown(input, { key: 'Tab', shiftKey: true }))
    expect(document.activeElement).toBe(last)

    await act(async () => root!.unmount())
    root = null
    expect(document.activeElement).toBe(opener)
  })

  it('Subtotal: focuses the first control on open and returns focus to the opener on close', async () => {
    const opener = document.createElement('button')
    document.body.appendChild(opener)
    opener.focus()
    const fields = [
      { colIndex: 0, label: 'Region' },
      { colIndex: 1, label: 'Sales' },
    ]
    const host = await mount(
      createElement(SubtotalDialog, { fields, onCreate: () => null, onClose: () => {} }),
    )
    const dialog = host.querySelector<HTMLElement>('[role="dialog"]')!
    expect(dialog.getAttribute('aria-modal')).toBe('true')
    expect(dialog.contains(document.activeElement)).toBe(true)
    await act(async () => root!.unmount())
    root = null
    expect(document.activeElement).toBe(opener)
  })

  it('keeps focus on an autoFocus child that is not the first control', async () => {
    const host = await mount(
      createElement(AllowEditRangesDialog, {
        ranges: [{ name: 'Range1', sqref: 'A1:B2' }],
        defaultRef: 'C3',
        onApply: () => null,
        onClose: () => {},
      }),
    )
    const dialog = host.querySelector<HTMLElement>('[role="dialog"]')!
    const first = dialog.querySelector<HTMLElement>('button, input')!
    const autoFocused = dialog.querySelectorAll('input')[1]!
    expect(document.activeElement).not.toBe(first)
    expect(document.activeElement).toBe(autoFocused)
  })

  it('Tab from body (after clicking non-focusable chrome) re-enters the card', async () => {
    const host = await mount(
      createElement(GoToDialog, { names: [], onGo: () => null, onClose: () => {} }),
    )
    const dialog = host.querySelector<HTMLElement>('[role="dialog"]')!
    const input = dialog.querySelector('input')!
    const cancel = dialog.querySelector<HTMLButtonElement>('button.secondary')!
    input.blur()
    expect(document.activeElement).toBe(document.body)
    await act(async () => void keydown(window, { key: 'Tab' }))
    expect(document.activeElement).toBe(input)
    input.blur()
    await act(async () => void keydown(window, { key: 'Tab', shiftKey: true }))
    expect(document.activeElement).toBe(cancel)
  })

  it('trapTab swallows Tab when the card has no focusable control', () => {
    const card = document.createElement('div')
    const event = { shiftKey: false, preventDefault: vi.fn() }
    trapTab(card, event)
    expect(event.preventDefault).toHaveBeenCalled()
  })
})
