import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  installRibbonPeekDismiss,
  isRibbonToggleShortcut,
  readRibbonCollapsed,
} from '@genoffice/ui'

const press = (target: Element) =>
  target.dispatchEvent(new Event('pointerdown', { bubbles: true, composed: true }))

describe('ribbon collapse', () => {
  let ribbon: HTMLDivElement
  let doc: HTMLDivElement
  let close: ReturnType<typeof vi.fn<() => void>>
  let off: () => void

  beforeEach(() => {
    document.body.innerHTML = ''
    ribbon = document.createElement('div')
    doc = document.createElement('div')
    document.body.append(ribbon, doc)
    close = vi.fn<() => void>()
    off = installRibbonPeekDismiss(() => ribbon, close)
  })
  afterEach(() => {
    off()
    document.documentElement.classList.remove('genoffice-popover-open')
  })

  it('a press inside the ribbon keeps the peek open', () => {
    press(ribbon)
    expect(close).not.toHaveBeenCalled()
  })

  it('a press in the document closes the peek', () => {
    press(doc)
    expect(close).toHaveBeenCalledTimes(1)
  })

  it('Escape closes the peek; teardown removes the listeners', () => {
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    expect(close).toHaveBeenCalledTimes(1)
    off()
    press(doc)
    expect(close).toHaveBeenCalledTimes(1)
  })

  it('a press inside a portaled popover (unmounted by its own dismissal) does not close', async () => {
    const popover = document.createElement('div')
    document.body.append(popover)
    document.documentElement.classList.add('genoffice-popover-open')
    press(popover)
    // the popover's own dismiss listener unmounts it and drops the html class
    popover.remove()
    document.documentElement.classList.remove('genoffice-popover-open')
    await new Promise((r) => setTimeout(r, 0))
    expect(close).not.toHaveBeenCalled()
  })

  it('a document press that only dismissed an open popover still closes the peek', async () => {
    document.documentElement.classList.add('genoffice-popover-open')
    press(doc)
    expect(close).not.toHaveBeenCalled()
    document.documentElement.classList.remove('genoffice-popover-open')
    await new Promise((r) => setTimeout(r, 0))
    expect(close).toHaveBeenCalledTimes(1)
  })

  it('a press inside a popover that stays open does not close', async () => {
    document.documentElement.classList.add('genoffice-popover-open')
    press(doc)
    await new Promise((r) => setTimeout(r, 0))
    expect(close).not.toHaveBeenCalled()
  })

  it('Escape with a ribbon popover open leaves the peek to the popover', () => {
    document.documentElement.classList.add('genoffice-popover-open')
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    expect(close).not.toHaveBeenCalled()
    document.documentElement.classList.remove('genoffice-popover-open')
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    expect(close).toHaveBeenCalledTimes(1)
  })

  it('a held shortcut (key repeat) does not re-toggle', () => {
    expect(
      isRibbonToggleShortcut(
        new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true, repeat: true }),
      ),
    ).toBe(false)
  })

  it('Ctrl+F1 is the toggle shortcut, F1 alone is not', () => {
    expect(isRibbonToggleShortcut(new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true }))).toBe(
      true,
    )
    expect(isRibbonToggleShortcut(new KeyboardEvent('keydown', { key: 'F1' }))).toBe(false)
    expect(
      isRibbonToggleShortcut(
        new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true, altKey: true }),
      ),
    ).toBe(false)
  })

  it('reads the persisted flag, defaulting to expanded', () => {
    localStorage.removeItem('t.ribbon')
    expect(readRibbonCollapsed('t.ribbon')).toBe(false)
    localStorage.setItem('t.ribbon', '1')
    expect(readRibbonCollapsed('t.ribbon')).toBe(true)
  })
})
