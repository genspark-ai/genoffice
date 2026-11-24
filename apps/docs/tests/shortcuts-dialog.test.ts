/**
 * The Keyboard Shortcuts sheet: renders the registry grouped,
 * filters as you type, and closes on Escape / backdrop like the other modals.
 */
import { describe, expect, it, vi } from 'vitest'
import { createElement, act } from 'react'
import { createRoot } from 'react-dom/client'
import { ShortcutsDialog } from '../src/renderer/components/ShortcutsDialog'
import { SHORTCUTS, shortcutKeys } from '../src/renderer/shortcuts'
import { t } from '../src/renderer/i18n/locale'

function render(onClose = () => {}) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(ShortcutsDialog, { onClose })))
  return {
    container,
    rows: () => [...container.querySelectorAll('.sc-row')],
    type: (value: string) => {
      const input = container.querySelector<HTMLInputElement>('.sc-filter')!
      // React tracks the value node, so assigning .value directly is invisible to onChange
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      act(() => {
        setter.call(input, value)
        input.dispatchEvent(new Event('input', { bubbles: true }))
      })
    },
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

describe('ShortcutsDialog', () => {
  it('lists every registry entry that has a chord on this platform under a group heading', () => {
    const view = render()
    const withChord = SHORTCUTS.filter((def) => shortcutKeys(def))
    expect(withChord.length).toBeLessThan(SHORTCUTS.length)
    expect(view.rows().length).toBe(withChord.length)
    expect(view.container.querySelectorAll('.sc-group h3').length).toBeGreaterThan(3)
    view.unmount()
  })

  it('filters by command name and by chord', () => {
    const view = render()
    view.type(t('ribbonPageBreak'))
    expect(view.rows().length).toBe(1)
    view.type('F9')
    const onF9 = SHORTCUTS.filter((def) => shortcutKeys(def).includes('F9'))
    expect(onF9.length).toBeGreaterThanOrEqual(2)
    expect(view.rows().map((row) => row.querySelector('.sc-label')!.textContent)).toEqual(
      onF9.map((def) => t(def.labelKey)),
    )
    view.type('nothing here')
    expect(view.rows().length).toBe(0)
    expect(view.container.querySelector('.sc-empty')).not.toBeNull()
    view.unmount()
  })

  it('closes on Escape and on a backdrop click', () => {
    const onClose = vi.fn()
    const view = render(onClose)
    const backdrop = view.container.querySelector<HTMLElement>('.modal-backdrop')!
    act(() => {
      backdrop.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
      )
    })
    expect(onClose).toHaveBeenCalledTimes(1)
    act(() => {
      backdrop.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }))
    })
    expect(onClose).toHaveBeenCalledTimes(2)
    view.unmount()
  })
})
