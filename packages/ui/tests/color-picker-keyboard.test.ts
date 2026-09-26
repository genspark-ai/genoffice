// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { ColorPicker } from '../src/color-picker'

const strings = { themeColors: 'Theme Colors', standardColors: 'Standard Colors' }
let root: Root | null = null

function mount(value: string | null, onPick = vi.fn()): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  act(() => root!.render(createElement(ColorPicker, { value, strings, onPick })))
  return container
}

afterEach(() => {
  act(() => root?.unmount())
  document.body.innerHTML = ''
})

const swatches = (c: HTMLElement) => [...c.querySelectorAll<HTMLButtonElement>('.gcp-swatch')]
const press = (el: HTMLElement, key: string) =>
  act(() => {
    el.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
  })

describe('ColorPicker keyboard grid', () => {
  it('exposes grid semantics and marks the current colour selected', () => {
    const c = mount('#FF0000')
    expect(c.querySelectorAll('[role="grid"]').length).toBeGreaterThan(0)
    const all = swatches(c)
    expect(all.every((b) => b.getAttribute('role') === 'gridcell')).toBe(true)
    const selected = all.filter((b) => b.getAttribute('aria-selected') === 'true')
    expect(selected.map((b) => b.getAttribute('aria-label'))).toEqual(['Red'])
    expect(all.filter((b) => b.tabIndex === 0)).toEqual(selected)
  })

  it('roves focus with arrow keys and keeps a single tab stop', () => {
    const c = mount(null)
    const all = swatches(c)
    expect(all.filter((b) => b.tabIndex === 0)).toEqual([all[0]])
    act(() => all[0].focus())
    press(all[0], 'ArrowRight')
    expect(document.activeElement).toBe(all[1])
    press(all[1], 'ArrowDown')
    expect(document.activeElement).toBe(all[11])
    expect(all.filter((b) => b.tabIndex === 0)).toEqual([all[11]])
    press(all[11], 'End')
    expect(document.activeElement).toBe(all[19])
    press(all[19], 'ArrowRight')
    expect(document.activeElement).toBe(all[20])
  })
})
