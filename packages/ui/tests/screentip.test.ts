/** @vitest-environment jsdom */
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { installScreenTips } from '../src/screentip'

let uninstall: () => void
let button: HTMLButtonElement

function rect(left: number, top: number, width: number, height: number): DOMRect {
  return {
    left,
    top,
    width,
    height,
    right: left + width,
    bottom: top + height,
    x: left,
    y: top,
    toJSON: () => ({}),
  } as DOMRect
}

function tipEl(): HTMLElement {
  return document.getElementById('ui-screentip') as HTMLElement
}

function hoverAt(el: Element, x: number, y: number): void {
  el.dispatchEvent(new MouseEvent('pointerover', { bubbles: true, clientX: x, clientY: y }))
  el.dispatchEvent(new MouseEvent('pointermove', { bubbles: true, clientX: x, clientY: y }))
}

beforeEach(() => {
  vi.useFakeTimers()
  Object.defineProperty(window, 'innerWidth', { configurable: true, value: 1200 })
  Object.defineProperty(window, 'innerHeight', { configurable: true, value: 800 })
  // a 24px toolbar icon near the top of the window
  button = document.createElement('button')
  button.setAttribute('data-tip', 'Bold')
  button.getBoundingClientRect = () => rect(100, 10, 24, 24)
  document.body.appendChild(button)
  uninstall = installScreenTips(document)
  // the tip measures 80x24 once created
  vi.spyOn(HTMLDivElement.prototype, 'getBoundingClientRect').mockImplementation(function (
    this: HTMLDivElement,
  ) {
    return this.id === 'ui-screentip' ? rect(0, 0, 80, 24) : rect(0, 0, 0, 0)
  })
})

afterEach(() => {
  uninstall()
  button.remove()
  vi.restoreAllMocks()
  vi.useRealTimers()
})

it('places the tip below the mouse cursor, not under it', () => {
  hoverAt(button, 112, 22)
  vi.advanceTimersByTime(500)
  expect(tipEl().style.visibility).toBe('visible')
  // a 2x-size arrow cursor hangs ~40px below its hotspot (y=22)
  expect(parseInt(tipEl().style.top, 10)).toBeGreaterThanOrEqual(22 + 40)
})

it('keeps the tip right under the control for keyboard focus', () => {
  button.dispatchEvent(new FocusEvent('focusin', { bubbles: true }))
  vi.advanceTimersByTime(500)
  expect(tipEl().style.visibility).toBe('visible')
  expect(tipEl().style.top).toBe(`${34 + 6}px`)
})

it('flips the tip above the control when below would clip', () => {
  button.getBoundingClientRect = () => rect(100, 770, 24, 24)
  hoverAt(button, 112, 782)
  vi.advanceTimersByTime(500)
  expect(tipEl().style.top).toBe(`${770 - 24 - 6}px`)
})
