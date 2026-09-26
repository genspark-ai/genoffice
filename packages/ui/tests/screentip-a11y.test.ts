// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { installScreenTips } from '../src/screentip'

let uninstall: () => void
let button: HTMLButtonElement

beforeEach(() => {
  vi.useFakeTimers()
  uninstall = installScreenTips()
  button = document.createElement('button')
  button.setAttribute('data-tip', 'Bold')
  button.setAttribute('data-tip-kbd', 'Ctrl+B')
  document.body.appendChild(button)
})

afterEach(() => {
  uninstall()
  button.remove()
  vi.useRealTimers()
})

const tip = () => document.querySelector<HTMLElement>('.ui-screentip')

describe('screentip keyboard and ARIA', () => {
  it('shows on focus and links the anchor through aria-describedby', () => {
    button.focus()
    vi.advanceTimersByTime(500)
    expect(tip()?.style.visibility).toBe('visible')
    expect(tip()?.getAttribute('role')).toBe('tooltip')
    expect(button.getAttribute('aria-describedby')).toBe(tip()!.id)
    expect(tip()!.id).not.toBe('')
  })

  it('hides on Escape and drops the association', () => {
    button.focus()
    vi.advanceTimersByTime(500)
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    expect(tip()?.style.visibility).toBe('hidden')
    expect(button.hasAttribute('aria-describedby')).toBe(false)
  })

  it('hides on blur', () => {
    button.focus()
    vi.advanceTimersByTime(500)
    button.blur()
    expect(tip()?.style.visibility).toBe('hidden')
    expect(button.hasAttribute('aria-describedby')).toBe(false)
  })

  it('still shows on hover', () => {
    button.dispatchEvent(new PointerEvent('pointerover', { bubbles: true }))
    vi.advanceTimersByTime(500)
    expect(tip()?.style.visibility).toBe('visible')
    expect(button.getAttribute('aria-describedby')).toBe(tip()!.id)
  })
})
