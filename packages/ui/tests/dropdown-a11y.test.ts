/** @vitest-environment jsdom */
import { act, createElement, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Dropdown, type DropdownOption } from '../src/dropdown'

const OPTIONS: DropdownOption[] = [
  { value: 'a', label: 'Alpha' },
  { value: 'b', label: 'Beta' },
  { value: 'c', label: 'Gamma' },
]

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>
let onPick: ReturnType<typeof vi.fn>

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  onPick = vi.fn()
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

function Harness({
  options = OPTIONS,
  initial = 'a',
}: {
  options?: DropdownOption[]
  initial?: string
}) {
  const [value, setValue] = useState(initial)
  return createElement(Dropdown, {
    value,
    options,
    onPick: (next: string) => {
      setValue(next)
      onPick(next)
    },
  })
}

const mount = (props: { options?: DropdownOption[]; initial?: string } = {}) => {
  act(() => root.render(createElement(Harness, props)))
  return host
}

const trigger = (scope: HTMLElement): HTMLButtonElement => scope.querySelector('.gs-dd-btn')!

const press = (scope: HTMLElement, key: string, init: KeyboardEventInit = {}) =>
  act(() => {
    trigger(scope).dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, ...init }))
  })

const listbox = (scope: HTMLElement): HTMLElement | null => scope.querySelector('[role="listbox"]')

const activeOption = (scope: HTMLElement): HTMLElement | null => {
  const id = trigger(scope).getAttribute('aria-activedescendant')
  return id ? document.getElementById(id) : null
}

describe('Dropdown active option exposure', () => {
  it('has no activedescendant or controls reference while closed', () => {
    mount()
    expect(trigger(host).getAttribute('aria-activedescendant')).toBeNull()
    expect(trigger(host).getAttribute('aria-controls')).toBeNull()
    expect(trigger(host).getAttribute('aria-expanded')).toBe('false')
  })

  it('points the trigger at the listbox it controls once open', () => {
    mount()
    press(host, 'ArrowDown')
    const controls = trigger(host).getAttribute('aria-controls')
    expect(controls).toBeTruthy()
    expect(listbox(host)!.id).toBe(controls)
    expect(trigger(host).getAttribute('aria-expanded')).toBe('true')
  })

  it('references the active option and gives every option an id', () => {
    mount()
    press(host, 'ArrowDown')
    const options = [...host.querySelectorAll<HTMLElement>('[role="option"]')]
    expect(options).toHaveLength(3)
    for (const option of options) expect(option.id).toBeTruthy()
    expect(new Set(options.map((o) => o.id)).size).toBe(3)
    expect(activeOption(host)!.id).toBe(options[0]!.id)
  })

  it('moves the reference as the arrows walk the list', () => {
    mount()
    press(host, 'ArrowDown')
    const ids = [...host.querySelectorAll<HTMLElement>('[role="option"]')].map((o) => o.id)
    press(host, 'ArrowDown')
    expect(activeOption(host)!.id).toBe(ids[1])
    press(host, 'ArrowDown')
    expect(activeOption(host)!.id).toBe(ids[2])
    press(host, 'ArrowUp')
    expect(activeOption(host)!.id).toBe(ids[1])
    press(host, 'End')
    expect(activeOption(host)!.id).toBe(ids[2])
    press(host, 'Home')
    expect(activeOption(host)!.id).toBe(ids[0])
  })

  it('keeps the reference on the option that also carries the active class', () => {
    mount()
    press(host, 'ArrowDown')
    press(host, 'ArrowDown')
    const active = host.querySelector('.gs-dd-item.active')!
    expect(active.getAttribute('id')).toBe(activeOption(host)!.id)
  })

  it('never points at a disabled option', () => {
    mount({
      options: [
        { value: 'a', label: 'Alpha' },
        { value: 'b', label: 'Beta', disabled: true },
        { value: 'c', label: 'Gamma' },
      ],
    })
    press(host, 'ArrowDown')
    expect(activeOption(host)!.textContent).toBe('Alpha')
    // #945 made the cursor walk past a disabled row instead of landing on it,
    // so neither the reference nor the .active class ever points at one —
    // in either direction
    press(host, 'ArrowDown')
    expect(activeOption(host)!.textContent).toBe('Gamma')
    expect(host.querySelector('.gs-dd-item.active')!.textContent).toBe('Gamma')
    expect(activeOption(host)!.hasAttribute('disabled')).toBe(false)
    press(host, 'ArrowUp')
    expect(activeOption(host)!.textContent).toBe('Alpha')
  })

  it('keeps focus on the trigger while the reference changes', () => {
    mount()
    press(host, 'ArrowDown')
    trigger(host).focus()
    press(host, 'ArrowDown')
    expect(document.activeElement).toBe(trigger(host))
    expect(activeOption(host)!.textContent).toBe('Beta')
  })

  it('drops the reference when the list closes and restores it on reopen', () => {
    mount()
    press(host, 'ArrowDown')
    press(host, 'Escape')
    expect(trigger(host).getAttribute('aria-activedescendant')).toBeNull()
    expect(trigger(host).getAttribute('aria-controls')).toBeNull()
    press(host, 'ArrowDown')
    expect(activeOption(host)!.textContent).toBe('Alpha')
  })

  it('opens on the current value', () => {
    mount({ initial: 'c' })
    press(host, 'Enter')
    expect(activeOption(host)!.textContent).toBe('Gamma')
  })

  it('gives each instance its own ids', () => {
    act(() => {
      root.render(
        createElement(
          'div',
          null,
          createElement(Harness, { key: 'one' }),
          createElement(Harness, { key: 'two' }),
        ),
      )
    })
    const triggers = [...host.querySelectorAll<HTMLButtonElement>('.gs-dd-btn')]
    expect(triggers).toHaveLength(2)
    const [one, two] = triggers
    act(() => {
      one!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }))
      two!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }))
    })
    const listIds = [...host.querySelectorAll<HTMLElement>('[role="listbox"]')].map((l) => l.id)
    expect(listIds).toHaveLength(2)
    expect(listIds[0]).not.toBe(listIds[1])
    expect(listIds).toContain(one!.getAttribute('aria-controls'))
    expect(listIds).toContain(two!.getAttribute('aria-controls'))
    for (const option of host.querySelectorAll<HTMLElement>('[role="option"]')) {
      expect(option.id).toBeTruthy()
    }
    expect(
      new Set([...host.querySelectorAll<HTMLElement>('[role="option"]')].map((o) => o.id)).size,
    ).toBe(6)
  })
})
