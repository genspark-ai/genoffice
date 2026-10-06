/**
 * @vitest-environment jsdom
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { useRedactLabelField } from '../src/use-redact-label'

/**
 * Five editors shared this behaviour by accident, which means it is where a
 * bug would hide: five places to forget that an empty label must not submit,
 * or that the preview has to be the sanitised text rather than what was typed.
 *
 * The second one is the one that matters. What the reader types is not what
 * the model sees — the carrier cannot hold a `*` or a `/` in a comment — so a
 * preview of the raw text would promise something the save does not do.
 */

const actEnvironment = globalThis as typeof globalThis & {
  IS_REACT_ACT_ENVIRONMENT?: boolean
}
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root
afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

interface Harness {
  latest: ReturnType<typeof useRedactLabelField> | null
  onSubmit: ReturnType<typeof vi.fn>
  onCancel: ReturnType<typeof vi.fn>
}

function render(seed = ''): Harness {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
  const h: Harness = { latest: null, onSubmit: vi.fn(), onCancel: vi.fn() }
  function Probe() {
    h.latest = useRedactLabelField(seed, h.onSubmit, h.onCancel)
    return null
  }
  act(() => root.render(createElement(Probe)))
  return h
}

function type(h: Harness, text: string): void {
  act(() => h.latest!.setValue(text))
}

function press(h: Harness, key: string): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  act(() => {
    h.latest!.onKeyDown(event as never)
  })
  return event
}

describe('the label field of a withhold dialog', () => {
  it('offers the seed as a starting label', () => {
    const h = render('客户电话')
    expect(h.latest!.value).toBe('客户电话')
    expect(h.latest!.label).toBe('客户电话')
  })

  it('previews the marker the model will see, as the reader types', () => {
    const h = render()
    type(h, 'client phone')
    expect(h.latest!.marker).toBe('{{client phone}}')
  })

  it('previews the sanitised text, not what was typed', () => {
    // the carrier cannot hold these, so a preview of the raw text would promise
    // something the save does not do
    const h = render()
    type(h, 'a*b/c')
    expect(h.latest!.label).toBe('a*b/c')
    // the shared sanitiser keeps them; a carrier that cannot says so at its own
    // call site rather than by showing a lie here
    type(h, 'a{b}c')
    expect(h.latest!.label).toBe('abc')
    expect(h.latest!.marker).toBe('{{abc}}')
  })

  it('never submits an empty label', () => {
    const h = render()
    expect(h.latest!.empty).toBe(true)
    press(h, 'Enter')
    expect(h.onSubmit).not.toHaveBeenCalled()
    type(h, '   ')
    expect(h.latest!.empty).toBe(true)
    press(h, 'Enter')
    expect(h.onSubmit).not.toHaveBeenCalled()
  })

  it('submits the sanitised label on Enter, and stops the key reaching the form', () => {
    const h = render()
    type(h, 'a{b}c')
    const event = press(h, 'Enter')
    expect(h.onSubmit).toHaveBeenCalledWith('abc')
    expect(event.defaultPrevented).toBe(true)
  })

  it('cancels on Escape', () => {
    const h = render()
    type(h, 'client')
    press(h, 'Escape')
    expect(h.onCancel).toHaveBeenCalled()
    expect(h.onSubmit).not.toHaveBeenCalled()
  })

  it('caps the input at the shared maximum', () => {
    const h = render()
    expect(h.latest!.maxLength).toBe(40)
  })
})
