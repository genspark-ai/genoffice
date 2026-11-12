/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { AiComposer } from '../src/AiComposer'
import { createImeEnterGuard } from '../src/ime-enter'

describe('createImeEnterGuard', () => {
  it('lets a plain Enter through when no composition ever happened', () => {
    const guard = createImeEnterGuard(() => {})
    expect(guard.isImeConfirm({ isComposing: false, keyCode: 13 })).toBe(false)
  })

  it('blocks Enter while composing (own flag, Chromium ordering)', () => {
    const guard = createImeEnterGuard(() => {})
    guard.handleCompositionStart()
    expect(guard.isImeConfirm({ isComposing: false, keyCode: 13 })).toBe(true)
  })

  it('blocks Enter on the event signals alone (isComposing / keyCode 229)', () => {
    const guard = createImeEnterGuard(() => {})
    expect(guard.isImeConfirm({ isComposing: true, keyCode: 13 })).toBe(true)
    expect(guard.isImeConfirm({ isComposing: false, keyCode: 229 })).toBe(true)
  })

  it('blocks the WebKit-ordered confirm: compositionend first, then a normal-looking Enter', () => {
    const resets: Array<() => void> = []
    const guard = createImeEnterGuard((reset) => resets.push(reset))
    guard.handleCompositionStart()
    guard.handleCompositionEnd()
    // Safari: this keydown reports isComposing=false with an ordinary keyCode
    expect(guard.isImeConfirm({ isComposing: false, keyCode: 13 })).toBe(true)
    // the next tick's deliberate Enter submits again
    resets.forEach((reset) => reset())
    expect(guard.isImeConfirm({ isComposing: false, keyCode: 13 })).toBe(false)
  })
})

describe('AiComposer IME Enter', () => {
  let host: HTMLDivElement
  let root: ReturnType<typeof createRoot>
  beforeEach(() => {
    vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
    host = document.createElement('div')
    document.body.append(host)
    root = createRoot(host)
  })
  afterEach(() => {
    act(() => root.unmount())
    host.remove()
    vi.restoreAllMocks()
  })

  function renderComposer(onSend: () => void): HTMLTextAreaElement {
    act(() =>
      root.render(
        createElement(AiComposer, {
          value: 'かな',
          busy: false,
          placeholder: '',
          hintIdle: '',
          hintBusy: '',
          sendLabel: 'send',
          stopLabel: 'stop',
          onChange: () => {},
          onSend,
          onStop: () => {},
        }),
      ),
    )
    return host.querySelector('textarea')!
  }

  /** returns whether the textarea's default (inserting a newline) was prevented */
  const pressEnter = (ta: HTMLTextAreaElement, isComposing: boolean): boolean => {
    const ev = new KeyboardEvent('keydown', {
      key: 'Enter',
      bubbles: true,
      cancelable: true,
      isComposing,
    })
    act(() => {
      ta.dispatchEvent(ev)
    })
    return ev.defaultPrevented
  }

  it('does not send on the Enter that confirms a candidate right after compositionend', async () => {
    const onSend = vi.fn()
    const ta = renderComposer(onSend)
    act(() => {
      ta.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
      ta.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'かな' }))
    })
    // swallowed outright: no send, and no newline inserted into the draft either
    expect(pressEnter(ta, false)).toBe(true)
    expect(onSend).not.toHaveBeenCalled()

    // the flag clears on the next tick, so a deliberate Enter still sends
    await new Promise((resolve) => setTimeout(resolve, 0))
    pressEnter(ta, false)
    expect(onSend).toHaveBeenCalledTimes(1)
  })

  it('does not send while the keydown itself reports composing', () => {
    const onSend = vi.fn()
    const ta = renderComposer(onSend)
    pressEnter(ta, true)
    expect(onSend).not.toHaveBeenCalled()
  })

  it('sends on a plain Enter', () => {
    const onSend = vi.fn()
    const ta = renderComposer(onSend)
    pressEnter(ta, false)
    expect(onSend).toHaveBeenCalledTimes(1)
  })
})
