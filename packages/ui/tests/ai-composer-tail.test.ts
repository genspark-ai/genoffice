/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { AiComposer } from '../src/AiComposer'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const baseProps = {
  value: 'a draft',
  busy: false,
  placeholder: 'Ask anything',
  hintIdle: 'Enter sends',
  hintBusy: 'Esc stops',
  sendLabel: 'Send',
  stopLabel: 'Stop',
  onChange: () => {},
  onSend: () => {},
  onStop: () => {},
}

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

describe('AiComposer footerEnd', () => {
  it('renders send directly in the footer when the slot is unused', () => {
    act(() => root.render(createElement(AiComposer, baseProps)))
    expect(host.querySelector('.ai-input-tail')).toBeNull()
    expect(host.querySelector('.ai-input-footer > .ai-send-btn')).not.toBeNull()
  })

  it('pins the slot and send together at the right of the footer', () => {
    act(() =>
      root.render(
        createElement(AiComposer, {
          ...baseProps,
          footerEnd: createElement('button', { id: 'mic' }, 'mic'),
        }),
      ),
    )
    const tail = host.querySelector('.ai-input-tail')
    expect(tail).not.toBeNull()
    expect(tail?.querySelector('#mic')).not.toBeNull()
    expect(tail?.querySelector('.ai-send-btn')).not.toBeNull()
    expect(tail?.lastElementChild?.classList.contains('ai-send-btn')).toBe(true)
  })

  it('keeps the slot while a reply is running, so stop stays reachable', () => {
    act(() =>
      root.render(
        createElement(AiComposer, {
          ...baseProps,
          busy: true,
          footerEnd: createElement('button', { id: 'mic' }, 'mic'),
        }),
      ),
    )
    const tail = host.querySelector('.ai-input-tail')
    expect(tail?.querySelector('#mic')).not.toBeNull()
    expect(tail?.querySelector('.ai-stop-btn')).not.toBeNull()
  })
})
