import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import { trapTab, useModalKeys } from '../src/renderer/components/modal-keys'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

function keydown(
  key: string,
  shiftKey = false,
): { key: string; shiftKey: boolean; prevented: boolean; preventDefault: () => void } {
  const e = {
    key,
    shiftKey,
    prevented: false,
    preventDefault() {
      e.prevented = true
    },
  }
  return e
}

function modal(): { root: HTMLDivElement; buttons: HTMLButtonElement[]; cleanup: () => void } {
  const root = document.createElement('div')
  const buttons = [document.createElement('button'), document.createElement('button')]
  for (const b of buttons) root.appendChild(b)
  document.body.appendChild(root)
  return { root, buttons, cleanup: () => root.remove() }
}

describe('trapTab', () => {
  it('wraps Tab on the last control back to the first', () => {
    const { root, buttons, cleanup } = modal()
    try {
      buttons[1]!.focus()
      const e = keydown('Tab')
      trapTab(root, e)
      expect(e.prevented).toBe(true)
      expect(document.activeElement).toBe(buttons[0])
    } finally {
      cleanup()
    }
  })

  it('wraps Shift+Tab on the first control to the last', () => {
    const { root, buttons, cleanup } = modal()
    try {
      buttons[0]!.focus()
      const e = keydown('Tab', true)
      trapTab(root, e)
      expect(e.prevented).toBe(true)
      expect(document.activeElement).toBe(buttons[1])
    } finally {
      cleanup()
    }
  })

  it('leaves mid-list tabs and other keys alone', () => {
    const { root, buttons, cleanup } = modal()
    try {
      buttons[0]!.focus()
      const plain = keydown('Tab')
      trapTab(root, plain)
      expect(plain.prevented).toBe(false)
      const other = keydown('Enter')
      trapTab(root, other)
      expect(other.prevented).toBe(false)
    } finally {
      cleanup()
    }
  })

  it('prevents tabbing out of an empty modal', () => {
    const root = document.createElement('div')
    document.body.appendChild(root)
    try {
      const e = keydown('Tab')
      trapTab(root, e)
      expect(e.prevented).toBe(true)
    } finally {
      root.remove()
    }
  })
})

function Modal({ onClose }: { onClose: () => void }) {
  const keys = useModalKeys(onClose)
  return createElement(
    'div',
    { ref: keys.ref, onKeyDown: keys.onKeyDown },
    createElement('button', null, 'ok'),
  )
}

function mountModal(): { onClose: ReturnType<typeof vi.fn>; cleanup: () => void } {
  const onClose = vi.fn()
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(Modal, { onClose })))
  return {
    onClose,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

const windowEscape = () =>
  act(() => {
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
  })

describe('useModalKeys Escape fallback', () => {
  it('closes on Escape when focus has left the modal', () => {
    const { onClose, cleanup } = mountModal()
    try {
      ;(document.activeElement as HTMLElement | null)?.blur()
      expect(document.activeElement).toBe(document.body)
      windowEscape()
      expect(onClose).toHaveBeenCalledTimes(1)
    } finally {
      cleanup()
    }
  })

  it('leaves Escape to the modal handler while focus is inside', () => {
    const { onClose, cleanup } = mountModal()
    try {
      expect(document.activeElement?.tagName).toBe('BUTTON')
      windowEscape()
      expect(onClose).not.toHaveBeenCalled()
    } finally {
      cleanup()
    }
  })
})
