import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { PANNING_CLASS, attachHandPan } from '../src/renderer/hand-pan'

const pointer = (type: string, x: number, y: number, button = 0): MouseEvent => {
  const e = new MouseEvent(type, { bubbles: true, clientX: x, clientY: y, button })
  Object.defineProperty(e, 'pointerId', { value: 1 })
  return e
}

describe('hand tool drag-to-pan', () => {
  let el: HTMLDivElement
  let child: HTMLSpanElement
  let detach: () => void

  beforeEach(() => {
    el = document.createElement('div')
    child = document.createElement('span')
    el.append(child)
    document.body.append(el)
    Object.defineProperty(el, 'clientWidth', { value: 500 })
    Object.defineProperty(el, 'clientHeight', { value: 400 })
    el.scrollLeft = 100
    el.scrollTop = 200
    detach = attachHandPan(el)
  })
  afterEach(() => {
    detach()
    el.remove()
  })

  it('scrolls against the drag and swallows the click that ends it', () => {
    const clicked = vi.fn()
    el.addEventListener('click', clicked)
    child.dispatchEvent(pointer('pointerdown', 50, 50))
    window.dispatchEvent(pointer('pointermove', 20, 10))
    expect(el.scrollLeft).toBe(130)
    expect(el.scrollTop).toBe(240)
    expect(el.classList.contains(PANNING_CLASS)).toBe(true)
    window.dispatchEvent(pointer('pointerup', 20, 10))
    expect(el.classList.contains(PANNING_CLASS)).toBe(false)
    child.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    expect(clicked).not.toHaveBeenCalled()
  })

  it('leaves a press without travel as a plain click', () => {
    const clicked = vi.fn()
    el.addEventListener('click', clicked)
    child.dispatchEvent(pointer('pointerdown', 50, 50))
    window.dispatchEvent(pointer('pointermove', 51, 51))
    window.dispatchEvent(pointer('pointerup', 51, 51))
    child.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    expect(el.scrollLeft).toBe(100)
    expect(clicked).toHaveBeenCalledTimes(1)
  })

  it('does not pan from a control, the scrollbar or another button', () => {
    const button = document.createElement('button')
    el.append(button)
    button.dispatchEvent(pointer('pointerdown', 50, 50))
    window.dispatchEvent(pointer('pointermove', 0, 0))
    el.dispatchEvent(pointer('pointerdown', 505, 50))
    window.dispatchEvent(pointer('pointermove', 400, 0))
    child.dispatchEvent(pointer('pointerdown', 50, 50, 2))
    window.dispatchEvent(pointer('pointermove', 0, 0))
    expect([el.scrollLeft, el.scrollTop]).toEqual([100, 200])
  })

  it('stops panning once detached', () => {
    detach()
    child.dispatchEvent(pointer('pointerdown', 50, 50))
    window.dispatchEvent(pointer('pointermove', 0, 0))
    expect(el.scrollLeft).toBe(100)
  })
})
