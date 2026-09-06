import { describe, expect, it } from 'vitest'
import { rangeSlot } from '../src/renderer/dom-range'

describe('rangeSlot', () => {
  it('hands out one Range per slot and keeps slots independent', () => {
    const a = rangeSlot()
    const b = rangeSlot()
    const r1 = a()
    expect(a()).toBe(r1)
    expect(b()).not.toBe(r1)
    expect(b()).toBe(b())
  })

  it('creates the Range lazily and reuses it after repositioning', () => {
    const slot = rangeSlot()
    const t = document.createTextNode('hello world')
    document.body.appendChild(t)
    const r = slot()
    r.setStart(t, 0)
    r.setEnd(t, 5)
    expect(r.toString()).toBe('hello')
    slot().setStart(t, 6)
    slot().setEnd(t, 11)
    expect(r.toString()).toBe('world')
    t.remove()
  })
})
