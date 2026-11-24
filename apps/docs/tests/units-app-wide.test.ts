import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  MEASUREMENT_UNIT_KEY,
  connectAppWideMeasurementUnit,
  measurementUnit,
  setMeasurementUnit,
  subscribeMeasurementUnit,
} from '../src/renderer/units'

/** a shell stand-in: answers get(), records set(), lets the test broadcast */
function fakeShell(initial: unknown) {
  let handler: ((state: unknown) => void) | null = null
  const set = vi.fn()
  return {
    source: {
      get: () => Promise.resolve(initial),
      set,
      onChanged: (h: (state: unknown) => void) => {
        handler = h
        return () => {
          handler = null
        }
      },
    },
    set,
    broadcast: (state: unknown) => handler?.(state),
  }
}

const settle = () => new Promise((r) => setTimeout(r, 0))

let disconnect: (() => void) | null = null
afterEach(() => {
  disconnect?.()
  disconnect = null
  localStorage.clear()
})

describe('app-wide measurement unit in Docs', () => {
  it('follows the shell, including later changes', async () => {
    const shell = fakeShell({ unit: 'cm', explicit: false })
    const seen = vi.fn()
    const off = subscribeMeasurementUnit(seen)
    disconnect = connectAppWideMeasurementUnit(shell.source)
    await settle()
    expect(measurementUnit()).toBe('cm')
    shell.broadcast({ unit: 'in', explicit: true })
    expect(measurementUnit()).toBe('in')
    expect(seen).toHaveBeenCalled()
    off()
  })

  it('writes a choice through the shell, not local storage', async () => {
    const shell = fakeShell({ unit: 'cm', explicit: false })
    disconnect = connectAppWideMeasurementUnit(shell.source)
    await settle()
    setMeasurementUnit('mm')
    expect(shell.set).toHaveBeenCalledWith('mm')
    expect(localStorage.getItem(MEASUREMENT_UNIT_KEY)).toBeNull()
  })

  it('carries a unit picked in Docs before over to the shell once', async () => {
    localStorage.setItem(MEASUREMENT_UNIT_KEY, 'pt')
    const shell = fakeShell({ unit: 'cm', explicit: false })
    disconnect = connectAppWideMeasurementUnit(shell.source)
    await settle()
    expect(shell.set).toHaveBeenCalledWith('pt')
    expect(localStorage.getItem(MEASUREMENT_UNIT_KEY)).toBeNull()
  })

  it('does not override a unit already picked in Settings', async () => {
    localStorage.setItem(MEASUREMENT_UNIT_KEY, 'pt')
    const shell = fakeShell({ unit: 'mm', explicit: true })
    disconnect = connectAppWideMeasurementUnit(shell.source)
    await settle()
    expect(shell.set).not.toHaveBeenCalled()
    expect(measurementUnit()).toBe('mm')
  })

  it('ignores an unreadable answer', async () => {
    const shell = fakeShell({ unit: 'furlong', explicit: true })
    disconnect = connectAppWideMeasurementUnit(shell.source)
    await settle()
    expect(['in', 'cm', 'mm', 'pt', 'pi']).toContain(measurementUnit())
  })
})
