// Slides follows the app-wide measurement unit the shell hands out (Settings ▸ General).
import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  connectMeasurementUnit,
  measurementUnit,
  resetMeasurementUnitForTests,
} from '../src/renderer/measurement-unit'

function fakeShell(initial: unknown) {
  let handler: ((state: unknown) => void) | null = null
  return {
    source: {
      get: () => Promise.resolve(initial),
      onChanged: (h: (state: unknown) => void) => {
        handler = h
        return () => {
          handler = null
        }
      },
    },
    broadcast: (state: unknown) => handler?.(state),
  }
}

const settle = () => new Promise((r) => setTimeout(r, 0))

afterEach(() => resetMeasurementUnitForTests())

describe('app-wide measurement unit in Slides', () => {
  it('takes the shell unit and later changes', async () => {
    const shell = fakeShell({ unit: 'mm', explicit: true })
    const off = connectMeasurementUnit(shell.source)
    await settle()
    expect(measurementUnit()).toBe('mm')
    shell.broadcast({ unit: 'pt', explicit: true })
    expect(measurementUnit()).toBe('pt')
    off()
  })

  it('keeps the locale default when the shell answer is unreadable', async () => {
    vi.spyOn(navigator, 'language', 'get').mockReturnValue('id-ID')
    const off = connectMeasurementUnit(fakeShell({ unit: 'px' }).source)
    await settle()
    expect(measurementUnit()).toBe('cm')
    off()
    vi.restoreAllMocks()
  })
})
