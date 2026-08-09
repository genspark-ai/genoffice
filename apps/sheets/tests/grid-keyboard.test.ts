import {
  CommandType,
  EDITOR_ACTIVATED,
  FOCUSING_COMMON_DRAWINGS,
  FOCUSING_SHEET,
  FOCUSING_UNIVER_EDITOR,
} from '@univerjs/core'
import { ICommandService } from '@univerjs/core'
import { IShortcutService } from '@univerjs/ui'
import { describe, expect, it, vi } from 'vitest'

import { GridNavigationCommand, installGridKeyboardShortcuts } from '../src/renderer/grid-keyboard'

function makeHarness() {
  const commands: { id: string; handler: () => boolean; type: CommandType }[] = []
  const shortcuts: {
    id: string
    binding: number
    preconditions?: (context: { getContextValue(key: string): boolean }) => boolean
  }[] = []
  const disposals = vi.fn()
  const active = { getRow: () => 40, getColumn: () => 5 }
  const dataEnd = { getRow: () => 40, getColumn: () => 12 }
  const getRange = vi.fn((row: number, column: number) => ({
    getRow: () => row,
    getColumn: () => column,
  }))
  const scrollToCell = vi.fn()
  const setActiveRange = vi.fn()
  const selection = { getActiveRange: () => active, getNextDataRange: () => dataEnd }
  const worksheet = {
    getSelection: () => selection,
    getRange,
    getMaxRows: () => 100,
    getMaxColumns: () => 20,
    scrollToCell,
  }
  const commandService = {
    registerCommand: (command: { id: string; handler: () => boolean; type: CommandType }) => {
      commands.push(command)
      return { dispose: disposals }
    },
  }
  const shortcutService = {
    registerShortcut: (shortcut: (typeof shortcuts)[number]) => {
      shortcuts.push(shortcut)
      return { dispose: disposals }
    },
  }
  const runtime = {
    univer: {
      __getInjector: () => ({
        get: (token: unknown) => (token === ICommandService ? commandService : shortcutService),
      }),
    },
    univerAPI: {
      Enum: { Direction: { RIGHT: 1 } },
      getActiveWorkbook: () => ({ getActiveSheet: () => worksheet, setActiveRange }),
    },
  }
  return { commands, shortcuts, disposals, getRange, scrollToCell, setActiveRange, runtime }
}

describe('installGridKeyboardShortcuts', () => {
  it("registers navigation plus Delete and Backspace in Univer's dispatcher", () => {
    const harness = makeHarness()
    const disposable = installGridKeyboardShortcuts(harness.runtime as never)

    expect(harness.commands.map(({ id }) => id)).toEqual(Object.values(GridNavigationCommand))
    expect(harness.commands.every(({ type }) => type === CommandType.OPERATION)).toBe(true)
    expect(harness.shortcuts.map(({ binding }) => binding)).toEqual([36, 35, 33, 34, 46, 8])
    expect(harness.shortcuts.slice(-2).map(({ id }) => id)).toEqual([
      'sheet.command.clear-selection-content',
      'sheet.command.clear-selection-content',
    ])

    harness.commands.find(({ id }) => id === GridNavigationCommand.Home)?.handler()
    harness.commands.find(({ id }) => id === GridNavigationCommand.End)?.handler()
    harness.commands.find(({ id }) => id === GridNavigationCommand.PageUp)?.handler()
    harness.commands.find(({ id }) => id === GridNavigationCommand.PageDown)?.handler()
    expect(harness.getRange).toHaveBeenCalledWith(40, 0)
    expect(harness.scrollToCell).toHaveBeenNthCalledWith(1, 40, 0)
    expect(harness.scrollToCell).toHaveBeenNthCalledWith(2, 40, 12)
    expect(harness.scrollToCell).toHaveBeenNthCalledWith(3, 10, 5)
    expect(harness.scrollToCell).toHaveBeenNthCalledWith(4, 70, 5)

    disposable.dispose()
    expect(harness.disposals).toHaveBeenCalledTimes(10)
  })

  it("accepts Univer's hidden grid editor but rejects actual text editing", () => {
    const { shortcuts, runtime } = makeHarness()
    installGridKeyboardShortcuts(runtime as never)
    const precondition = shortcuts[0]?.preconditions
    const context = (values: Record<string, boolean>) => ({
      getContextValue: (key: string) => values[key] ?? false,
    })

    expect(
      precondition?.(context({ [FOCUSING_SHEET]: true, [FOCUSING_UNIVER_EDITOR]: true })),
    ).toBe(true)
    expect(
      precondition?.(
        context({
          [FOCUSING_SHEET]: true,
          [FOCUSING_UNIVER_EDITOR]: true,
          [EDITOR_ACTIVATED]: true,
        }),
      ),
    ).toBe(false)
    expect(
      precondition?.(
        context({
          [FOCUSING_SHEET]: true,
          [FOCUSING_UNIVER_EDITOR]: true,
          [FOCUSING_COMMON_DRAWINGS]: true,
        }),
      ),
    ).toBe(false)
  })
})
