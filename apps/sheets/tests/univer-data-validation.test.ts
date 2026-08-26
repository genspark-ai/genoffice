import { DataValidationRenderMode } from '@univerjs/core'
import { describe, expect, it, vi } from 'vitest'

import {
  dataValidationInputMessageHeight,
  getDataValidationInputMessage,
  installActiveCellDataValidationChrome,
  openActiveDataValidationDropdown,
  shouldShowDataValidationDropdown,
} from '../src/renderer/data-validation-dropdown'
import { toUniverDvRule } from '../src/renderer/univer-sync'

const listRule = {
  ranges: [{ startRow: 1, startColumn: 4, endRow: 100, endColumn: 4 }],
  ruleType: 'list',
  formulas: ['"Return,Exchange"'],
  allowBlank: true,
  suppressDropdown: false,
  showInputMessage: true,
  showErrorMessage: true,
}

describe('toUniverDvRule', () => {
  it('keeps plain text rendering without changing dropdown semantics', () => {
    expect(toUniverDvRule(listRule, 'file-dv-sheet-1-0')).toMatchObject({
      type: 'list',
      formula1: 'Return,Exchange',
      showDropDown: true,
      renderMode: DataValidationRenderMode.TEXT,
    })
  })

  it('preserves an OOXML request to suppress the dropdown', () => {
    expect(
      toUniverDvRule({ ...listRule, suppressDropdown: true }, 'file-dv-sheet-1-0'),
    ).toMatchObject({
      type: 'list',
      showDropDown: false,
      renderMode: DataValidationRenderMode.TEXT,
    })
  })

  it('drops a list rule formula2 — Univer reads it as per-item chip colors', () => {
    // LibreOffice writes junk formula2 ("0") on list validations; passing it
    // through painted validated cells black.
    const rule = toUniverDvRule(
      { ...listRule, formulas: ['Hitab!$K$1:$K$4', '0'] },
      'file-dv-sheet-1-0',
    )
    expect(rule).not.toHaveProperty('formula2')
    expect(rule).toMatchObject({ formula1: '=Hitab!$K$1:$K$4' })
  })

  it('keeps formula2 on non-list rules', () => {
    expect(
      toUniverDvRule(
        { ...listRule, ruleType: 'whole', formulas: ['1', '10'], operator: 'between' },
        'file-dv-sheet-1-0',
      ),
    ).toMatchObject({ formula1: '1', formula2: '10' })
  })
})

function createDropdownHarness() {
  const active = { unitId: 'book-1', sheetId: 'sheet-1', row: 17, col: 4 }
  const validation = {
    present: true,
    criteria: 'list',
    showDropDown: true,
    showInputMessage: false,
    promptTitle: 'Type',
    prompt: 'Return or Exchange',
    merged: false,
  }
  const handlers = new Map<string, (event?: { id: string }) => void>()
  const eventDisposes = new Map<string, ReturnType<typeof vi.fn>>()
  const floatingDisposes: Array<ReturnType<typeof vi.fn>> = []
  const range = {
    getRow: () => active.row,
    getLastRow: () => active.row + (validation.merged ? 1 : 0),
    getDataValidation: () =>
      validation.present
        ? {
            rule: {
              showDropDown: validation.showDropDown,
              showInputMessage: validation.showInputMessage,
              promptTitle: validation.promptTitle,
              prompt: validation.prompt,
            },
            getCriteriaType: () => validation.criteria,
          }
        : null,
  }
  const addFloatDomToRange = vi.fn(
    (_range: unknown, _layer: unknown, _layout: unknown, _id: string) => {
      const dispose = vi.fn()
      floatingDisposes.push(dispose)
      return { dispose }
    },
  )
  const worksheet = {
    getSheetId: () => active.sheetId,
    getRange: () => range,
    getCellMergeData: () => (validation.merged ? range : null),
    getRowHeight: () => 20,
    addFloatDomToRange,
  }
  const executeCommand = vi.fn()
  const componentDisposes: Array<ReturnType<typeof vi.fn>> = []
  const runtime = {
    univerAPI: {
      Event: {
        SelectionChanged: 'selection-changed',
        ActiveSheetChanged: 'active-sheet-changed',
        CommandExecuted: 'command-executed',
      },
      getActiveWorkbook: () => ({
        getId: () => active.unitId,
        getActiveSheet: () => worksheet,
        getActiveCell: () => ({
          getRow: () => active.row,
          getColumn: () => active.col,
        }),
      }),
      addEvent: vi.fn((event: string, handler: (event?: { id: string }) => void) => {
        const dispose = vi.fn()
        handlers.set(event, handler)
        eventDisposes.set(event, dispose)
        return { dispose }
      }),
      registerComponent: vi.fn(() => {
        const dispose = vi.fn()
        componentDisposes.push(dispose)
        return { dispose }
      }),
      executeCommand,
    },
  }
  return {
    active,
    validation,
    handlers,
    eventDisposes,
    floatingDisposes,
    range,
    addFloatDomToRange,
    executeCommand,
    componentDisposes,
    runtime,
  }
}

describe('active-cell data-validation chrome', () => {
  it('shows only list rules whose dropdown is not suppressed', () => {
    expect(
      shouldShowDataValidationDropdown({
        rule: { showDropDown: true },
        getCriteriaType: () => 'list',
      }),
    ).toBe(true)
    expect(
      shouldShowDataValidationDropdown({
        rule: { showDropDown: false },
        getCriteriaType: () => 'list',
      }),
    ).toBe(false)
    expect(
      shouldShowDataValidationDropdown({
        rule: { showDropDown: true },
        getCriteriaType: () => 'whole',
      }),
    ).toBe(false)
    expect(shouldShowDataValidationDropdown(null)).toBe(false)
  })

  it('shows and sizes only enabled non-empty input messages', () => {
    const message = getDataValidationInputMessage({
      rule: {
        showInputMessage: true,
        promptTitle: 'Type',
        prompt: 'Return or Exchange',
      },
      getCriteriaType: () => 'list',
    })
    expect(message).toEqual({ title: 'Type', prompt: 'Return or Exchange' })
    expect(dataValidationInputMessageHeight(message!)).toBe(40)
    expect(
      dataValidationInputMessageHeight({
        title: 'Type',
        prompt: 'A long validation message that wraps across several display lines',
      }),
    ).toBeGreaterThan(40)
    expect(
      getDataValidationInputMessage({
        rule: { showInputMessage: false, prompt: 'Hidden' },
        getCriteriaType: () => 'list',
      }),
    ).toBeNull()
    expect(
      getDataValidationInputMessage({
        rule: { showInputMessage: true },
        getCriteriaType: () => 'list',
      }),
    ).toBeNull()
  })

  it('anchors one button outside the active validation cell', () => {
    const harness = createDropdownHarness()
    const disposable = installActiveCellDataValidationChrome(harness.runtime as never)

    expect(harness.addFloatDomToRange).toHaveBeenCalledTimes(1)
    expect(harness.addFloatDomToRange.mock.calls[0]?.[2]).toEqual({
      width: 14,
      height: 14,
      marginX: '100%',
      marginY: 0,
      verticalOffsetAlign: 'bottom',
    })
    harness.handlers.get('selection-changed')?.()
    expect(harness.addFloatDomToRange).toHaveBeenCalledTimes(1)

    disposable.dispose()
  })

  it('anchors the input message from the cell midpoint below the cell', () => {
    const harness = createDropdownHarness()
    harness.validation.showInputMessage = true
    const disposable = installActiveCellDataValidationChrome(harness.runtime as never)

    expect(harness.addFloatDomToRange).toHaveBeenCalledTimes(2)
    expect(harness.addFloatDomToRange.mock.calls[1]?.[2]).toEqual({
      width: 108,
      height: 40,
      marginX: '50%',
      marginY: '100%',
    })

    disposable.dispose()
  })

  it('shows an input message for non-list validation without a dropdown', () => {
    const harness = createDropdownHarness()
    harness.validation.criteria = 'whole'
    harness.validation.showInputMessage = true
    const disposable = installActiveCellDataValidationChrome(harness.runtime as never)

    expect(harness.addFloatDomToRange).toHaveBeenCalledTimes(1)
    expect(harness.addFloatDomToRange.mock.calls[0]?.[2]).toMatchObject({
      marginX: '50%',
      marginY: '100%',
    })

    disposable.dispose()
  })

  it('anchors the button to the edge of a merged validation cell', () => {
    const harness = createDropdownHarness()
    harness.validation.merged = true
    const disposable = installActiveCellDataValidationChrome(harness.runtime as never)

    expect(harness.addFloatDomToRange.mock.calls[0]?.[2]).toEqual({
      width: 14,
      height: 14,
      marginX: '100%',
      marginY: 0,
      verticalOffsetAlign: 'bottom',
    })

    disposable.dispose()
  })

  it('moves with selection and disappears outside the validation range', () => {
    const harness = createDropdownHarness()
    const disposable = installActiveCellDataValidationChrome(harness.runtime as never)

    harness.active.row = 18
    harness.handlers.get('selection-changed')?.()
    expect(harness.floatingDisposes[0]).toHaveBeenCalledTimes(1)
    expect(harness.addFloatDomToRange).toHaveBeenCalledTimes(2)

    harness.validation.present = false
    harness.handlers.get('selection-changed')?.()
    expect(harness.floatingDisposes[1]).toHaveBeenCalledTimes(1)
    expect(harness.addFloatDomToRange).toHaveBeenCalledTimes(2)

    disposable.dispose()
  })

  it('reacts when a validation rule arrives after workbook streaming', () => {
    const harness = createDropdownHarness()
    harness.validation.present = false
    const disposable = installActiveCellDataValidationChrome(harness.runtime as never)
    expect(harness.addFloatDomToRange).not.toHaveBeenCalled()

    harness.validation.present = true
    harness.handlers.get('command-executed')?.({ id: 'other.command' })
    expect(harness.addFloatDomToRange).not.toHaveBeenCalled()
    harness.handlers.get('command-executed')?.({ id: 'data-validation.mutation.addRule' })
    expect(harness.addFloatDomToRange).toHaveBeenCalledTimes(1)

    disposable.dispose()
  })

  it('opens the dropdown for the current active cell', () => {
    const harness = createDropdownHarness()
    const preventDefault = vi.fn()
    const stopPropagation = vi.fn()

    openActiveDataValidationDropdown(
      harness.runtime as never,
      {
        preventDefault,
        stopPropagation,
      } as never,
    )

    expect(preventDefault).toHaveBeenCalledTimes(1)
    expect(stopPropagation).toHaveBeenCalledTimes(1)
    expect(harness.executeCommand).toHaveBeenCalledWith(
      'sheet.operation.show-data-validation-dropdown',
      {
        unitId: 'book-1',
        subUnitId: 'sheet-1',
        row: 17,
        column: 4,
      },
    )

    harness.validation.showDropDown = false
    openActiveDataValidationDropdown(harness.runtime as never)
    expect(harness.executeCommand).toHaveBeenCalledTimes(1)
  })

  it('disposes the floating control, events, and component', () => {
    const harness = createDropdownHarness()
    const disposable = installActiveCellDataValidationChrome(harness.runtime as never)
    disposable.dispose()

    expect(harness.floatingDisposes[0]).toHaveBeenCalledTimes(1)
    expect(harness.componentDisposes).toHaveLength(2)
    expect(harness.componentDisposes[0]).toHaveBeenCalledTimes(1)
    expect(harness.componentDisposes[1]).toHaveBeenCalledTimes(1)
    expect(harness.eventDisposes.get('selection-changed')).toHaveBeenCalledTimes(1)
    expect(harness.eventDisposes.get('active-sheet-changed')).toHaveBeenCalledTimes(1)
    expect(harness.eventDisposes.get('command-executed')).toHaveBeenCalledTimes(1)
  })
})
