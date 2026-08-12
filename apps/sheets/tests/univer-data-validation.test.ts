import { DataValidationRenderMode } from '@univerjs/core'
import { describe, expect, it, vi } from 'vitest'

import { installPopulatedDataValidationArrow } from '../src/renderer/data-validation-arrow'
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
})

describe('installPopulatedDataValidationArrow', () => {
  it('overlays an arrow only for populated cells without repainting their value', () => {
    const drawWith = vi.fn((..._args: unknown[]) => undefined)
    const isHit = vi.fn((..._args: unknown[]) => true)
    const rule = { renderMode: DataValidationRenderMode.TEXT }
    const validator = {
      canvasRender: {
        drawWith,
        isHit,
        _dataValidationModel: { getRuleByLocation: () => rule },
      },
    }
    const runtime = {
      univer: {
        __getInjector: () => ({
          get: () => ({ getValidatorItem: () => validator }),
        }),
      },
    }
    const location = { unitId: 'book-1', subUnitId: 'sheet-1', row: 17, col: 4 }
    const populated = { ...location, data: { v: 'Return' } }
    const empty = { ...location, data: { v: null } }

    const disposable = installPopulatedDataValidationArrow(runtime as never)
    validator.canvasRender.drawWith({}, populated)
    validator.canvasRender.drawWith({}, empty)
    expect(drawWith).toHaveBeenCalledTimes(1)
    expect(drawWith.mock.calls[0]?.[1]).toMatchObject({ data: { v: '' } })
    expect(validator.canvasRender.isHit({}, populated)).toBe(true)
    expect(validator.canvasRender.isHit({}, empty)).toBe(false)
    expect(rule.renderMode).toBe(DataValidationRenderMode.TEXT)

    disposable.dispose()
  })
})
