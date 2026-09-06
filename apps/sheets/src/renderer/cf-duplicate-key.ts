/**
 * Excel's duplicateValues / uniqueValues conditional formats compare cells by
 * displayed text: the number 1981233 and the text "1981233" are duplicates,
 * and text compares case-insensitively. Univer's HighlightCellCalculateUnit
 * keys its occurrence map on the raw cell value, so a column that stores the
 * same id as both numbers and shared strings only flags one kind.
 *
 * The unit reads cells through a per-rule `context.getCellValue`; this module
 * swaps in a reader that hands the unit a synthetic text cell carrying the
 * display key whenever the rule's current sub-type is unique/duplicate.
 */
import { BooleanNumber, CellValueType, type ICellData } from '@univerjs/core'
import {
  CFRuleType,
  CFSubRuleType,
  ConditionalFormattingViewModel,
} from '@univerjs/preset-sheets-conditional-formatting'

import type { UniverRuntime } from './univer-state'

/// Excel keeps 15 significant digits; folding to that hides double noise
/// (0.1+0.2) the way General display does.
function generalNumberText(value: number): string {
  if (!Number.isFinite(value)) return String(value)
  return String(Number(value.toPrecision(15)))
}

export function cfDisplayKey(cell: ICellData | null | undefined): string | null {
  if (!cell) return null
  if (cell.t === CellValueType.BOOLEAN) {
    return cell.v === BooleanNumber.TRUE || cell.v === true ? 'TRUE' : 'FALSE'
  }
  let value: unknown = cell.v
  if (value === null || value === undefined || value === '') {
    const stream = cell.p?.body?.dataStream?.replace(/\r\n$/, '')
    if (!stream) return null
    value = stream
  }
  if (typeof value === 'number') return generalNumberText(value)
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  const text = String(value)
  return text.trim() === '' ? null : text.toUpperCase()
}

interface CalculateUnitLike {
  _rule?: { rule?: { type?: string; subType?: string } }
  _context?: { getCellValue: (row: number, col: number) => ICellData }
}

interface ViewModelLike {
  _createRuleCalculateUnitInstance: (
    unitId: string,
    subUnitId: string,
    rule: { rule: { type: string } },
  ) => CalculateUnitLike | undefined
}

function isDisplayKeyRule(unit: CalculateUnitLike): boolean {
  const subType = unit._rule?.rule?.subType
  return subType === CFSubRuleType.duplicateValues || subType === CFSubRuleType.uniqueValues
}

export function wrapCalculateUnit(
  unit: CalculateUnitLike | undefined,
): CalculateUnitLike | undefined {
  const context = unit?._context
  if (!unit || !context || typeof context.getCellValue !== 'function') return unit
  const raw = context.getCellValue
  unit._context = {
    ...context,
    getCellValue: (row, col) => {
      const cell = raw(row, col)
      if (!isDisplayKeyRule(unit)) return cell
      const key = cfDisplayKey(cell)
      return key === null ? {} : { v: key, t: CellValueType.STRING }
    },
  }
  return unit
}

export function wrapCfViewModel(viewModel: ViewModelLike): { dispose(): void } {
  const original = viewModel._createRuleCalculateUnitInstance
  viewModel._createRuleCalculateUnitInstance = function (unitId, subUnitId, rule) {
    const unit = original.call(this, unitId, subUnitId, rule)
    return rule.rule.type === CFRuleType.highlightCell ? wrapCalculateUnit(unit) : unit
  }
  return {
    dispose() {
      viewModel._createRuleCalculateUnitInstance = original
    },
  }
}

export function installCfDisplayKeyCompare(runtime: UniverRuntime): { dispose(): void } {
  let viewModel: ViewModelLike
  try {
    viewModel = runtime.univer
      .__getInjector()
      .get(ConditionalFormattingViewModel) as unknown as ViewModelLike
  } catch {
    return { dispose() {} }
  }
  if (typeof viewModel._createRuleCalculateUnitInstance !== 'function') return { dispose() {} }
  return wrapCfViewModel(viewModel)
}
