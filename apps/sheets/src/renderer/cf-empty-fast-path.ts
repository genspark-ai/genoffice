/**
 * Univer's conditional-formatting view model answers "which rules cover this
 * cell" on every cell read the renderer makes — and on a sheet with no rules
 * it still builds a cache key, misses (empty results are never cached) and
 * runs an R-tree search before returning nothing. Skip straight to the empty
 * answer when the sheet has no rules at all.
 */
import { ConditionalFormattingViewModel } from '@univerjs/preset-sheets-conditional-formatting'

interface RuleModelHost {
  _conditionalFormattingRuleModel?: {
    getSubunitRules(unitId: string, subUnitId: string): readonly unknown[] | null | undefined
  }
}

let installed = false

export function installCfEmptySheetFastPath(): void {
  if (installed) return
  installed = true
  const proto = ConditionalFormattingViewModel.prototype as unknown as {
    getCellCfs(unitId: string, subUnitId: string, row: number, col: number): unknown[]
  }
  const original = proto.getCellCfs
  if (typeof original !== 'function') return
  proto.getCellCfs = function (
    this: RuleModelHost,
    unitId: string,
    subUnitId: string,
    row: number,
    col: number,
  ) {
    const rules = this._conditionalFormattingRuleModel?.getSubunitRules(unitId, subUnitId)
    if (!rules || rules.length === 0) return []
    return original.call(this, unitId, subUnitId, row, col)
  }
}
