import { ConditionalFormattingViewModel } from '@univerjs/preset-sheets-conditional-formatting'
import { describe, expect, it } from 'vitest'

import { installCfEmptySheetFastPath } from '../src/renderer/cf-empty-fast-path'

describe('conditional formatting empty-sheet fast path', () => {
  it('answers no rules without touching the R-tree when the sheet has none (null or empty)', () => {
    installCfEmptySheetFastPath()
    const getCellCfs = (
      ConditionalFormattingViewModel.prototype as unknown as {
        getCellCfs(this: unknown, u: string, s: string, r: number, c: number): unknown[]
      }
    ).getCellCfs
    for (const rules of [null, undefined, []]) {
      const host = {
        _conditionalFormattingRuleModel: { getSubunitRules: () => rules },
        _createCacheKey: () => {
          throw new Error('should not reach the original lookup')
        },
      }
      expect(getCellCfs.call(host, 'u', 's', 3, 4)).toEqual([])
    }
  })
})
