/**
 * COUNTIF/SUMIFS-family criteria compares reuse the inverted-index cache
 * once a column has been compared before. The cached '=' fast path
 * materializes only exact-key matches, and the number-sensitive repair pass
 * (filterSameValueObjectResult) maps only materialized slots — so a criteria
 * Univer coerced to a number ("2026-06" → date serial 46174) matches nothing
 * on a text column from the second evaluation on: the first SUMIFS in a
 * budget sheet is right and every later one is 0. Force the dense compare
 * path whenever a non-string '=' criteria meets a column whose cache holds
 * string keys; same-type compares keep the cache fast path.
 */
import {
  ArrayValueObject,
  CELL_INVERTED_INDEX_CACHE,
  ERROR_TYPE_SET,
} from '@univerjs/engine-formula'

type BatchOperatorValue = (
  this: SheetArrayLike,
  valueObject: CriteriaLike,
  column: number,
  result: unknown[],
  batchOperatorType: unknown,
  operator?: string,
  isCaseSensitive?: boolean,
) => void

interface CriteriaLike {
  isString?(): boolean
  isError?(): boolean
}

interface SheetArrayLike {
  getUnitId(): string
  getSheetId(): string
  getCurrentColumn(): number
}

interface InvertedIndexCacheLike {
  canUseCache(
    unitId: string,
    sheetId: string,
    column: number,
    startRow: number,
    endRow: number,
  ): unknown
  getCellValuePositions(
    unitId: string,
    sheetId: string,
    column: number,
  ): ReadonlyMap<unknown, unknown> | null | undefined
}

/// Keys are only ever added to a column's cache map (clear() replaces the
/// map object), so a verdict memoized by map identity stays valid until the
/// key count changes — a filled-down COUNTIF over an all-number column must
/// not rescan every distinct key per cell.
const stringKeyVerdicts = new WeakMap<object, { size: number; hasStringKey: boolean }>()

function cacheHoldsStringKeys(
  cache: InvertedIndexCacheLike,
  array: SheetArrayLike,
  column: number,
): boolean {
  const keyed = cache.getCellValuePositions(
    array.getUnitId(),
    array.getSheetId(),
    column + array.getCurrentColumn(),
  )
  if (!keyed) return false
  const memo = stringKeyVerdicts.get(keyed)
  if (memo && memo.size === keyed.size) return memo.hasStringKey
  const errorLiterals = ERROR_TYPE_SET as ReadonlySet<string>
  let hasStringKey = false
  for (const key of keyed.keys()) {
    // Error literals are stored as strings too but never coerce to numbers.
    if (typeof key === 'string' && !errorLiterals.has(key)) {
      hasStringKey = true
      break
    }
  }
  stringKeyVerdicts.set(keyed, { size: keyed.size, hasStringKey })
  return hasStringKey
}

export function installCriteriaCompareCacheFix(): { dispose(): void } {
  const proto = ArrayValueObject.prototype as unknown as Record<
    '_batchOperatorValue',
    BatchOperatorValue | undefined
  >
  const original = proto._batchOperatorValue
  // Renamed upstream: degrade to stock behavior — repeat compares on mixed
  // columns would regress again, but nothing crashes.
  if (typeof original !== 'function') return { dispose() {} }
  const cache = CELL_INVERTED_INDEX_CACHE as unknown as InvertedIndexCacheLike
  const patched: BatchOperatorValue = function (
    valueObject,
    column,
    result,
    batchOperatorType,
    operator,
    isCaseSensitive,
  ) {
    if (
      operator === '=' &&
      valueObject?.isString?.() === false &&
      valueObject.isError?.() === false &&
      cacheHoldsStringKeys(cache, this, column)
    ) {
      const canUseCache = cache.canUseCache
      cache.canUseCache = () => ({ rowsInCache: [], rowsNotInCache: [] })
      try {
        return original.call(
          this,
          valueObject,
          column,
          result,
          batchOperatorType,
          operator,
          isCaseSensitive,
        )
      } finally {
        cache.canUseCache = canUseCache
      }
    }
    return original.call(
      this,
      valueObject,
      column,
      result,
      batchOperatorType,
      operator,
      isCaseSensitive,
    )
  }
  proto._batchOperatorValue = patched
  return {
    dispose() {
      if (proto._batchOperatorValue === patched) proto._batchOperatorValue = original
    },
  }
}
