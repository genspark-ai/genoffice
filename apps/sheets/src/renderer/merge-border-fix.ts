/**
 * Univer skips borders on merged ranges when only the merge's main cell
 * carries them (the common file shape — Excel writes the style on the
 * top-left cell). `_setStylesCacheForOneCell` passes a `mergeRange` only for
 * covered cells (`isMerged`), never for the main cell (`isMergedMainCell`),
 * and the non-merge branch bails via `intersectMergeRange` — so the main
 * cell's borders never reach the render cache. Passing the merge range for
 * the main cell routes it through `_setMergeBorderProps`, which walks the
 * range's edge cells and emits exactly the per-cell segments Excel draws
 * (a border on the main cell alone renders only that cell's edge segment).
 */
import type { IRange, IStyleData, Nullable } from '@univerjs/core'
import { SpreadsheetSkeleton } from '@univerjs/engine-render'

interface MergeInfoLike {
  isMergedMainCell?: boolean
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

interface SkeletonProtoLike {
  worksheet?: { getCellInfoInMergeData(row: number, col: number): MergeInfoLike }
  _setBorderStylesCache(
    row: number,
    col: number,
    style: Nullable<IStyleData>,
    options: { mergeRange?: IRange; cacheItem?: unknown } | undefined,
  ): void
}

let installed = false

export function installMergeBorderFix(): void {
  if (installed) return
  installed = true
  const proto = SpreadsheetSkeleton.prototype as unknown as SkeletonProtoLike
  const original = proto._setBorderStylesCache
  if (typeof original !== 'function') return
  proto._setBorderStylesCache = function (
    this: SkeletonProtoLike,
    row: number,
    col: number,
    style: Nullable<IStyleData>,
    options: { mergeRange?: IRange; cacheItem?: unknown } | undefined,
  ): void {
    if (style?.bd && !options?.mergeRange) {
      const info = this.worksheet?.getCellInfoInMergeData(row, col)
      if (
        info?.isMergedMainCell &&
        (info.endRow > info.startRow || info.endColumn > info.startColumn)
      ) {
        options = {
          ...options,
          mergeRange: {
            startRow: info.startRow,
            startColumn: info.startColumn,
            endRow: info.endRow,
            endColumn: info.endColumn,
          },
        }
      }
    }
    original.call(this, row, col, style, options)
  }
}
