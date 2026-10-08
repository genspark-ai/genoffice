/**
 * Shared-formula groups from the sidecar's formula index. The index lists
 * only masters and ordinary formula cells; a follower's text is the master's
 * formula shifted by the follower's offset, derived here on demand so a
 * 150k-follower column costs one entry instead of 150k strings.
 */
import { offsetFormulaRefs } from '@genoffice/xlsx-gateway/domain/formula-shift'

import type { WorkbookSharedFormulaGroup } from '../shared/desktop-api'

const ROW_BUCKET = 256

export function sharedFormulaText(
  group: WorkbookSharedFormulaGroup,
  row: number,
  column: number,
): string {
  return offsetFormulaRefs(group.formula, row - group.row, column - group.column)
}

function spanArea(range: NonNullable<WorkbookSharedFormulaGroup['range']>): number {
  return (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1)
}

export class SharedFormulaLookup {
  /** follower cells across every group (masters excluded) */
  readonly followerCount: number
  private readonly spansByRowBucket = new Map<number, WorkbookSharedFormulaGroup[]>()
  private readonly explicit = new Map<string, WorkbookSharedFormulaGroup>()

  readonly groups: readonly WorkbookSharedFormulaGroup[]

  constructor(groups: readonly WorkbookSharedFormulaGroup[] | undefined) {
    this.groups = groups ?? []
    let followers = 0
    for (const group of this.groups) {
      if (group.range) {
        followers += spanArea(group.range) - 1
        const first = Math.floor(group.range.startRow / ROW_BUCKET)
        const last = Math.floor(group.range.endRow / ROW_BUCKET)
        for (let bucket = first; bucket <= last; bucket += 1) {
          let list = this.spansByRowBucket.get(bucket)
          if (!list) {
            list = []
            this.spansByRowBucket.set(bucket, list)
          }
          list.push(group)
        }
      }
      for (const [row, column] of group.cells ?? []) {
        this.explicit.set(`${row}:${column}`, group)
        followers += 1
      }
    }
    this.followerCount = followers
  }

  /** The group a follower at (row, column) belongs to; masters resolve too. */
  groupAt(row: number, column: number): WorkbookSharedFormulaGroup | undefined {
    const explicit = this.explicit.get(`${row}:${column}`)
    if (explicit) return explicit
    for (const group of this.spansByRowBucket.get(Math.floor(row / ROW_BUCKET)) ?? []) {
      const range = group.range
      if (
        range &&
        row >= range.startRow &&
        row <= range.endRow &&
        column >= range.startColumn &&
        column <= range.endColumn
      ) {
        return group
      }
    }
    return undefined
  }

  textAt(row: number, column: number): string | undefined {
    const group = this.groupAt(row, column)
    return group ? sharedFormulaText(group, row, column) : undefined
  }

  /** Every follower cell (file coordinates); the master is not visited. */
  forEachFollower(
    visit: (row: number, column: number, group: WorkbookSharedFormulaGroup) => void,
  ): void {
    for (const group of this.groups) {
      if (group.range) {
        const { startRow, endRow, startColumn, endColumn } = group.range
        for (let row = startRow; row <= endRow; row += 1) {
          for (let column = startColumn; column <= endColumn; column += 1) {
            if (row !== group.row || column !== group.column) visit(row, column, group)
          }
        }
      }
      for (const [row, column] of group.cells ?? []) visit(row, column, group)
    }
  }
}

/** Masters plus derived followers, as the closure analyzer consumes them. */
export function indexedFormulas(
  cells: readonly { row: number; column: number; formula?: string | undefined }[],
  shared: SharedFormulaLookup | undefined,
): { row: number; column: number; formula: string }[] {
  const out: { row: number; column: number; formula: string }[] = []
  for (const cell of cells) {
    if (cell.formula) out.push({ row: cell.row, column: cell.column, formula: cell.formula })
  }
  shared?.forEachFollower((row, column, group) => {
    out.push({ row, column, formula: sharedFormulaText(group, row, column) })
  })
  return out
}
