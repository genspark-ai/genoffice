/**
 * Shared-formula (si) materialization for the edit journal.
 *
 * A tiled clipboard paste (and fill) writes follower cells as `{si, v, s}` —
 * a shared-formula id with NO formula text; only the group's master carries
 * `f`. The journal recorded such cells as plain values at best, and the
 * follow-up `{v: null}` recalc mutation then wiped even that, so saving
 * dropped the whole pasted block (alpha feedback: copy C2:E2 → tile-paste
 * C3:E21 → save lost every cell). OOXML shared formulas cannot round-trip
 * Univer's string ids, so materialize the follower's real formula at journal
 * time instead: find the group's master cell, shift its formula by the
 * follower's offset — the same rule Univer applies when evaluating.
 */
import { LexerTreeBuilder } from '@univerjs/engine-formula'

import type { UniverRuntime } from './univer-state'

export type SharedFormulaResolver = (row: number, column: number, si: string) => string | null

interface MatrixLike {
  forValue(cb: (row: number, column: number, cell: unknown) => boolean | void): unknown
}

interface MasterCell {
  row: number
  column: number
  f: string
}

export function makeSharedFormulaResolver(
  getMatrix: () => MatrixLike | null | undefined,
  moveRefOffset: (formula: string, dx: number, dy: number) => string,
): SharedFormulaResolver {
  // one full-matrix scan per distinct si, cached for the whole batch
  const masters = new Map<string, MasterCell | null>()
  return (row, column, si) => {
    let master: MasterCell | null | undefined = masters.get(si)
    if (master === undefined) {
      let found: MasterCell | null = null
      getMatrix()?.forValue((cellRow, cellColumn, cell) => {
        const data = cell as { f?: unknown; si?: unknown } | null
        if (data && typeof data.f === 'string' && data.f.length > 0 && data.si === si) {
          found = { row: cellRow, column: cellColumn, f: data.f }
          return false
        }
      })
      master = found
      masters.set(si, master)
    }
    if (!master) return null
    if (master.row === row && master.column === column) return master.f
    return moveRefOffset(master.f, column - master.column, row - master.row)
  }
}

export function sharedFormulaResolverFor(
  runtime: UniverRuntime,
  sheetId: string,
): SharedFormulaResolver {
  const injector = runtime.univer.__getInjector()
  return makeSharedFormulaResolver(
    () =>
      runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheetId)?.getSheet().getCellMatrix(),
    (formula, dx, dy) => injector.get(LexerTreeBuilder).moveFormulaRefOffset(formula, dx, dy),
  )
}
