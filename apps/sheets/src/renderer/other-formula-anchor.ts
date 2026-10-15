/**
 * Univer registers conditional-formatting and data-validation formulas as
 * "other formulas": one dependency tree at the first covered cell plus
 * virtual trees offset from it. The first tree never gets its row/column
 * set, so it sits at (-1, -1) and every virtual tree inherits the shift —
 * argument-less ROW()/COLUMN() answer 0-based values (`=MOD(ROW(),2)=0`
 * paints row 1) and dirty-range hit tests miss by one cell. Anchor the
 * first tree at its range's top-left right after registration.
 */
import { FormulaDependencyGenerator } from '@univerjs/engine-formula'

export interface OtherFormulaTreeLike {
  isVirtual: boolean
  formulaId?: string | null
  unitId: string
  subUnitId: string
  row: number
  column: number
}

export type OtherFormulaDataLike = Record<
  string,
  | Record<
      string,
      | Record<
          string,
          { ranges?: readonly { startRow: number; startColumn: number }[] } | undefined
        >
      | undefined
    >
  | undefined
>

export function anchorOtherFormulaTrees(
  treeList: readonly OtherFormulaTreeLike[],
  otherFormulaData: OtherFormulaDataLike,
  from: number,
): void {
  for (let i = from; i < treeList.length; i += 1) {
    const tree = treeList[i]
    if (!tree || tree.isVirtual || tree.formulaId == null || tree.row !== -1) continue
    const range = otherFormulaData[tree.unitId]?.[tree.subUnitId]?.[tree.formulaId]?.ranges?.[0]
    if (!range) continue
    tree.row = range.startRow
    tree.column = range.startColumn
  }
}

type RegisterOtherFormulas = (
  otherFormulaData: OtherFormulaDataLike,
  otherFormulaDataKeys: string[],
  treeList: OtherFormulaTreeLike[],
) => void

let installed = false

export function installOtherFormulaAnchor(): void {
  if (installed) return
  installed = true
  const proto = FormulaDependencyGenerator.prototype as unknown as {
    _registerOtherFormulas?: RegisterOtherFormulas
  }
  const original = proto._registerOtherFormulas
  if (typeof original !== 'function') return
  proto._registerOtherFormulas = function (otherFormulaData, otherFormulaDataKeys, treeList) {
    const from = treeList.length
    original.call(this, otherFormulaData, otherFormulaDataKeys, treeList)
    anchorOtherFormulaTrees(treeList, otherFormulaData, from)
  }
}
