/**
 * Sheet-rename case-change fix. Univer's rename check (sheets-ui
 * nameRepeatCheck) only passes a case-sensitive self-comparison before
 * calling Workbook.checkSheetName, which matches case-insensitively — so
 * renaming Sheet1 → SHEET1 collides with itself and is rejected as a
 * duplicate. Exempt exactly that case: a candidate differing from the
 * active sheet's name only by letter case is a duplicate only if it also
 * matches some other sheet.
 */
import { Workbook } from '@univerjs/core'

interface SheetLike {
  getName(): string
}

export interface WorkbookProtoLike {
  getActiveSheet(allowNull: true): SheetLike | null | undefined
  getSheets(): SheetLike[]
  checkSheetName(name: string): boolean
  uniqueSheetName(name?: string): string
  generateNewSheetName(name: string): string
}

export function patchCheckSheetName(proto: WorkbookProtoLike): { dispose(): void } {
  const originalCheck = proto.checkSheetName
  const originalUnique = proto.uniqueSheetName
  const originalGenerate = proto.generateNewSheetName
  // uniqueSheetName/generateNewSheetName mint names for new/copied/loaded
  // sheets; the rename exemption must not weaken their dedupe (a minted
  // "Sheet2" must still count as colliding with an active "sheet2").
  let minting = false
  proto.uniqueSheetName = function (name?: string) {
    minting = true
    try {
      return originalUnique.call(this, name)
    } finally {
      minting = false
    }
  }
  proto.generateNewSheetName = function (name: string) {
    minting = true
    try {
      return originalGenerate.call(this, name)
    } finally {
      minting = false
    }
  }
  proto.checkSheetName = function (name: string) {
    if (!minting) {
      const active = this.getActiveSheet(true)
      const activeName = active?.getName()
      if (
        activeName !== undefined &&
        activeName !== name &&
        activeName.toLowerCase() === name.toLowerCase()
      ) {
        return this.getSheets().some(
          (sheet) => sheet !== active && sheet.getName().toLowerCase() === name.toLowerCase(),
        )
      }
    }
    return originalCheck.call(this, name)
  }
  return {
    dispose() {
      proto.checkSheetName = originalCheck
      proto.uniqueSheetName = originalUnique
      proto.generateNewSheetName = originalGenerate
    },
  }
}

export function installSheetRenameFix(): { dispose(): void } {
  // Patch the class prototype: no workbook unit exists yet at install time
  // (units are created later by loadSnapshotIntoUniver/loadWorkbookSkeleton).
  return patchCheckSheetName(Workbook.prototype as unknown as WorkbookProtoLike)
}
