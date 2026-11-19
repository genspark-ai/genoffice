interface Sheet {
  getSheetId(): string
}
interface Workbook {
  getActiveSheet(allowNull: true): Sheet | null
  setActiveSheet(sheet: unknown): void
}
interface SheetTarget extends Sheet {
  getWorkbook?: () => Workbook
  _fWorkbook?: { setActiveSheet(sheetId: string): unknown }
}

const focusVersions = new WeakMap<object, number>()

/// Row/col property commands and SetRangeValuesCommand tail a selection op
/// onto the written sheet, and Univer's ActiveWorksheetController then
/// asynchronously activates whichever sheet the selection landed on.
/// Streaming file content into a background (even hidden) sheet must not
/// steal the active one. The activation runs after the command's promise
/// chain, so a synchronous restore alone loses the race — re-check across
/// the microtask and task queues too. Only a flip TO the patched sheet is
/// undone. Explicit MCP navigation invalidates older restores, including
/// ones aimed at the newly selected sheet.
export function keepActiveSheet<T>(worksheet: Sheet, run: () => T): T {
  const facade = worksheet as SheetTarget
  const workbook = facade.getWorkbook?.()
  const version = workbook ? focusVersions.get(workbook) : undefined
  const before = workbook?.getActiveSheet(true)
  const patchedId = worksheet.getSheetId()
  const restore = (): void => {
    if (!workbook || !before || before.getSheetId() === patchedId) return
    if (focusVersions.get(workbook) !== version) return
    const current = workbook.getActiveSheet(true)
    if (current && current !== before && current.getSheetId() === patchedId) {
      // Restore through the full SetWorksheetActiveOperation, not the bare
      // model setter: the stray activation also moved the render skeleton's
      // current sheet, and a model-only restore leaves canvas and model
      // pointing at different sheets — resolveRenderedSheetId then "heals"
      // the model back to the patched sheet, making the theft permanent.
      const fWorkbook = facade._fWorkbook
      if (fWorkbook) fWorkbook.setActiveSheet(before.getSheetId())
      else workbook.setActiveSheet(before)
    }
  }
  try {
    return run()
  } finally {
    restore()
    queueMicrotask(restore)
    setTimeout(restore, 0)
    setTimeout(restore, 60)
  }
}

/** Give explicit navigation priority over previously queued background restores. */
export function focusWorksheet(worksheet: Sheet, activate: () => void): void {
  const workbook = (worksheet as SheetTarget).getWorkbook?.()
  if (workbook) focusVersions.set(workbook, (focusVersions.get(workbook) ?? 0) + 1)
  activate()
}
