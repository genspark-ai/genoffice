/**
 * Excel's veryHidden sheets stay hidden in Univer (formulas still see them)
 * but no UI path may unhide them: the sheet-bar menu, Unhide dialog and the
 * facade all end in SetWorksheetShowCommand, which is vetoed here.
 */
import { CustomCommandExecutionError, ICommandService } from '@univerjs/core'

import type { WorkbookFile } from '../shared/desktop-api'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

const SHOW_COMMAND = 'sheet.command.set-worksheet-show'

export function isVeryHiddenSheet(file: WorkbookFile | undefined, sheetId: string): boolean {
  return file?.sheets.find((sheet) => sheet.id === sheetId)?.veryHidden === true
}

export function vetoesUnhide(
  file: WorkbookFile | undefined,
  command: { id: string; params?: unknown },
): boolean {
  if (command.id !== SHOW_COMMAND) return false
  const subUnitId = (command.params as { subUnitId?: unknown } | undefined)?.subUnitId
  return typeof subUnitId === 'string' && isVeryHiddenSheet(file, subUnitId)
}

/// Cancels the command being dispatched. Univer pops its execution stack only
/// on the success path, so the entry pushed for this dispatch is removed by
/// identity first (private field; if renamed the veto still works and only
/// this cleanup degrades).
export function vetoUniverCommand(
  commandService: ICommandService,
  command: object,
  reason: string,
): never {
  const stack = (commandService as unknown as { _commandExecutionStack?: unknown[] })
    ._commandExecutionStack
  const index = stack?.indexOf(command) ?? -1
  if (index >= 0) stack?.splice(index, 1)
  throw new CustomCommandExecutionError(reason)
}

export function installVeryHiddenSheetGuard(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): void {
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  commandService.beforeCommandExecuted((command) => {
    if (vetoesUnhide(lazyWorkbookRef.current?.file, command)) {
      vetoUniverCommand(commandService, command, 'veryHidden sheet')
    }
  })
}
