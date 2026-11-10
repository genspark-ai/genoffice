/**
 * The View tab's Right-to-Left sheet toggle — sheetView/@rightToLeft, Excel's
 * "Sheet Right-to-Left" display direction (column A at the right edge).
 *
 * The grid mirroring itself lives in rtl-grid-mirror, which reads the
 * worksheet config on every paint: flipping the config and rebuilding the
 * skeleton is the whole live switch. The rebuild re-runs the mirrored
 * viewport anchors, which on a sheet's first RTL activation wipe any
 * scroll state poisoned by LTR clamping back to the flush-right home and
 * repaint; flipping off restores the LTR margins and replays the recorded
 * scroll state, landing back at column A on the left.
 *
 * Univer 0.25.1 ships the right-to-left command and mutation but never
 * registers either in the plugin's tables (ToggleGridlinesMutation ships
 * registered; this pair does not), so a dispatch would silently no-op —
 * installRightToLeftCommand registers them once per session.
 */
import { BooleanNumber, ICommandService } from '@univerjs/core'
import { IRenderManagerService } from '@univerjs/engine-render'
import { SetWorksheetRightToLeftCommand, SetWorksheetRightToLeftMutation } from '@univerjs/sheets'
import { SheetSkeletonManagerService } from '@univerjs/sheets-ui'

import type { UniverRuntime } from './univer-state'

interface RtlWorksheetLike {
  getSheet(): { getConfig(): { rightToLeft?: BooleanNumber } }
}

interface InjectorLike {
  get<T>(token: unknown): T
}

interface SkeletonManagerLike {
  reCalculate(): void
}

interface RenderUnitLike {
  with<T>(token: unknown): T
}

export function sheetIsRightToLeft(worksheet: RtlWorksheetLike | null | undefined): boolean {
  try {
    return worksheet?.getSheet().getConfig().rightToLeft === BooleanNumber.TRUE
  } catch {
    return false
  }
}

const installedRuntimes = new WeakSet<object>()

export function installRightToLeftCommand(runtime: UniverRuntime): void {
  if (installedRuntimes.has(runtime.univer)) return
  try {
    const injector = (
      runtime.univer as unknown as { __getInjector(): InjectorLike }
    ).__getInjector()
    const commands = injector.get<{
      registerCommand(command: unknown): unknown
    }>(ICommandService)
    commands.registerCommand(SetWorksheetRightToLeftCommand)
    commands.registerCommand(SetWorksheetRightToLeftMutation)
    installedRuntimes.add(runtime.univer)
  } catch {
    // The command service is not up yet (still booting); the next workbook
    // open retries the registration.
  }
}

/// Flip the active sheet's grid direction. The command wraps the mutation
/// with its own undo entry; the mutation listener refreshes the render (and
/// journals the state), so Univer's undo/redo of the flip re-renders too.
export function applyRightToLeftView(runtime: UniverRuntime, rtl: boolean): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return
  runtime.univerAPI.syncExecuteCommand('sheet.command.set-worksheet-right-to-left', {
    unitId: workbook.getId(),
    subUnitId: worksheet.getSheetId(),
    rightToLeft: rtl ? BooleanNumber.TRUE : BooleanNumber.FALSE,
  })
}

/// Rebuild the sheet skeleton so the mirrored viewport anchors re-run after
/// any right-to-left mutation — the ribbon's own flip, and Univer's undo or
/// redo of one, which dispatches the mutation without the command.
export function refreshSheetRender(runtime: UniverRuntime | null | undefined): void {
  if (!runtime) return
  try {
    const workbook = runtime.univerAPI.getActiveWorkbook()
    if (!workbook) return
    const injector = (
      runtime.univer as unknown as { __getInjector(): InjectorLike }
    ).__getInjector()
    const render = injector
      .get<{ getRenderById(id: string): RenderUnitLike | null }>(IRenderManagerService)
      .getRenderById(workbook.getId())
    render?.with<SkeletonManagerLike>(SheetSkeletonManagerService).reCalculate()
  } catch {
    // No render unit yet (still booting) — the next skeleton rebuild picks
    // the flag up.
  }
}
