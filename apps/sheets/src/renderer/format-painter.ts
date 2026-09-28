/**
 * Excel-style Format Painter button semantics on top of Univer's painter
 * service: a single click paints the next selection once, a double-click
 * keeps the painter on until Esc or another click on the button.
 */
import { CommandType, ICommandService, toDisposable } from '@univerjs/core'
import type { IDisposable } from '@univerjs/core'
import { IShortcutService, KeyCode } from '@univerjs/ui'
import {
  FormatPainterStatus,
  IFormatPainterService,
  SetFormatPainterOperation,
  whenSheetEditorFocused,
} from '@univerjs/sheets-ui'
import type { UniverRuntime } from './univer-state'

export type FormatPainterClick = 'once' | 'lock' | 'cancel' | 'ignore'

const CANCEL_FORMAT_PAINTER_ID = 'genoffice.command.cancel-format-painter'

/**
 * Maps a ribbon click to the painter transition. A double-click reaches
 * the button as click(detail 1) → click(detail 2), so the first click has
 * already armed the one-shot painter when the second one arrives and only
 * needs to escalate it; a second click that follows a cancel is ignored.
 */
export function formatPainterClickAction(
  status: FormatPainterStatus,
  detail: number,
): FormatPainterClick {
  if (detail >= 2) return status === FormatPainterStatus.ONCE ? 'lock' : 'ignore'
  return status === FormatPainterStatus.OFF ? 'once' : 'cancel'
}

const STATUS_AFTER_CLICK: Record<Exclude<FormatPainterClick, 'ignore'>, FormatPainterStatus> = {
  once: FormatPainterStatus.ONCE,
  lock: FormatPainterStatus.INFINITE,
  cancel: FormatPainterStatus.OFF,
}

export function formatPainterTurnedOff(
  previous: FormatPainterStatus,
  next: FormatPainterStatus,
): boolean {
  return previous !== FormatPainterStatus.OFF && next === FormatPainterStatus.OFF
}

export function applyFormatPainterClick(
  runtime: UniverRuntime,
  detail: number,
): FormatPainterClick {
  const injector = runtime.univer.__getInjector()
  const painter = injector.get(IFormatPainterService)
  const action = formatPainterClickAction(painter.getStatus(), detail)
  if (action === 'ignore') return action
  void injector
    .get(ICommandService)
    .executeCommand(SetFormatPainterOperation.id, { status: STATUS_AFTER_CLICK[action] })
  return action
}

/**
 * Mirrors the painter status into the ribbon and binds Esc in the grid to
 * cancel an active painter (Univer only uses Esc to leave the cell editor).
 * `turnedOff` flags the off transition from any path (button, Esc, the
 * one-shot paint completing) so the status line can drop the painter hint.
 */
export function installFormatPainter(
  runtime: UniverRuntime,
  onStatusChange: (active: boolean, turnedOff: boolean) => void,
): IDisposable {
  const injector = runtime.univer.__getInjector()
  const painter = injector.get(IFormatPainterService)
  const commandService = injector.get(ICommandService)
  const shortcutService = injector.get(IShortcutService)

  let previous = painter.getStatus()
  const subscription = painter.status$.subscribe((status) => {
    onStatusChange(status !== FormatPainterStatus.OFF, formatPainterTurnedOff(previous, status))
    previous = status
  })
  const commandDisposable = commandService.registerCommand({
    id: CANCEL_FORMAT_PAINTER_ID,
    type: CommandType.COMMAND,
    handler: (accessor) =>
      accessor
        .get(ICommandService)
        .executeCommand(SetFormatPainterOperation.id, { status: FormatPainterStatus.OFF }),
  })
  const shortcutDisposable = shortcutService.registerShortcut({
    id: CANCEL_FORMAT_PAINTER_ID,
    binding: KeyCode.ESC,
    priority: 100,
    preconditions: (contextService) =>
      painter.getStatus() !== FormatPainterStatus.OFF && whenSheetEditorFocused(contextService),
  })

  return toDisposable(() => {
    subscription.unsubscribe()
    commandDisposable.dispose()
    shortcutDisposable.dispose()
  })
}
