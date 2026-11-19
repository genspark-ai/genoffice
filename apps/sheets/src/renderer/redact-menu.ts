/**
 * The grid's right-click item that withholds the selected cells from the AI.
 *
 * The cell context menu belongs to Univer, not to this app, so the item is
 * contributed through `IMenuManagerService.mergeMenu` — the sanctioned
 * extension point — rather than by replacing the menu. That keeps every
 * built-in item (insert rows, sort, clear formatting…) working untouched, and
 * it is why the label is a Univer locale key: Univer resolves `title` through
 * its own LocaleService at render, so the item follows a language switch for
 * free and is worded exactly like its neighbours.
 *
 * The gesture is one item doing both jobs, as in slides: marking a fresh
 * selection asks what it stands for, and a selection that is already withheld
 * offers to stop withholding it. The command only reports the selection —
 * deciding what to do with it belongs to the app, which owns the dialog.
 */
import { CommandType, ICommandService, UniverInstanceType, toDisposable } from '@univerjs/core'
import {
  ContextMenuGroup,
  ContextMenuPosition,
  IMenuManagerService,
  MenuItemType,
  getMenuHiddenObservable,
  type IMenuButtonItem,
} from '@univerjs/ui'

import type { UniverRuntime } from './univer-state'

export const HIDE_SELECTION_COMMAND_ID = 'genoffice.sheets.command.hide-selection-from-ai'

/** The rectangle a right-click acted on, in zero-based screen coordinates. */
export interface SelectionRequest {
  readonly sheetId: string
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
  readonly isSingleCell: boolean
}

export function installRedactMenu(
  runtime: UniverRuntime,
  onRequest: (selection: SelectionRequest) => void,
) {
  const injector = runtime.univer.__getInjector()
  const commandService = injector.get(ICommandService)

  const commandDisposable = commandService.registerCommand({
    id: HIDE_SELECTION_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: () => {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      const worksheet = workbook?.getActiveSheet()
      // No active range means the click landed outside the grid (a header, the
      // formula bar); there is nothing to withhold, so the item does nothing
      // rather than guessing a target.
      const range = worksheet?.getActiveRange()?.getRange()
      if (!worksheet || !range) return false
      onRequest({
        sheetId: worksheet.getSheetId(),
        startRow: range.startRow,
        endRow: range.endRow,
        startColumn: range.startColumn,
        endColumn: range.endColumn,
        isSingleCell: range.startRow === range.endRow && range.startColumn === range.endColumn,
      })
      return true
    },
  })

  injector.get(IMenuManagerService).mergeMenu({
    [ContextMenuPosition.MAIN_AREA]: {
      [ContextMenuGroup.OTHERS]: {
        [HIDE_SELECTION_COMMAND_ID]: {
          order: 100,
          menuItemFactory: (accessor): IMenuButtonItem => ({
            id: HIDE_SELECTION_COMMAND_ID,
            type: MenuItemType.BUTTON,
            title: 'sheets-ui.rightClick.hideSelectionFromAi',
            hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_SHEET),
          }),
        },
      },
    },
  })

  return toDisposable(() => {
    commandDisposable.dispose()
    // mergeMenu has no inverse, so the item outlives this disposable; the
    // command it invokes is disposed, which is enough to make it inert.
  })
}
