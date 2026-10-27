/**
 * Register Univer's filter-toggle command, which 0.25.x leaves unregistered.
 *
 * `SmartToggleSheetsFilterCommand` (`sheet.command.smart-toggle-filter`) is
 * defined and exported by `@univerjs/sheets-filter`, and Univer's own filter
 * menu and its ⇧⌘L shortcut both point at it — but
 * `UniverSheetsFilterPlugin` never passes it to `registerCommand`. The plugin
 * registers the five commands around it (set range, remove, set criteria,
 * clear criteria, recalculate) and skips the toggle.
 *
 * A click on "Filter" therefore asks the command service for a command that
 * does not exist: nothing is registered, nothing runs, and nothing is said.
 * The sheet keeps exactly the state it had, which reads as the button being
 * dead (user report: select a header row, Sort & Filter ▸ Filter, and
 * nothing happens).
 *
 * Registering it here is the whole fix: the handler is upstream's, and it
 * removes the filter when the sheet already has one and adds it otherwise.
 * Univer's own menu item and shortcut work again for the same reason.
 */
import { ICommandService } from '@univerjs/core'
import { SmartToggleSheetsFilterCommand } from '@univerjs/sheets-filter'

import type { UniverRuntime } from './univer-state'

/**
 * Make sure the command service can run the filter toggle.
 *
 * A newer Univer may register the command itself. Registering it again would
 * replace a working command with the same handler, but there is no reason to
 * take that risk: only fill the gap.
 */
export function registerFilterToggleCommand(commandService: ICommandService): {
  dispose(): void
} {
  if (commandService.hasCommand(SmartToggleSheetsFilterCommand.id)) {
    return { dispose: () => {} }
  }
  const disposable = commandService.registerCommand(SmartToggleSheetsFilterCommand)
  return { dispose: () => disposable.dispose() }
}

export function installFilterToggleCommand(runtime: UniverRuntime): { dispose(): void } {
  return registerFilterToggleCommand(runtime.univer.__getInjector().get(ICommandService))
}
