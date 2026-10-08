/**
 * Decision for a right-click inside the sheets document: whether our status
 * stats menu should open, and whether the event must be preventDefault()ed.
 *
 * The preventDefault matters beyond the stats menu: Univer draws its own
 * grid context menu from a window-level listener and never cancels the DOM
 * event, so a surviving event reaches Electron's webContents
 * 'context-menu' handler and the native edit menu stacks on top of it
 * (genoffice#1816). Cancelling the event suppresses the native menu entirely, which
 * is exactly how renderer-drawn menus elsewhere stay exclusive (see
 * electron-utils' context-menu.ts). Editable surfaces (formula bar, the
 * in-cell editor) keep the native cut/copy/paste menu — Univer shows nothing
 * of its own there.
 */
export interface SheetContextMenuAction {
  /** Open our statistics menu (the footer strip below the sheet tabs). */
  readonly statsMenu: boolean
  /** Cancel the event so Electron's native menu does not stack on top. */
  readonly suppressNativeMenu: boolean
}

export function sheetContextMenuAction(target: EventTarget | null): SheetContextMenuAction {
  if (!(target instanceof Element)) return { statsMenu: false, suppressNativeMenu: false }
  const footer = target.closest('#univer-container section[data-range-selector]')
  // Univer owns the footer; right-clicking it outside the tab strip opens ours.
  if (footer && !footer.firstElementChild?.contains(target))
    return { statsMenu: true, suppressNativeMenu: true }
  if (target.closest('#univer-container')) {
    // Editable surfaces keep the native cut/copy/paste menu — Univer shows
    // nothing of its own there. closest() also covers elements inside a
    // contenteditable host (the in-cell editor's internals).
    const editable =
      target.closest(
        'input, textarea, [contenteditable="true"], [contenteditable="plaintext-only"]',
      ) !== null
    return { statsMenu: false, suppressNativeMenu: !editable }
  }
  return { statsMenu: false, suppressNativeMenu: false }
}
