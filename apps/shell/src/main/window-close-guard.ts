import type { BrowserWindow } from 'electron'

import { docsQueryDirty, requestDocsClose } from '../../../docs/src/main/docs-main'
import { requestSheetsClose, resetSheetsShuttingDown } from '../../../sheets/src/main/sheets-main'
import { requestSlidesClose } from '../../../slides/src/main/slides-main'
import { requestPdfClose } from '../../../pdf/src/main/pdf-main'
import { requestMarkdownClose } from '../../../markdown/src/main/markdown-main'
import { requestHtmlClose } from '../../../html/src/main/html-main'
import type { TabManager } from './tab-manager'

/**
 * Unsaved-changes guard for the shell window: closing the whole window walks
 * every dirty sheets/pdf/slides/docs tab through the same
 * save/don't-save/cancel prompt the tab close path uses, and any cancel
 * aborts the close. docs dirtiness lives renderer-side, so any live docs tab
 * forces the async path and gets queried there (clean tabs pass through
 * without activation).
 *
 * Installed separately from index.ts (same helpers as the detached-window
 * guard) so the handler itself stays unit-testable: it takes the window and
 * the tab manager as arguments and therefore needs no Electron to run.
 */
export function installShellCloseGuard(win: BrowserWindow, manager: TabManager): void {
  let closeConfirmed = false
  // a second close event while the save prompts are still open (double ⌘Q, an
  // OS retry) must not re-enter the whole prompt chain
  let closePromptInFlight = false

  win.on('close', (event) => {
    if (closeConfirmed) return
    const dirtySheets = manager.dirtySheetsTabs()
    const dirtyPdf = manager.dirtyPdfTabs()
    const dirtyMarkdown = manager.dirtyMarkdownTabs()
    const dirtyHtml = manager.dirtyHtmlTabs()
    const dirtySlides = manager.dirtySlidesTabs()
    const docsTabs = manager.docsTabs()
    // Nothing to protect, so let this close event through untouched. Cancelling
    // it and re-issuing win.close() would also cancel an in-flight app.quit();
    // on macOS, where window-all-closed does not quit, ⌘Q with a clean
    // document would then close the only window and leave the app running
    // with none.
    if (
      dirtySheets.length === 0 &&
      dirtyPdf.length === 0 &&
      dirtyMarkdown.length === 0 &&
      dirtyHtml.length === 0 &&
      dirtySlides.length === 0 &&
      docsTabs.length === 0
    ) {
      return
    }
    // Only the branch that is about to prompt cancels the close. The debounce
    // below still needs that cancel before it returns, or a second close
    // would land on top of the prompt that is already open.
    event.preventDefault()
    if (closePromptInFlight) return
    closePromptInFlight = true
    void (async () => {
      try {
        const denied = await (async () => {
          for (const tab of dirtySheets) {
            manager.activateTab(tab.id)
            if (!(await requestSheetsClose(tab.webContents, win))) return true
          }
          for (const tab of dirtyPdf) {
            manager.activateTab(tab.id)
            if (!(await requestPdfClose(tab.webContents, win))) return true
          }
          for (const tab of dirtyMarkdown) {
            manager.activateTab(tab.id)
            if (!(await requestMarkdownClose(tab.webContents, win))) return true
          }
          for (const tab of dirtyHtml) {
            manager.activateTab(tab.id)
            if (!(await requestHtmlClose(tab.webContents, win))) return true
          }
          for (const tab of dirtySlides) {
            manager.activateTab(tab.id)
            if (!(await requestSlidesClose(tab.webContents, win))) return true
          }
          for (const tab of docsTabs) {
            if (!(await docsQueryDirty(tab.webContents))) continue
            manager.activateTab(tab.id)
            if (!(await requestDocsClose(tab.webContents, win))) return true
          }
          return false
        })()
        // a denied close vetoes any quit that was in flight: the sheets close
        // guard must prompt again on later closes instead of silently proceeding
        if (denied) resetSheetsShuttingDown()
        else {
          closeConfirmed = true
          if (!win.isDestroyed()) win.close()
        }
      } catch (err) {
        // a throw aborts the close like a denial: the window stays open, so the
        // quit that may have latched sheets as shutting down is vetoed too
        console.error('[shell] window close prompt failed:', err)
        resetSheetsShuttingDown()
      } finally {
        closePromptInFlight = false
      }
    })()
  })
}
