import { app, dialog, type BrowserWindow } from 'electron'

import { format } from '@genoffice/i18n'

import { docsQueryDirty, requestDocsClose } from '../../../docs/src/main/docs-main'
import { requestSheetsClose, resetSheetsShuttingDown } from '../../../sheets/src/main/sheets-main'
import { requestSlidesClose } from '../../../slides/src/main/slides-main'
import { requestPdfClose } from '../../../pdf/src/main/pdf-main'
import { requestMarkdownClose } from '../../../markdown/src/main/markdown-main'
import { requestHtmlClose } from '../../../html/src/main/html-main'
import { MIN_DOCUMENTS_TO_CONFIRM_CLOSE } from './close-confirm'
import type { TabManager } from './tab-manager'

/**
 * Unsaved-changes guard for the shell window: closing the whole window walks
 * every dirty sheets/pdf/slides/docs tab through the same
 * save/don't-save/cancel prompt the tab close path uses, and any cancel
 * aborts the close. docs dirtiness lives renderer-side, so any live docs tab
 * forces the async path and gets queried there (clean tabs pass through
 * without activation).
 *
 * On top of that (genoffice#2490), closing the window with more than one
 * document tab open is confirmed first with a "Close N documents?" dialog, so
 * the OS close button cannot toss several clean documents at once. Cancelling
 * that dialog keeps every tab open — the user can then close just one tab
 * from the strip. It can be opted out with "Don't ask again and close".
 *
 * Cancelling the close event also cancels an in-flight app.quit(), and on
 * macOS window-all-closed does not quit, so once the prompts have all been
 * confirmed a close that began as a quit must be finished with app.quit()
 * rather than win.close(): the latter closes the only window and leaves the
 * app running with none (genoffice#2098, ⌘Q needing two presses whenever a
 * Docs tab is open).
 *
 * Installed separately from index.ts (same helpers as the detached-window
 * guard) so the handler itself stays unit-testable: it takes the window, the
 * tab manager and the quit tracker as arguments and therefore needs no
 * Electron to run. The multi-document confirmation is injectable so its tests
 * need no native dialog either.
 */
export interface QuitTracker {
  /** a quit is in progress: between before-quit and its windows' close events */
  quitting(): boolean
  /** keep quitting() alive past the close event while prompts run */
  hold(): void
  release(): void
  /** the quit was cancelled: a prompt was denied */
  reset(): void
  quit(): void
}

/**
 * Electron fires no event for a cancelled quit, so the flag may not stay
 * latched: it is true only during the synchronous window-close pass that
 * app.quit() runs (cleared on the next macrotask) unless a guard that
 * prevented one of those closes holds it while its prompts are open.
 */
function electronQuitTracker(): QuitTracker {
  let quitting = false
  let holds = 0
  app.on('before-quit', () => {
    quitting = true
    setImmediate(() => {
      if (holds === 0) quitting = false
    })
  })
  return {
    quitting: () => quitting,
    hold: () => {
      holds += 1
    },
    release: () => {
      holds -= 1
      if (holds === 0) quitting = false
    },
    reset: () => {
      quitting = false
    },
    quit: () => app.quit(),
  }
}

let shared: QuitTracker | undefined

export function shellQuitTracker(): QuitTracker {
  return (shared ??= electronQuitTracker())
}

/** collaborators for the multi-document confirmation; all optional in production
 *  (the native dialog, the settings file, the main locale table) */
export interface ShellCloseGuardOptions {
  showMessageBox?: (
    win: BrowserWindow,
    options: {
      type: 'question'
      message: string
      detail: string
      buttons: string[]
      defaultId: number
      cancelId: number
    },
  ) => Promise<{ response: number }>
  /** read the opt-out on every close; defaults to always asking */
  isCloseConfirmEnabled?: () => boolean
  /** persist the opt-out after "Don't ask again and close" */
  onCloseConfirmDisabled?: () => void
  /** main-process translator; defaults to the English strings below */
  t?: (key: string, params?: Record<string, string | number>) => string
}

const CLOSE_CONFIRM_EN: Record<string, string> = {
  closeAllTitle: 'Close {count} documents?',
  closeAllDetail:
    'All open documents in this window will be closed. Unsaved changes will still be prompted before closing.',
  closeAllClose: 'Close all',
  closeAllDontAsk: "Don't ask again and close",
  btnCancel: 'Cancel',
}

function defaultCloseConfirmT(key: string, params?: Record<string, string | number>): string {
  return format(CLOSE_CONFIRM_EN[key] ?? key, params)
}

/** document tabs sharing the window (Home is not one); the manager's list is the
 *  source of truth, so a stub without one counts as no documents */
function openDocumentCount(manager: TabManager): number {
  try {
    return manager.list().filter((tab) => tab.closable).length
  } catch {
    return 0
  }
}

export function installShellCloseGuard(
  win: BrowserWindow,
  manager: TabManager,
  quitTracker: QuitTracker = shellQuitTracker(),
  options: ShellCloseGuardOptions = {},
): void {
  let closeConfirmed = false
  // a second close event while the save prompts are still open (double ⌘Q, an
  // OS retry) must not re-enter the whole prompt chain
  let closePromptInFlight = false
  let holdingQuit = false

  win.on('close', (event) => {
    if (closeConfirmed) return
    const dirtySheets = manager.dirtySheetsTabs()
    const dirtyPdf = manager.dirtyPdfTabs()
    const dirtyMarkdown = manager.dirtyMarkdownTabs()
    const dirtyHtml = manager.dirtyHtmlTabs()
    const dirtySlides = manager.dirtySlidesTabs()
    const docsTabs = manager.docsTabs()
    const openCount = openDocumentCount(manager)
    // The multi-document confirmation guards every clean document at once, so
    // it stands on its own: a window with several clean tabs has nothing dirty
    // yet still needs to be asked before all of them go away.
    const needsCloseConfirm =
      openCount >= MIN_DOCUMENTS_TO_CONFIRM_CLOSE && (options.isCloseConfirmEnabled?.() ?? true)
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
      docsTabs.length === 0 &&
      !needsCloseConfirm
    ) {
      return
    }
    // Only the branch that is about to prompt cancels the close. The debounce
    // below still needs that cancel before it returns, or a second close
    // would land on top of the prompt that is already open.
    event.preventDefault()
    // a ⌘Q that lands while a red-button prompt is already open still turns
    // that close into a quit
    if (quitTracker.quitting() && !holdingQuit) {
      holdingQuit = true
      quitTracker.hold()
    }
    if (closePromptInFlight) return
    closePromptInFlight = true
    void (async () => {
      try {
        // Ask about the whole window before any per-document save prompt: a
        // cancel here keeps everything open and skips the save walk entirely.
        if (needsCloseConfirm) {
          const t = options.t ?? defaultCloseConfirmT
          const showMessageBox =
            options.showMessageBox ?? ((w, opts) => dialog.showMessageBox(w, opts))
          const { response } = await showMessageBox(win, {
            type: 'question',
            message: t('closeAllTitle', { count: openCount }),
            detail: t('closeAllDetail'),
            buttons: [t('closeAllClose'), t('closeAllDontAsk'), t('btnCancel')],
            defaultId: 0,
            cancelId: 2,
          })
          // "Don't ask again" closes like "Close all" but also persists the
          // opt-out, so the next close goes straight to the save prompts.
          if (response === 1) options.onCloseConfirmDisabled?.()
          // cancelling keeps every tab open, like a denied save prompt
          if (response !== 0 && response !== 1) {
            quitTracker.reset()
            resetSheetsShuttingDown()
            return
          }
        }
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
        if (denied) {
          quitTracker.reset()
          resetSheetsShuttingDown()
        } else {
          closeConfirmed = true
          if (win.isDestroyed()) return
          // read now, not at the close event: a detached window may have vetoed
          // the quit in the meantime
          if (quitTracker.quitting()) quitTracker.quit()
          else win.close()
        }
      } catch (err) {
        // a throw aborts the close like a denial: the window stays open, so the
        // quit that may have latched sheets as shutting down is vetoed too
        console.error('[shell] window close prompt failed:', err)
        quitTracker.reset()
        resetSheetsShuttingDown()
      } finally {
        closePromptInFlight = false
        if (holdingQuit) {
          holdingQuit = false
          quitTracker.release()
        }
      }
    })()
  })
}
