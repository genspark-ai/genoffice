import { app, type BrowserWindow } from 'electron'

import { appForKind, resetSheetsShuttingDown } from './app-registry'
import type { DocumentTabKind } from '../shared/tabs-api'
import type { TabManager } from './tab-manager'

/**
 * The sweep walks the apps in this fixed order — the same order the guard has
 * always prompted in — and asks each one's registered close guard. Kinds whose
 * dirty answer needs a renderer round-trip (docs) are swept with every live
 * tab instead of a pre-filtered dirty list.
 */
const CLOSE_PROMPT_ORDER: readonly DocumentTabKind[] = [
  'sheets',
  'pdf',
  'markdown',
  'html',
  'slides',
  'docs',
]

/**
 * Unsaved-changes guard for the shell window: closing the whole window walks
 * every dirty document tab through the same save/don't-save/cancel prompt the
 * tab close path uses (the prompts themselves come from the app registry), and
 * any cancel aborts the close. docs dirtiness lives renderer-side, so any live
 * docs tab forces the async path and gets queried there (clean tabs pass
 * through without activation).
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
 * Electron to run.
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

export function installShellCloseGuard(
  win: BrowserWindow,
  manager: TabManager,
  quitTracker: QuitTracker = shellQuitTracker(),
): void {
  let closeConfirmed = false
  // a second close event while the save prompts are still open (double ⌘Q, an
  // OS retry) must not re-enter the whole prompt chain
  let closePromptInFlight = false
  let holdingQuit = false

  win.on('close', (event) => {
    if (closeConfirmed) return
    // One sweep list per app, computed now so the walk below uses exactly the
    // tabs that were dirty (or, for the async family, live) at close time: the
    // sync-dirty families get their dirty tabs, docs gets every live tab
    // because its dirtiness is only known after a renderer round-trip.
    const sweep = CLOSE_PROMPT_ORDER.map((kind) => ({
      kind,
      tabs: appForKind(kind)?.queryDirty ? manager.tabsOfKind(kind) : manager.dirtyTabsOfKind(kind),
    }))
    // Nothing to protect, so let this close event through untouched. Cancelling
    // it and re-issuing win.close() would also cancel an in-flight app.quit();
    // on macOS, where window-all-closed does not quit, ⌘Q with a clean
    // document would then close the only window and leave the app running
    // with none.
    if (sweep.every((entry) => entry.tabs.length === 0)) return
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
        const denied = await (async () => {
          for (const entry of sweep) {
            const descriptor = appForKind(entry.kind)
            if (!descriptor) continue
            for (const tab of entry.tabs) {
              if (descriptor.queryDirty && !(await descriptor.queryDirty(tab.webContents))) continue
              manager.activateTab(tab.id)
              if (!(await descriptor.requestClose(tab.webContents, win))) return true
            }
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
