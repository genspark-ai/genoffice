import { statSync } from 'node:fs'
import { isAbsolute } from 'node:path'

/**
 * Reopen the documents of the previous session (genoffice#2162). On unless
 * switched off: the shell window's file tabs are stored in app-settings.json
 * as it closes and reopened at the next launch, or when the window is
 * recreated from the Dock on macOS, where closing it leaves the app running.
 *
 * The list survives the restore (#2891): the old consume cleared it before
 * the documents opened, so one crashed launch — a fluke, a late-mounting
 * network drive — lost the whole session. Now the restore is armed while it
 * runs and the arm is cleared once the documents opened: a crash after that
 * still restores next time. A crash *during* the restore counts as a failed
 * attempt; after two failed attempts in a row the list is dropped, which is
 * the crash-loop guard — a document that kills the app at launch must not be
 * restored forever.
 */
export const RESTORE_LAST_SESSION_KEY = 'restoreLastSession'
export const LAST_SESSION_DOCUMENTS_KEY = 'lastSessionDocuments'
export const LAST_SESSION_RESTORE_IN_FLIGHT_KEY = 'lastSessionRestoreInFlight'
export const LAST_SESSION_RESTORE_FAILURES_KEY = 'lastSessionRestoreFailures'

/** a hand-edited settings file must not open hundreds of tabs at launch */
const MAX_RESTORED_DOCUMENTS = 50
/** failed restore attempts in a row before the saved list is dropped */
const MAX_RESTORE_FAILURES = 2

export function restoreLastSessionEnabled(settings: Record<string, unknown>): boolean {
  return settings[RESTORE_LAST_SESSION_KEY] !== false
}

/** the saved tab paths, once each, absolute only, in tab order */
export function readLastSessionDocuments(settings: Record<string, unknown>): string[] {
  const raw = settings[LAST_SESSION_DOCUMENTS_KEY]
  if (!Array.isArray(raw)) return []
  const paths = raw.filter((p): p is string => typeof p === 'string' && isAbsolute(p))
  return [...new Set(paths)].slice(0, MAX_RESTORED_DOCUMENTS)
}

function isFile(path: string): boolean {
  try {
    return statSync(path).isFile()
  } catch {
    return false
  }
}

/**
 * What to open at launch: the previous session's documents that still exist,
 * then the files the launch itself names, so a double-clicked file ends up the
 * active tab. A file already restored is not opened twice.
 */
export function launchPathsWithLastSession(
  settings: Record<string, unknown>,
  launchPaths: readonly string[],
  exists: (path: string) => boolean = isFile,
): string[] {
  if (!restoreLastSessionEnabled(settings)) return [...launchPaths]
  const restored = readLastSessionDocuments(settings).filter(
    (p) => !launchPaths.includes(p) && exists(p),
  )
  return [...restored, ...launchPaths]
}

/** the writes the caller applies around a restore, in order */
export interface LastSessionConsumption {
  /** the paths to open: the restored session first, then the launch's own */
  paths: string[]
  /** settings updates to apply before opening (arm the restore, count a failure, drop) */
  writes: Array<Record<string, unknown>>
  /** settings updates to apply once the documents opened (disarm the restore) */
  completedWrite: Record<string, unknown> | null
  /** true when two restore attempts in a row had crashed: the list was dropped */
  dropped: boolean
}

/**
 * The launch's restore decision, as data: what to open and which settings
 * writes go before and after, so the policy is testable without Electron.
 */
export function consumeLastSessionPlan(
  settings: Record<string, unknown>,
  launchPaths: readonly string[],
  exists: (path: string) => boolean = isFile,
): LastSessionConsumption {
  const none: LastSessionConsumption = {
    paths: [...launchPaths],
    writes: [],
    completedWrite: null,
    dropped: false,
  }
  if (settings[LAST_SESSION_RESTORE_IN_FLIGHT_KEY] === true) {
    // the previous launch armed a restore and never disarmed: it crashed
    // while the documents were opening
    const failures =
      typeof settings[LAST_SESSION_RESTORE_FAILURES_KEY] === 'number'
        ? (settings[LAST_SESSION_RESTORE_FAILURES_KEY] as number)
        : 0
    if (failures >= MAX_RESTORE_FAILURES - 1) {
      // second failed attempt in a row: a document that crashes the app at
      // launch must not be restored forever — drop the list
      return {
        paths: [...launchPaths],
        writes: [
          {
            [LAST_SESSION_DOCUMENTS_KEY]: undefined,
            [LAST_SESSION_RESTORE_IN_FLIGHT_KEY]: undefined,
            [LAST_SESSION_RESTORE_FAILURES_KEY]: 0,
          },
        ],
        completedWrite: null,
        dropped: true,
      }
    }
    // first failure: give the session one more chance — a fluke (a network
    // drive still mounting, a slow disk) must not cost the whole list. The
    // failure is recorded up front, so crashing again drops the list.
    if (!restoreLastSessionEnabled(settings)) {
      return {
        paths: [...launchPaths],
        writes: [
          {
            [LAST_SESSION_DOCUMENTS_KEY]: undefined,
            [LAST_SESSION_RESTORE_IN_FLIGHT_KEY]: undefined,
            [LAST_SESSION_RESTORE_FAILURES_KEY]: 0,
          },
        ],
        completedWrite: null,
        dropped: false,
      }
    }
    const restored = readLastSessionDocuments(settings).filter(
      (p) => !launchPaths.includes(p) && exists(p),
    )
    if (restored.length === 0) {
      // nothing left to restore: clear the failed state, keep the list
      return {
        paths: [...launchPaths],
        writes: [
          {
            [LAST_SESSION_RESTORE_IN_FLIGHT_KEY]: undefined,
            [LAST_SESSION_RESTORE_FAILURES_KEY]: 0,
          },
        ],
        completedWrite: null,
        dropped: false,
      }
    }
    return {
      paths: [...restored, ...launchPaths],
      writes: [{ [LAST_SESSION_RESTORE_FAILURES_KEY]: failures + 1 }],
      completedWrite: {
        [LAST_SESSION_RESTORE_IN_FLIGHT_KEY]: false,
        [LAST_SESSION_RESTORE_FAILURES_KEY]: 0,
      },
      dropped: false,
    }
  }
  if (!restoreLastSessionEnabled(settings)) return none
  const restored = readLastSessionDocuments(settings).filter(
    (p) => !launchPaths.includes(p) && exists(p),
  )
  if (restored.length === 0) return none
  return {
    paths: [...restored, ...launchPaths],
    writes: [
      { [LAST_SESSION_RESTORE_IN_FLIGHT_KEY]: true, [LAST_SESSION_RESTORE_FAILURES_KEY]: 0 },
    ],
    completedWrite: {
      [LAST_SESSION_RESTORE_IN_FLIGHT_KEY]: false,
      [LAST_SESSION_RESTORE_FAILURES_KEY]: 0,
    },
    dropped: false,
  }
}

/** the settings write that disarms the restore once its documents opened */
export function completeLastSessionRestore(): Record<string, unknown> {
  return {
    [LAST_SESSION_RESTORE_IN_FLIGHT_KEY]: false,
    [LAST_SESSION_RESTORE_FAILURES_KEY]: 0,
  }
}
