/// Convergence wrapper for the close guard. A save that raced with typing writes
/// a snapshot that is already one step behind, and reporting that as success lets
/// the window close over unpersisted edits — so keep saving until the file caught
/// up with the editor.

export interface SaveUntilPersistedDeps {
  /** Runs one full save; false means the save itself failed. */
  save: () => Promise<boolean>
  /** True when the last save wrote successfully but the editor changed mid-flight. */
  wasIncomplete: () => boolean
  /**
   * False for a document with no path yet: that save goes through the Save As
   * dialog, so retrying would prompt again — and a modal dialog means the user
   * was not typing anyway.
   */
  hasPath: () => boolean
  /** Passes before giving up (a user typing without pause never converges). */
  maxPasses?: number
}

/**
 * Serializes save passes. A call that arrives while a save is in flight waits
 * for it instead of failing (the old immediate `false` made "Save and close"
 * silently give up during a blur autosave), reuses the finished save's success
 * when `canReuse` says nothing is left to persist, and otherwise runs its own
 * pass. Callers always get the real outcome.
 */
export function createSaveSerializer(): (
  runSave: () => Promise<boolean>,
  canReuse: () => boolean,
) => Promise<boolean> {
  let tail: Promise<boolean> | null = null
  return (runSave, canReuse) => {
    // Strict FIFO chaining: every caller queues behind the current tail, so two
    // save passes can never run concurrently — the old "wait then re-check a
    // shared inFlight slot" let several waiters observe the cleared slot together
    // and start overlapping passes.
    const prior = tail
    const pass = (async () => {
      if (prior) {
        const ok = await prior.catch(() => false)
        if (ok && canReuse()) return true
      }
      return runSave()
    })()
    tail = pass
    pass
      .catch(() => false)
      .then(() => {
        if (tail === pass) tail = null
      })
    return pass
  }
}

/**
 * True only when the document is fully on disk. False means the caller must not
 * treat the save as complete — for the close guard, keep the window open.
 */
export async function saveUntilPersisted(deps: SaveUntilPersistedDeps): Promise<boolean> {
  const { save, wasIncomplete, hasPath, maxPasses = 3 } = deps
  for (let pass = 0; pass < maxPasses; pass++) {
    if (!(await save())) return false
    if (!wasIncomplete()) return true
    if (!hasPath()) return true
  }
  return false
}
