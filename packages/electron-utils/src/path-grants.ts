/// Per-view path grants for renderer-facing READ channels (attachment text /
/// image reads, document opens): the paths a view's user actually put in
/// front of the app — picked through a dialog, dropped from the OS, opened
/// from a listing — are recorded at the moment the main process sees them,
/// and the read handlers verify against that record.
///
/// Without the check, every such handler trusts the path the renderer sends
/// and only an extension whitelist stands between a compromised renderer and
/// arbitrary file reads (the defense-in-depth gap the #2746 review left as
/// follow-up; the html preview iframe cannot reach the contextBridge on this
/// Electron, so these reads are renderer-compromise hardening, not
/// document-reachable).
///
/// Grants are per webContents id: one editor view's picks never widen
/// another view's reads. The ledger holds no Electron imports — callers wire
/// `release` to their own destroyed handlers and keep this unit-testable.

import { resolve } from 'node:path'

/** grants per view; soft-capped so a runaway caller cannot grow it forever */
const MAX_GRANTS_PER_VIEW = 512

export interface PathGrants {
  /** record paths this view legitimately put in front of the app */
  grant(wcId: number, paths: readonly string[]): void
  /** whether this view may read this exact path */
  has(wcId: number, path: string): boolean
  /** drop a view's grants (webContents destroyed / tab closed) */
  release(wcId: number): void
}

/** Symbol.for: every app bundle ships its own copy of this module, and the
 *  shell hosts all of them in one process — the ledger must be ONE table or
 *  a path the shell saw (drop, dialog) is invisible to the editor's reads. */
const SHARED = Symbol.for('genoffice.shared-path-grants')

/** The process-wide ledger the drop bridge, dialog picks and read handlers
 *  share. */
export function sharedPathGrants(): PathGrants {
  const holder = globalThis as Record<symbol, PathGrants | undefined>
  holder[SHARED] ??= createPathGrants()
  return holder[SHARED]!
}

export function createPathGrants(): PathGrants {
  const byView = new Map<number, Set<string>>()
  return {
    grant(wcId, paths) {
      let set = byView.get(wcId)
      if (!set) {
        set = new Set()
        byView.set(wcId, set)
      }
      for (const raw of paths) {
        if (typeof raw !== 'string' || raw.length === 0 || raw.length > 4096) continue
        if (set.size >= MAX_GRANTS_PER_VIEW) set.delete(set.values().next().value!)
        set.add(resolve(raw))
      }
    },
    has(wcId, path) {
      const set = byView.get(wcId)
      return set !== undefined && set.has(resolve(path))
    },
    release(wcId) {
      byView.delete(wcId)
    },
  }
}
