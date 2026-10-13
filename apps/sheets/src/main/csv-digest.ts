import { statSync } from 'node:fs'

/** the two session fields `refreshCsvDigest` reads; the rest of a sheets session is irrelevant to it */
export interface CsvDigestSession {
  readonly csvSourcePath?: string
  readonly csvSourceSha?: string
}

/** dev+ino for a path, or undefined when it names nothing. Two spellings of one
 *  path get the same pair. On filesystems that report no inode (Windows
 *  FAT/exFAT report ino 0 for every file) the spelling itself is the identity
 *  again — a wrong answer beats two different files comparing equal. */
function fileIdentity(path: string): string | undefined {
  try {
    const st = statSync(path)
    if (st.ino === 0) return path
    return `${st.dev}:${st.ino}`
  } catch {
    return undefined
  }
}

/**
 * Point every session whose CSV source is the file just written at its new
 * digest, so its next Save does not read this export as an external change.
 *
 * Matched on dev+ino rather than spelling. macOS and Windows are
 * case-insensitive by default, so a session opened as `Report.CSV` and an
 * export written to `Report.csv` name one file and compare unequal — which
 * leaves the stale digest in place, and the save guard that runs next (hashing
 * content, so already case-immune) refuses the following ⌘S as an external
 * change to a file the user never touched. `path.resolve` does not close that
 * gap: it normalises `.`, `..` and separators, never case.
 *
 * Kept out of `sheets-main.ts` because that module touches `electron.app` at
 * import time, and the export path itself is only reachable through Electron's
 * file dialog.
 */
export function refreshCsvDigest<T extends CsvDigestSession>(
  sessions: Map<string, T>,
  targetPath: string,
  writtenSha: string,
): void {
  const targetId = fileIdentity(targetPath)
  if (targetId === undefined) return
  for (const [sessionId, session] of sessions) {
    if (session.csvSourcePath === undefined) continue
    const sourceId = fileIdentity(session.csvSourcePath)
    if (sourceId !== undefined && sourceId === targetId) {
      sessions.set(sessionId, { ...session, csvSourceSha: writtenSha })
    }
  }
}
