/** Disk snapshot recorded at the last read/write of a document path. */
export interface DiskFileState {
  mtimeMs: number
  size: number
  hash: string
}

/**
 * True when the file on disk no longer matches the recorded state, i.e. another
 * program wrote it since we last read/saved it. No record (path never tracked)
 * or a missing file (deleted externally) is not a conflict — the save proceeds
 * and recreates the file. The hash read only runs when mtime+size already
 * disagree, so the common no-conflict save never rereads the file.
 *
 * Handles edge cases gracefully:
 * - Invalid or corrupt timestamps (logs warning, falls back to hash comparison)
 * - Hash read failures (returns false to allow save, avoiding false-positive blocks)
 * - Network drive clock skew (hash-based validation catches content changes)
 */
export async function isExternallyModified(
  recorded: DiskFileState | undefined,
  current: { mtimeMs: number; size: number } | null,
  readHash: () => string | null | Promise<string | null>,
): Promise<boolean> {
  if (!recorded || !current) return false

  // Validate timestamps: catch invalid mtimes from network drives or filesystem issues
  if (!isValidTimestamp(recorded.mtimeMs) || !isValidTimestamp(current.mtimeMs)) {
    console.warn(
      `[external-change] Invalid mtime detected: recorded=${recorded.mtimeMs}, current=${current.mtimeMs}. Falling back to hash comparison.`,
    )
    // Clock skew or invalid timestamp: skip mtime check, go straight to hash
    const hash = await readHash()
    return hash !== null && hash !== recorded.hash
  }

  if (current.mtimeMs === recorded.mtimeMs && current.size === recorded.size) return false

  // Hash comparison for content verification
  try {
    const hash = await readHash()
    return hash !== null && hash !== recorded.hash
  } catch (error) {
    // Hash read failed (file locked, permission error, network issue)
    console.error(
      `[external-change] Hash read failed, allowing save to proceed:`,
      error instanceof Error ? error.message : String(error),
    )
    // Fail-safe: return false to avoid blocking legitimate saves with false-positive conflict
    return false
  }
}

/**
 * Validate filesystem timestamp: catches invalid values from clock skew,
 * network drives, or corrupted filesystem metadata.
 */
function isValidTimestamp(mtimeMs: number): boolean {
  // Valid range: Unix epoch (1970-01-01) to far future (year 3000)
  // Rejects: 0, negative, NaN, Infinity, or absurdly far future
  return (
    Number.isFinite(mtimeMs) && mtimeMs > 0 && mtimeMs < 32503680000000 // 3000-01-01 in ms
  )
}
