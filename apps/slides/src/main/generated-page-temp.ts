import { lstat, readdir, unlink } from 'node:fs/promises'
import { join } from 'node:path'

export const GENERATED_PAGE_TTL_MS = 7 * 24 * 60 * 60 * 1000

/// randomUUID-shaped .pptx page files are the only things the suite writes
/// into its generated-page directories; nothing else is touched.
const GENERATED_PAGE_NAME =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}\.pptx$/i

// genoffice-cloud-pages is the retired cloud pipeline's directory: still swept so
// files from older builds expire
const GENERATED_PAGE_DIRS = ['genoffice-cloud-pages', 'genoffice-local-pages']

/**
 * Remove only expired generated-page files from the suite's app-owned temp
 * directories (startup sweep, mirroring the sheets pasted-file cleanup).
 * Markers can be redeemed more than once (retries, multi-window), so files
 * are never deleted at land time — age alone decides. Arbitrary files,
 * directories, symlinks, and similarly named nested paths are untouched;
 * per-file errors are swallowed (another window/process may hold a file).
 */
export async function cleanupExpiredGeneratedPages(
  tempRoot: string,
  now = Date.now(),
): Promise<string[]> {
  const cutoff = now - GENERATED_PAGE_TTL_MS
  const removed: string[] = []
  for (const dirName of GENERATED_PAGE_DIRS) {
    const dir = join(tempRoot, dirName)
    let entries
    try {
      entries = await readdir(dir, { withFileTypes: true })
    } catch {
      // The directory may not exist yet — nothing to sweep.
      continue
    }
    for (const entry of entries) {
      if (!entry.isFile() || !GENERATED_PAGE_NAME.test(entry.name)) continue
      const path = join(dir, entry.name)
      try {
        const info = await lstat(path)
        if (!info.isFile() || info.mtimeMs >= cutoff) continue
        await unlink(path)
        removed.push(path)
      } catch {
        // Best-effort startup cleanup; the file may already be gone.
      }
    }
  }
  return removed
}
