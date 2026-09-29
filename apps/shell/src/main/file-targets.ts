/**
 * Gate for the destructive/editing file IPCs (rename, duplicate, delete).
 *
 * The Home UI can only show files from specific places: the folder roots,
 * recents, stars, cloud-project files, the slides recents, and the paths of
 * open or detached tabs. A renderer is supposed to pass back one of those —
 * but a compromised renderer can pass any absolute path, and without a gate
 * the main process would rename, duplicate or trash it (the audit's
 * "destructive file IPC is not root-scoped"). `insideAnyRoot` alone is not
 * enough: a legitimate recent can live outside every root (a file opened from
 * Downloads), so the check is membership in the union of every tracked
 * source, not just the roots.
 */

export interface FileTargetSources {
  /** every folder-root path (default save dir + extras) */
  insideAnyRoot: (path: string) => boolean
  /** recents, stars, project files, slides recents, open and detached tabs */
  trackedPaths: readonly string[]
}

export function isUserVisibleFile(path: string, sources: FileTargetSources): boolean {
  if (typeof path !== 'string' || path === '') return false
  if (sources.insideAnyRoot(path)) return true
  return sources.trackedPaths.includes(path)
}
