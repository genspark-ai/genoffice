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

export interface MoveSourceSources extends FileTargetSources {
  isDirectory: (path: string) => boolean
  isAnyRoot: (path: string) => boolean
}

/**
 * Gate for the move/drag-and-drop IPC (movePaths).
 *
 * A folder source keeps the tree's own rule — inside a root, never a root
 * itself. A file source gets the same union as the destructive file IPCs
 * instead of no gate at all: movePaths also drags out of the Recent list, so a
 * file outside every root can be legitimate, but a path the UI could never have
 * shown must not be movable by a compromised renderer.
 */
export function isMoveSource(path: string, sources: MoveSourceSources): boolean {
  if (typeof path !== 'string' || path === '') return false
  if (sources.isDirectory(path)) return sources.insideAnyRoot(path) && !sources.isAnyRoot(path)
  return isUserVisibleFile(path, sources)
}
