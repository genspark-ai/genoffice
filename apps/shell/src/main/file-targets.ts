export interface FileTargetSources {
  insideAnyRoot: (path: string) => boolean
  trackedPaths: readonly string[]
}

/** A recent may live outside every root, so roots alone cannot gate the file IPCs. */
export function isUserVisibleFile(path: string, sources: FileTargetSources): boolean {
  if (typeof path !== 'string' || path === '') return false
  if (sources.insideAnyRoot(path)) return true
  return sources.trackedPaths.includes(path)
}

export interface MoveSourceSources extends FileTargetSources {
  isDirectory: (path: string) => boolean
  isAnyRoot: (path: string) => boolean
}

export function isMoveSource(path: string, sources: MoveSourceSources): boolean {
  if (typeof path !== 'string' || path === '') return false
  if (sources.isDirectory(path)) return sources.insideAnyRoot(path) && !sources.isAnyRoot(path)
  return isUserVisibleFile(path, sources)
}
