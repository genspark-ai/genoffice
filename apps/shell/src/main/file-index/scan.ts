import { readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { isHiddenEntry, isSupportedTreeFile } from '../folder-tree'

export interface ScannedFile {
  path: string
  mtimeMs: number
  sizeBytes: number
}

export interface ScanResult {
  files: ScannedFile[]
  /** the walk hit a budget, so `files` is not the full set under the root */
  truncated: boolean
}

export const SCAN_MAX_DEPTH = 32
export const SCAN_MAX_FILES = 200_000

/**
 * Every supported, visible file under `root` with the stat fields the index keys
 * on. Symlinked directories are not followed (a Dirent reports them as
 * symlinks, not directories), and the walk is bounded by depth and file count.
 */
export function scanFiles(
  root: string,
  limits: { maxDepth?: number; maxFiles?: number } = {},
): ScanResult {
  const maxDepth = limits.maxDepth ?? SCAN_MAX_DEPTH
  const maxFiles = limits.maxFiles ?? SCAN_MAX_FILES
  const out: ScannedFile[] = []
  let truncated = false
  const walk = (dir: string, depth: number) => {
    if (depth > maxDepth) {
      truncated = true
      return
    }
    let dirents: import('node:fs').Dirent[]
    try {
      dirents = readdirSync(dir, { withFileTypes: true })
    } catch {
      return
    }
    for (const ent of dirents) {
      if (out.length >= maxFiles) {
        truncated = true
        return
      }
      const path = join(dir, ent.name)
      if (ent.isDirectory()) {
        if (!isHiddenEntry(dir, ent.name, true)) walk(path, depth + 1)
      } else if (
        ent.isFile() &&
        isSupportedTreeFile(ent.name) &&
        !isHiddenEntry(dir, ent.name, false)
      ) {
        const st = statOrNull(path)
        if (st) out.push(st)
      }
    }
  }
  walk(root, 0)
  return { files: out, truncated }
}

export function statOrNull(path: string): ScannedFile | null {
  try {
    const st = statSync(path)
    return st.isFile() ? { path, mtimeMs: st.mtimeMs, sizeBytes: st.size } : null
  } catch {
    return null
  }
}
