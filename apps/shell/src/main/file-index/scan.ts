import { readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { isHiddenEntry, isSupportedTreeFile } from '../folder-tree'

export interface ScannedFile {
  path: string
  mtimeMs: number
  sizeBytes: number
}

/** every supported, visible file under `root` with the stat fields the index keys on */
export function scanFiles(root: string): ScannedFile[] {
  const out: ScannedFile[] = []
  const walk = (dir: string) => {
    let dirents: import('node:fs').Dirent[]
    try {
      dirents = readdirSync(dir, { withFileTypes: true })
    } catch {
      return
    }
    for (const ent of dirents) {
      const path = join(dir, ent.name)
      if (ent.isDirectory()) {
        if (!isHiddenEntry(dir, ent.name, true)) walk(path)
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
  walk(root)
  return out
}

export function statOrNull(path: string): ScannedFile | null {
  try {
    const st = statSync(path)
    return st.isFile() ? { path, mtimeMs: st.mtimeMs, sizeBytes: st.size } : null
  } catch {
    return null
  }
}
