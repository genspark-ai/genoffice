import {
  accessSync,
  constants,
  cpSync,
  existsSync,
  mkdirSync,
  readdirSync,
  realpathSync,
  renameSync,
  rmSync,
  statSync,
  watch,
  type FSWatcher,
} from 'node:fs'
import { basename, dirname, extname, join, resolve, sep } from 'node:path'
import type {
  FolderEntry,
  FolderListing,
  FolderRoot,
  MoveConflictPolicy,
  MoveResult,
  RecentEntry,
  RenameResult,
} from '../shared/home-api'
import { isValidRenameName } from './rename-validation'
import { statPathEntries } from './recent-files'

/** files the home tree shows; mirrors the shell's open-dialog filter */
export const TREE_FILE_EXTENSIONS: ReadonlySet<string> = new Set([
  'docx',
  'doc',
  'xlsx',
  'xlsm',
  'xls',
  'csv',
  'tsv',
  'pptx',
  'ppt',
  'pdf',
  'md',
  'markdown',
  'html',
  'htm',
])

const HIDDEN_DIR_NAMES = new Set(['node_modules', '__macosx'])
const HIDDEN_FILE_NAMES = new Set(['thumbs.db', 'desktop.ini'])
/** the Markdown / HTML apps' image folder is app-owned when it holds this manifest */
const MD_ASSET_MANIFEST = '.genoffice-assets.json'

export function isSupportedTreeFile(name: string): boolean {
  return TREE_FILE_EXTENSIONS.has(extname(name).slice(1).toLowerCase())
}

/** dotfiles, Office lock files, OS junk and the Markdown `assets/` folder stay out of the tree */
export function isHiddenEntry(dir: string, name: string, isDir: boolean): boolean {
  if (name.startsWith('.') || name.startsWith('~$')) return true
  const lower = name.toLowerCase()
  if (isDir) {
    if (HIDDEN_DIR_NAMES.has(lower)) return true
    if (lower === 'assets' && existsSync(join(dir, name, MD_ASSET_MANIFEST))) return true
    return false
  }
  return HIDDEN_FILE_NAMES.has(lower)
}

function realOrResolved(path: string): string {
  try {
    return realpathSync.native(path)
  } catch {
    return resolve(path)
  }
}

/**
 * True when `path` is the root or lives under it. Existing paths compare by
 * real path (so a symlink pointing outside is rejected); a not-yet-existing
 * target compares by its nearest existing ancestor.
 */
export function isInsideRoot(root: string, path: string): boolean {
  const realRoot = realOrResolved(root)
  let probe = resolve(path)
  const missing: string[] = []
  while (!existsSync(probe)) {
    const parent = dirname(probe)
    if (parent === probe) return false
    missing.unshift(basename(probe))
    probe = parent
  }
  const real = join(realOrResolved(probe), ...missing)
  return real === realRoot || real.startsWith(realRoot + sep)
}

export function describeRoot(root: string): FolderRoot {
  let usable: boolean
  try {
    mkdirSync(root, { recursive: true })
    accessSync(root, constants.W_OK)
    usable = statSync(root).isDirectory()
  } catch {
    usable = false
  }
  return { path: root, name: basename(root) || root, usable, readable: usable, removable: false }
}

function hasVisibleSubfolder(dir: string): boolean {
  try {
    return readdirSync(dir, { withFileTypes: true }).some(
      (d) => d.isDirectory() && !isHiddenEntry(dir, d.name, true),
    )
  } catch {
    return false
  }
}

export function listFolder(dir: string, starredPaths: ReadonlySet<string>): FolderListing {
  const folders: FolderEntry[] = []
  const filePaths: string[] = []
  let dirents: import('node:fs').Dirent[]
  try {
    dirents = readdirSync(dir, { withFileTypes: true })
  } catch {
    return { dir, folders: [], files: [], missing: true }
  }
  for (const d of dirents) {
    const path = join(dir, d.name)
    if (d.isDirectory()) {
      if (isHiddenEntry(dir, d.name, true)) continue
      let mtimeMs: number
      try {
        mtimeMs = statSync(path).mtimeMs
      } catch {
        continue
      }
      folders.push({ path, name: d.name, mtimeMs, hasSubfolders: hasVisibleSubfolder(path) })
    } else if (d.isFile() && !isHiddenEntry(dir, d.name, false) && isSupportedTreeFile(d.name)) {
      filePaths.push(path)
    }
  }
  folders.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }))
  const files: RecentEntry[] = statPathEntries(filePaths, starredPaths)
    .filter((e) => !e.missing)
    .sort((a, b) => b.mtimeMs - a.mtimeMs)
  return { dir, folders, files }
}

/** the candidates below `dir` at any depth; bookkeeping filters tracked paths instead of walking the disk */
export function pathsUnder(dir: string, candidates: Iterable<string>): string[] {
  const prefix = resolve(dir) + sep
  const out = new Set<string>()
  for (const path of candidates) if (resolve(path).startsWith(prefix)) out.add(path)
  return [...out]
}

export interface FolderErrors {
  badArgs: string
  badName: string
  missing: string
  exists: string
  failed: string
}

export function createFolder(parent: string, name: string, errors: FolderErrors): RenameResult {
  const trimmed = name.trim()
  if (!isValidRenameName(trimmed)) return { ok: false, error: errors.badName }
  if (!existsSync(parent)) return { ok: false, error: errors.missing }
  const target = join(parent, trimmed)
  if (existsSync(target)) return { ok: false, error: errors.exists }
  try {
    mkdirSync(target)
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : errors.failed }
  }
  return { ok: true, path: target }
}

/** `name.ext` → `name (2).ext`, `name (3).ext`… — first free name inside dir */
export function uniqueNameIn(dir: string, name: string): string {
  if (!existsSync(join(dir, name))) return name
  const ext = extname(name)
  const base = name.slice(0, name.length - ext.length)
  for (let i = 2; ; i++) {
    const candidate = `${base} (${i})${ext}`
    if (!existsSync(join(dir, candidate))) return candidate
  }
}

/** rename, falling back to copy + delete when the target is on another volume */
export function moveOnDisk(from: string, to: string): void {
  try {
    renameSync(from, to)
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code !== 'EXDEV') throw err
    cpSync(from, to, { recursive: true, errorOnExist: true, force: false })
    rmSync(from, { recursive: true, force: true })
  }
}

/** the displaced target of a 'replace' move: finalized once the move succeeded, put back if it failed */
export interface ReplacedItem {
  commit: () => void
  rollback: () => void
}

export interface MoveOptions {
  /** park an existing target out of the way before the move (policy 'replace') */
  replaceExisting: (path: string) => ReplacedItem
}

/** True when `dir` is `path` itself or somewhere below it (a folder cannot move into itself). */
export function isSelfOrDescendant(path: string, dir: string): boolean {
  const a = resolve(path)
  const b = resolve(dir)
  return a === b || b.startsWith(a + sep)
}

/**
 * Move each path into `targetDir`. Returns what moved (old → new) so the
 * caller can re-key recents, stars, chat history and open tabs. With policy
 * 'ask' nothing that collides moves; the collisions come back for the UI to
 * ask about and retry with an explicit policy.
 */
export function movePathsInto(
  paths: readonly string[],
  targetDir: string,
  policy: MoveConflictPolicy,
  errors: FolderErrors,
  options: MoveOptions,
): MoveResult {
  const result: MoveResult = { moved: [], conflicts: [], failed: [] }
  if (!existsSync(targetDir)) {
    for (const p of paths) result.failed.push({ path: p, error: errors.missing })
    return result
  }
  for (const from of paths) {
    if (!existsSync(from)) {
      result.failed.push({ path: from, error: errors.missing })
      continue
    }
    if (dirname(resolve(from)) === resolve(targetDir)) continue
    let isDir: boolean
    try {
      isDir = statSync(from).isDirectory()
    } catch {
      result.failed.push({ path: from, error: errors.missing })
      continue
    }
    if (isDir && isSelfOrDescendant(from, targetDir)) {
      result.failed.push({ path: from, error: errors.badArgs })
      continue
    }
    const name = basename(from)
    let to = join(targetDir, name)
    let replaced: ReplacedItem | undefined
    if (existsSync(to)) {
      if (policy === 'ask' || policy === 'skip') {
        result.conflicts.push(from)
        continue
      }
      if (policy === 'keepBoth') {
        to = join(targetDir, uniqueNameIn(targetDir, name))
      } else {
        try {
          replaced = options.replaceExisting(to)
        } catch (err) {
          result.failed.push({
            path: from,
            error: err instanceof Error ? err.message : errors.failed,
          })
          continue
        }
      }
    }
    try {
      moveOnDisk(from, to)
      result.moved.push({ from, to })
      replaced?.commit()
    } catch (err) {
      // a failed move must leave the destination as it was
      replaced?.rollback()
      result.failed.push({ path: from, error: err instanceof Error ? err.message : errors.failed })
    }
  }
  return result
}

export function renameFolder(dir: string, newName: string, errors: FolderErrors): RenameResult {
  const name = newName.trim()
  if (!isValidRenameName(name)) return { ok: false, error: errors.badName }
  if (!existsSync(dir)) return { ok: false, error: errors.missing }
  const target = join(dirname(dir), name)
  if (target === dir) return { ok: true, path: dir }
  if (existsSync(target) && !isSameEntry(dir, target)) return { ok: false, error: errors.exists }
  try {
    renameSync(dir, target)
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : errors.failed }
  }
  return { ok: true, path: target }
}

function isSameEntry(a: string, b: string): boolean {
  try {
    const sa = statSync(a)
    const sb = statSync(b)
    return sa.dev === sb.dev && sa.ino === sb.ino
  } catch {
    return false
  }
}

/** `oldPrefix/…/file` → `newPrefix/…/file` for a file that lived under a renamed/moved folder */
export function rebasePath(path: string, oldDir: string, newDir: string): string {
  const rel = resolve(path).slice(resolve(oldDir).length)
  return join(newDir, rel)
}

const WATCH_DEBOUNCE_MS = 250

/**
 * Watches the root recursively and reports the folders whose direct
 * contents changed, debounced. Changes from Finder / Explorer / other apps
 * show up in the tree the same way as our own edits.
 */
export class FolderWatcher {
  private watcher: FSWatcher | null = null
  private pending = new Set<string>()
  private timer: NodeJS.Timeout | null = null

  /** false when the watch handle could not be created or has since errored */
  get active(): boolean {
    return this.watcher !== null
  }

  constructor(
    private readonly root: string,
    private readonly onChange: (dirs: string[]) => void,
  ) {
    try {
      this.watcher = watch(root, { recursive: true, persistent: false }, (_event, filename) => {
        this.note(filename)
      })
      this.watcher.on('error', () => this.close())
    } catch {
      this.watcher = null
    }
  }

  private note(filename: string | Buffer | null): void {
    const rel = filename == null ? '' : filename.toString()
    const changed = rel ? join(this.root, rel) : this.root
    // a created/removed folder changes its parent's listing; a file inside
    // changes its own folder's listing — report both candidates
    this.pending.add(dirname(changed))
    let isDir: boolean
    try {
      isDir = statSync(changed).isDirectory()
    } catch {
      isDir = false
    }
    if (isDir) this.pending.add(changed)
    if (this.timer) clearTimeout(this.timer)
    this.timer = setTimeout(() => {
      this.timer = null
      const dirs = [...this.pending]
      this.pending.clear()
      this.onChange(dirs)
    }, WATCH_DEBOUNCE_MS)
  }

  close(): void {
    if (this.timer) clearTimeout(this.timer)
    this.timer = null
    this.watcher?.close()
    this.watcher = null
  }
}
