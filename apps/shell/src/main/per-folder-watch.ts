import { EventEmitter } from 'node:events'
import { watch, type Dirent, type FSWatcher } from 'node:fs'
import { lstat, readdir } from 'node:fs/promises'
import { dirname, join, relative, resolve, sep } from 'node:path'

/** how long the walk places watches before it lets the event loop run again */
const WALK_SLICE_MS = 10

/**
 * Watches a tree on platforms without a recursive watch of their own (Linux),
 * in place of `fs.watch(root, { recursive: true })`. There Node walks the
 * whole tree synchronously inside that call, and the main process answers
 * nothing until the walk is over: seconds on a network mount with a few
 * hundred folders.
 *
 * This watches the root at once and finds the folders below it with an
 * asynchronous walk, nearest first. The listener gets what `fs.watch` gives
 * its own: the name of the entry that changed, relative to the root. `error`
 * is emitted when the watch on the root itself fails.
 */
export class PerFolderWatch extends EventEmitter {
  private readonly root: string
  private readonly watchers = new Map<string, FSWatcher>()
  private closed = false
  /** the system refused another watch, so asking for more is pointless */
  private exhausted = false

  constructor(
    root: string,
    private readonly listener: (filename: string) => void,
  ) {
    super()
    this.root = resolve(root)
    const watcher = this.open(this.root)
    this.watchers.set(this.root, watcher)
    void this.walk(this.root, watcher, false)
  }

  /** how many folders are watched right now */
  get size(): number {
    return this.watchers.size
  }

  close(): void {
    this.closed = true
    for (const watcher of this.watchers.values()) watcher.close()
    this.watchers.clear()
  }

  private report(path: string): void {
    this.listener(relative(this.root, path))
  }

  private open(dir: string): FSWatcher {
    const watcher = watch(dir, { persistent: false }, (event, name) => {
      if (this.closed) return
      const path = name == null ? dir : join(dir, name)
      this.report(path)
      // an entry appeared, went away or changed places
      if (event === 'rename' && name != null) void this.refresh(path)
    })
    watcher.on('error', (err) => {
      if (dir === this.root) this.emit('error', err)
      else this.drop(dir)
    })
    return watcher
  }

  private add(dir: string): FSWatcher | null {
    if (this.exhausted) return null
    try {
      const watcher = this.open(dir)
      this.watchers.set(dir, watcher)
      return watcher
    } catch (err) {
      // out of watches for this user or process: the next folder would fail the same way
      const code = (err as NodeJS.ErrnoException).code
      if (code === 'ENOSPC' || code === 'EMFILE') this.exhausted = true
      return null
    }
  }

  /** Stops watching `path` and everything below it. */
  private drop(path: string): void {
    if (!this.watchers.has(path)) return
    const below = path + sep
    for (const [dir, watcher] of this.watchers) {
      if (dir !== path && !dir.startsWith(below)) continue
      watcher.close()
      this.watchers.delete(dir)
    }
    this.exhausted = false
  }

  /**
   * Watches the folders below `start`, which `watcher` watches already.
   * `announce` reports each one once its watch is in place, for a folder that
   * appeared after the start: what was put into it before that raised no event.
   *
   * A folder is read only while the watch this walk knows is still the one on
   * it. A folder that was removed, or replaced and watched again in the
   * meantime, belongs to whatever watched it again.
   */
  private async walk(start: string, watcher: FSWatcher, announce: boolean): Promise<void> {
    const queue: [string, FSWatcher][] = [[start, watcher]]
    for (let next = queue.shift(); next !== undefined; next = queue.shift()) {
      const [dir, mine] = next
      if (this.closed || this.exhausted) return
      if (this.watchers.get(dir) !== mine) continue
      let entries: Dirent[]
      try {
        entries = await readdir(dir, { withFileTypes: true })
      } catch {
        // unreadable, or removed since it was listed
        continue
      }
      let sliceEnd = performance.now() + WALK_SLICE_MS
      for (const entry of entries) {
        if (!entry.isDirectory()) continue
        // fs.watch is synchronous: a folder with thousands of subfolders must not hold the event loop
        if (performance.now() > sliceEnd) {
          await new Promise((resume) => setImmediate(resume))
          sliceEnd = performance.now() + WALK_SLICE_MS
        }
        if (this.exhausted || this.watchers.get(dir) !== mine) break
        const path = join(dir, entry.name)
        if (this.watchers.has(path)) continue
        const added = this.add(path)
        if (!added) continue
        if (announce) this.report(path)
        queue.push([path, added])
      }
    }
  }

  /** Brings the watches at and below `path` in line with what is there now. */
  private async refresh(path: string): Promise<void> {
    let isFolder = false
    try {
      isFolder = (await lstat(path)).isDirectory()
    } catch {
      // gone
    }
    if (this.closed) return
    // the folder watched under this name is gone, or another one took its place
    this.drop(path)
    if (!isFolder || !this.watchers.has(dirname(path))) return
    const watcher = this.add(path)
    if (!watcher) return
    this.report(path)
    await this.walk(path, watcher, true)
  }
}
