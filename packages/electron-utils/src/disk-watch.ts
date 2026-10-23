import { stat, unwatchFile, watchFile, type Stats } from 'node:fs'

export interface DiskStamp {
  mtimeMs: number
  size: number
}

type StatsLike = Pick<Stats, 'mtimeMs' | 'size' | 'nlink'>

export interface DiskWatchFs {
  watchFile(path: string, interval: number, listener: (curr: StatsLike) => void): void
  unwatchFile(path: string, listener: (curr: StatsLike) => void): void
  stat(path: string): Promise<StatsLike | null>
}

export const nodeDiskWatchFs: DiskWatchFs = {
  watchFile: (path, interval, listener) =>
    watchFile(path, { persistent: false, interval }, listener as (curr: Stats) => void),
  unwatchFile: (path, listener) => unwatchFile(path, listener as (curr: Stats) => void),
  stat: (path) => new Promise((resolveStat) => stat(path, (err, s) => resolveStat(err ? null : s))),
}

/**
 * Watches the one file a document view has open and reports edits made by
 * other programs. Polling (watchFile) rather than fs.watch: editors replace
 * files by rename, which fs.watch misses on every platform.
 */
export interface DocumentDiskWatcher {
  /** follow a new path; the current stat becomes the baseline */
  watch(path: string | null): Promise<void>
  /** mute ticks while the owner writes; endOwnWrite re-baselines */
  beginOwnWrite(): void
  endOwnWrite(path: string): Promise<void>
  /** the renderer has read the changed file: it becomes the baseline a save is checked against */
  acknowledge(): Promise<void>
  /** false when the file on disk no longer matches the last baseline */
  isCurrent(): Promise<boolean>
  dispose(): void
}

const isGone = (s: StatsLike) => s.nlink === 0 || (s.mtimeMs === 0 && s.size === 0)
const sameStamp = (a: DiskStamp | null, b: StatsLike) =>
  a !== null && a.mtimeMs === b.mtimeMs && a.size === b.size

export function createDocumentDiskWatcher(
  onChanged: (path: string) => void,
  options: { intervalMs?: number; fs?: DiskWatchFs } = {},
): DocumentDiskWatcher {
  const fs = options.fs ?? nodeDiskWatchFs
  const intervalMs = options.intervalMs ?? 1000
  let path: string | null = null
  let baseline: DiskStamp | null = null
  /** the stamp last reported to the renderer; repeated polls of the same change stay quiet */
  let reported: DiskStamp | null = null
  let muted = false
  let listener: ((curr: StatsLike) => void) | null = null

  const stop = () => {
    if (path && listener) fs.unwatchFile(path, listener)
    listener = null
  }

  const onTick = (watched: string) => (curr: StatsLike) => {
    if (muted || watched !== path || isGone(curr)) return
    if (sameStamp(baseline, curr) || sameStamp(reported, curr)) return
    // the baseline stays put until the renderer acknowledges, so a save landing
    // in between is refused as a conflict instead of clobbering the file
    reported = { mtimeMs: curr.mtimeMs, size: curr.size }
    onChanged(watched)
  }

  const watcher: DocumentDiskWatcher = {
    async watch(next) {
      if (next === path) return
      stop()
      path = next
      baseline = null
      reported = null
      if (!next) return
      const s = await fs.stat(next)
      if (path !== next) return
      baseline = s && !isGone(s) ? { mtimeMs: s.mtimeMs, size: s.size } : null
      listener = onTick(next)
      fs.watchFile(next, intervalMs, listener)
    },
    beginOwnWrite() {
      muted = true
    },
    async endOwnWrite(written) {
      if (written !== path) {
        muted = false
        await watcher.watch(written)
        return
      }
      const s = await fs.stat(written)
      if (s && !isGone(s)) baseline = { mtimeMs: s.mtimeMs, size: s.size }
      reported = null
      muted = false
    },
    async acknowledge() {
      if (!path) return
      const s = await fs.stat(path)
      if (s && !isGone(s)) baseline = { mtimeMs: s.mtimeMs, size: s.size }
      reported = null
    },
    async isCurrent() {
      if (!path || !baseline) return true
      const s = await fs.stat(path)
      return !s || isGone(s) || sameStamp(baseline, s)
    },
    dispose() {
      stop()
      path = null
      baseline = null
      reported = null
    },
  }
  return watcher
}
