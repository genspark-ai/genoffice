import { describe, expect, it } from 'vitest'
import { createDocumentDiskWatcher, type DiskWatchFs } from '../src/disk-watch'

type S = { mtimeMs: number; size: number; nlink: number }

function fakeFs() {
  const files = new Map<string, S>()
  const listeners = new Map<string, Set<(curr: S) => void>>()
  const fs: DiskWatchFs = {
    watchFile: (path, _interval, listener) => {
      if (!listeners.has(path)) listeners.set(path, new Set())
      listeners.get(path)!.add(listener)
    },
    unwatchFile: (path, listener) => void listeners.get(path)?.delete(listener),
    stat: async (path) => files.get(path) ?? null,
  }
  return {
    fs,
    set(path: string, s: Partial<S>) {
      files.set(path, { nlink: 1, mtimeMs: 1, size: 1, ...files.get(path), ...s })
    },
    remove(path: string) {
      files.delete(path)
    },
    tick(path: string) {
      const curr = files.get(path) ?? { mtimeMs: 0, size: 0, nlink: 0 }
      for (const l of listeners.get(path) ?? []) l(curr)
    },
    watcherCount: (path: string) => listeners.get(path)?.size ?? 0,
  }
}

describe('createDocumentDiskWatcher', () => {
  it('reports an external edit once and re-baselines', async () => {
    const disk = fakeFs()
    disk.set('/a.md', { mtimeMs: 10, size: 5 })
    const seen: string[] = []
    const w = createDocumentDiskWatcher((p) => seen.push(p), { fs: disk.fs })
    await w.watch('/a.md')
    disk.tick('/a.md')
    expect(seen).toEqual([])
    disk.set('/a.md', { mtimeMs: 20 })
    disk.tick('/a.md')
    disk.tick('/a.md')
    expect(seen).toEqual(['/a.md'])
    // a save before the renderer has looked at the change must be refused
    expect(await w.isCurrent()).toBe(false)
    await w.acknowledge()
    expect(await w.isCurrent()).toBe(true)
    disk.tick('/a.md')
    expect(seen).toEqual(['/a.md'])
  })

  it('unwatches with the listener it registered', async () => {
    const disk = fakeFs()
    disk.set('/a.md', { mtimeMs: 10, size: 5 })
    const w = createDocumentDiskWatcher(() => {}, { fs: disk.fs })
    await w.watch('/a.md')
    expect(disk.watcherCount('/a.md')).toBe(1)
    w.dispose()
    expect(disk.watcherCount('/a.md')).toBe(0)
  })

  it('stays quiet across its own write, including the tick that lands mid-write', async () => {
    const disk = fakeFs()
    disk.set('/a.md', { mtimeMs: 10, size: 5 })
    const seen: string[] = []
    const w = createDocumentDiskWatcher((p) => seen.push(p), { fs: disk.fs })
    await w.watch('/a.md')
    w.beginOwnWrite()
    disk.set('/a.md', { mtimeMs: 30, size: 9 })
    disk.tick('/a.md')
    await w.endOwnWrite('/a.md')
    disk.tick('/a.md')
    expect(seen).toEqual([])
    expect(await w.isCurrent()).toBe(true)
  })

  it('isCurrent catches a change the poll has not seen yet', async () => {
    const disk = fakeFs()
    disk.set('/a.md', { mtimeMs: 10, size: 5 })
    const w = createDocumentDiskWatcher(() => {}, { fs: disk.fs })
    await w.watch('/a.md')
    disk.set('/a.md', { size: 6 })
    expect(await w.isCurrent()).toBe(false)
  })

  it('ignores a deleted file and fires when it is recreated with new content', async () => {
    const disk = fakeFs()
    disk.set('/a.md', { mtimeMs: 10, size: 5 })
    const seen: string[] = []
    const w = createDocumentDiskWatcher((p) => seen.push(p), { fs: disk.fs })
    await w.watch('/a.md')
    disk.remove('/a.md')
    disk.tick('/a.md')
    expect(seen).toEqual([])
    expect(await w.isCurrent()).toBe(true)
    disk.set('/a.md', { mtimeMs: 40, size: 5 })
    disk.tick('/a.md')
    expect(seen).toEqual(['/a.md'])
  })

  it('follows Save As onto the new path and drops the old listener', async () => {
    const disk = fakeFs()
    disk.set('/a.md', { mtimeMs: 10, size: 5 })
    disk.set('/b.md', { mtimeMs: 10, size: 5 })
    const seen: string[] = []
    const w = createDocumentDiskWatcher((p) => seen.push(p), { fs: disk.fs })
    await w.watch('/a.md')
    w.beginOwnWrite()
    await w.endOwnWrite('/b.md')
    expect(disk.watcherCount('/a.md')).toBe(0)
    expect(disk.watcherCount('/b.md')).toBe(1)
    disk.set('/a.md', { mtimeMs: 99 })
    disk.tick('/a.md')
    disk.set('/b.md', { mtimeMs: 99 })
    disk.tick('/b.md')
    expect(seen).toEqual(['/b.md'])
    w.dispose()
    expect(disk.watcherCount('/b.md')).toBe(0)
  })
})
