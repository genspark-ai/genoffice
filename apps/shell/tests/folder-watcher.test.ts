import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/** `watch` is the real one unless a test scripts it, so that every call can be inspected */
const real = vi.hoisted(() => ({ watch: null as unknown as typeof import('node:fs').watch }))
vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  real.watch = actual.watch
  return { ...actual, watch: vi.fn(actual.watch) }
})
/** lets a test hold back the listing of one folder, to fix the order of two walks */
const held = vi.hoisted(() => ({ before: null as null | ((dir: string) => Promise<void> | void) }))
vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  const readdir = async (dir: string, options: never): Promise<unknown> => {
    await held.before?.(dir)
    return actual.readdir(dir, options)
  }
  return { ...actual, readdir }
})
import { EventEmitter } from 'node:events'
import {
  mkdirSync,
  mkdtempSync,
  realpathSync,
  renameSync,
  rmSync,
  watch,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { FolderWatcher } from '../src/main/folder-tree'
import { PerFolderWatch } from '../src/main/per-folder-watch'

const hostPlatform = Object.getOwnPropertyDescriptor(process, 'platform')!
const usePlatform = (platform: NodeJS.Platform): void => {
  Object.defineProperty(process, 'platform', { ...hostPlatform, value: platform })
}

class ScriptedWatcher extends EventEmitter {
  closed = false
  constructor(readonly fire: (event: 'rename' | 'change', name: string | null) => void) {
    super()
  }
  close(): void {
    this.closed = true
  }
}

/** the watchers a scripted `watch` handed out, by folder; a closed one stays until its folder is watched again */
let scripted: Map<string, ScriptedWatcher>
/** makes the scripted `watch` fail for a folder, the way a used-up inotify limit does */
let refuse: (dir: string) => boolean
/** how long every scripted `watch` call takes, the way it does on a slow mount */
let slowWatchMs: number

function scriptWatch(): void {
  vi.mocked(watch).mockImplementation(((dir: string, _options: unknown, listener: unknown) => {
    for (const end = performance.now() + slowWatchMs; performance.now() < end;) {
      // busy: fs.watch is a synchronous call
    }
    if (refuse(dir)) {
      throw Object.assign(new Error('no space left on device'), { code: 'ENOSPC' })
    }
    const watcher = new ScriptedWatcher(listener as ScriptedWatcher['fire'])
    scripted.set(dir, watcher)
    return watcher
  }) as unknown as typeof watch)
}

const watchedDirs = (): string[] =>
  [...scripted]
    .filter(([, watcher]) => !watcher.closed)
    .map(([dir]) => dir)
    .sort()
const watchCalls = (): string[] => vi.mocked(watch).mock.calls.map(([dir]) => String(dir))
const eventually = (check: () => void): Promise<void> => vi.waitFor(check, { timeout: 5000 })
/** long enough for a walk or an lstat that should change nothing to have finished */
const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 150))

let root: string
/** what the watch under test reported: names relative to the root, as `fs.watch` gives them */
let seen: string[]
let open: { close(): void }[]
/** the listings of `dir` that are being held back; calling one lets it go on */
let gates: (() => void)[]
function holdListingOf(dir: string): void {
  held.before = (listed) =>
    listed === dir ? new Promise<void>((release) => gates.push(release)) : undefined
}

const at = (...parts: string[]): string => join(root, ...parts)
function folder(...parts: string[]): string {
  mkdirSync(at(...parts), { recursive: true })
  return at(...parts)
}
function file(...parts: string[]): string {
  writeFileSync(at(...parts), 'x')
  return at(...parts)
}
function perFolder(dir = root): PerFolderWatch {
  const watcher = new PerFolderWatch(dir, (filename) => {
    seen.push(filename)
  })
  open.push(watcher)
  return watcher
}

beforeEach(() => {
  root = realpathSync(mkdtempSync(join(tmpdir(), 'folder-watcher-')))
  seen = []
  open = []
  gates = []
  scripted = new Map()
  refuse = () => false
  slowWatchMs = 0
  vi.mocked(watch).mockClear()
})

afterEach(() => {
  for (const watcher of open) watcher.close()
  held.before = null
  for (const release of gates) release()
  vi.mocked(watch).mockImplementation(real.watch)
  Object.defineProperty(process, 'platform', hostPlatform)
  rmSync(root, { recursive: true, force: true })
})

describe('PerFolderWatch', () => {
  beforeEach(scriptWatch)

  it('returns once the root is watched and finds the folders below it afterwards', async () => {
    folder('a', 'b')
    folder('c')
    file('a', 'b', 'note.md')
    file('top.md')

    const watcher = perFolder()

    expect(watcher.size).toBe(1)
    await eventually(() => expect(watchedDirs()).toEqual([root, at('a'), at('a', 'b'), at('c')]))
    expect(watcher.size).toBe(4)
    for (const [, options] of vi.mocked(watch).mock.calls) {
      expect(options).toEqual({ persistent: false })
    }
  })

  it('reports a change under the name fs.watch gives it, relative to the root', async () => {
    folder('a', 'b')
    perFolder()
    await eventually(() => expect(watchedDirs()).toHaveLength(3))

    scripted.get(at('a', 'b'))!.fire('change', 'note.md')
    scripted.get(root)!.fire('change', 'top.md')
    scripted.get(at('a'))!.fire('change', null)
    scripted.get(root)!.fire('rename', null)

    expect(seen).toEqual([join('a', 'b', 'note.md'), 'top.md', 'a', ''])
  })

  it('watches a folder that appears later and reports it once its watch is in place', async () => {
    folder('first')
    const watcher = perFolder()
    await eventually(() => expect(watcher.size).toBe(2))
    await settle()
    folder('late', 'inner')
    file('late', 'inner', 'written-before-the-watch.md')

    scripted.get(root)!.fire('rename', 'late')

    await eventually(() =>
      expect(watchedDirs()).toEqual([root, at('first'), at('late'), at('late', 'inner')]),
    )
    // the event itself, then each new folder once it is watched
    expect(seen).toEqual(['late', 'late', join('late', 'inner')])
  })

  it('reports a folder that appears while the first walk is still reading it', async () => {
    folder('late', 'inner')
    holdListingOf(at('late'))
    perFolder()
    // the first walk watches `late` and waits for its listing
    await eventually(() => expect(gates).toHaveLength(1))
    const first = scripted.get(at('late'))!

    scripted.get(root)!.fire('rename', 'late')
    // the walk for the event watches `late` again and waits for the same listing
    await eventually(() => expect(gates).toHaveLength(2))
    expect(first.closed).toBe(true)
    gates[0]()
    await settle()
    gates[1]()

    await eventually(() => expect(watchedDirs()).toEqual([root, at('late'), at('late', 'inner')]))
    expect(seen).toEqual(['late', 'late', join('late', 'inner')])
    expect(watchCalls().filter((dir) => dir === at('late', 'inner'))).toHaveLength(1)
  })

  it('leaves a new file unwatched', async () => {
    const watcher = perFolder()
    file('note.md')

    scripted.get(root)!.fire('rename', 'note.md')

    await settle()
    expect(watcher.size).toBe(1)
    expect(seen).toEqual(['note.md'])
  })

  it('changes nothing for what a watched folder reports about itself', async () => {
    folder('a', 'b')
    perFolder()
    await eventually(() => expect(watchedDirs()).toHaveLength(3))

    // a removed or moved folder reports a rename of its own name
    scripted.get(at('a'))!.fire('rename', 'a')

    await settle()
    expect(watchedDirs()).toEqual([root, at('a'), at('a', 'b')])
    expect(seen).toEqual([join('a', 'a')])
  })

  it('stops watching a folder that went away, with everything below it', async () => {
    folder('a', 'b')
    folder('ab')
    perFolder()
    await eventually(() => expect(watchedDirs()).toHaveLength(4))
    const gone = [scripted.get(at('a'))!, scripted.get(at('a', 'b'))!]

    rmSync(at('a'), { recursive: true })
    scripted.get(root)!.fire('rename', 'a')

    await eventually(() => expect(watchedDirs()).toEqual([root, at('ab')]))
    expect(gone.map((watcher) => watcher.closed)).toEqual([true, true])
  })

  it('watches a renamed folder under its new name', async () => {
    folder('a', 'b')
    perFolder()
    await eventually(() => expect(watchedDirs()).toHaveLength(3))
    const old = [scripted.get(at('a'))!, scripted.get(at('a', 'b'))!]

    renameSync(at('a'), at('z'))
    scripted.get(root)!.fire('rename', 'a')
    scripted.get(root)!.fire('rename', 'z')

    await eventually(() => expect(watchedDirs()).toEqual([root, at('z'), at('z', 'b')]))
    expect(old.map((watcher) => watcher.closed)).toEqual([true, true])
  })

  it('watches the folder that took the place of a removed one', async () => {
    folder('a', 'old')
    perFolder()
    await eventually(() => expect(watchedDirs()).toHaveLength(3))
    const before = scripted.get(at('a'))!

    rmSync(at('a'), { recursive: true })
    folder('a', 'new')
    scripted.get(root)!.fire('rename', 'a')

    await eventually(() => expect(watchedDirs()).toEqual([root, at('a'), at('a', 'new')]))
    expect(before.closed).toBe(true)
    expect(scripted.get(at('a'))).not.toBe(before)
  })

  it('stops asking once the system refuses a watch, and keeps what it has', async () => {
    folder('a')
    folder('b')
    folder('c')
    refuse = () => scripted.size >= 2

    const watcher = perFolder()

    await eventually(() => expect(watchCalls()).toHaveLength(3))
    await settle()
    expect(watchCalls()).toHaveLength(3)
    expect(watcher.size).toBe(2)
  })

  it('has the folders nearest the root watched when the system refuses more', async () => {
    folder('a', 'a1', 'a2')
    folder('b', 'b1', 'b2')
    refuse = () => scripted.size >= 5

    const watcher = perFolder()

    await eventually(() => expect(watchCalls()).toHaveLength(6))
    await settle()
    expect(watchedDirs()).toEqual([root, at('a'), at('a', 'a1'), at('b'), at('b', 'b1')])
    expect(watcher.size).toBe(5)
  })

  it('asks again after a watched folder went away', async () => {
    folder('a')
    refuse = (dir) => dir === at('late')
    const watcher = perFolder()
    await eventually(() => expect(watcher.size).toBe(2))
    folder('late')
    scripted.get(root)!.fire('rename', 'late')
    await eventually(() => expect(watchCalls()).toContain(at('late')))
    await settle()
    refuse = () => false

    folder('refused-while-full')
    scripted.get(root)!.fire('rename', 'refused-while-full')
    await settle()
    expect(watchCalls()).not.toContain(at('refused-while-full'))

    rmSync(at('a'), { recursive: true })
    scripted.get(root)!.fire('rename', 'a')
    await eventually(() => expect(watcher.size).toBe(1))
    folder('after')
    scripted.get(root)!.fire('rename', 'after')

    await eventually(() => expect(watchedDirs()).toEqual([root, at('after')]))
  })

  it('lets the event loop run between the subfolders of one folder when watching is slow', async () => {
    const subfolders = 40
    for (let i = 0; i < subfolders; i += 1) folder(`sub-${i}`)
    slowWatchMs = 3
    const watcher = perFolder()
    /** watch calls made between two turns of the event loop */
    const perTurn: number[] = []
    let counted = watchCalls().length
    let counting = true
    const turn = (): void => {
      perTurn.push(watchCalls().length - counted)
      counted = watchCalls().length
      if (counting) setImmediate(turn)
    }
    setImmediate(turn)

    await eventually(() => expect(watcher.size).toBe(subfolders + 1))
    counting = false

    // 10 ms of 3 ms calls is four or five of them, never the whole folder
    expect(Math.max(...perTurn)).toBeLessThanOrEqual(8)
  })

  it('throws when the root cannot be watched', () => {
    refuse = () => true

    expect(() => perFolder()).toThrow('no space left on device')
  })

  it('emits error when the watch on the root fails, and drops any other failing folder', async () => {
    folder('a', 'b')
    folder('c')
    const watcher = perFolder()
    const errors: unknown[] = []
    watcher.on('error', (err) => errors.push(err))
    await eventually(() => expect(watchedDirs()).toHaveLength(4))

    scripted.get(at('a'))!.emit('error', new Error('a failed'))

    expect(watchedDirs()).toEqual([root, at('c')])
    expect(errors).toEqual([])

    const failure = new Error('root failed')
    scripted.get(root)!.emit('error', failure)

    expect(errors).toEqual([failure])
  })

  it('stops a walk in progress and reports nothing after close', async () => {
    folder('a', 'b')
    folder('c')
    const watcher = perFolder()
    const rootWatcher = scripted.get(root)!

    watcher.close()
    await settle()

    expect(watcher.size).toBe(0)
    expect(watchedDirs()).toEqual([])
    rootWatcher.fire('rename', 'a')
    await settle()
    expect(seen).toEqual([])
    expect(watchCalls()).toEqual([root])
  })

  it('stops a walk that is already below the root when closed', async () => {
    folder('a', 'b', 'c')
    holdListingOf(at('a', 'b'))
    const watcher = perFolder()
    await eventually(() => expect(gates).toHaveLength(1))

    watcher.close()
    gates[0]()
    await settle()

    expect(watchedDirs()).toEqual([])
    expect(watchCalls()).toEqual([root, at('a'), at('a', 'b')])
  })

  it('watches nothing for an event that is still being looked at when closed', async () => {
    const watcher = perFolder()
    await settle()
    folder('late')

    scripted.get(root)!.fire('rename', 'late')
    watcher.close()
    await settle()

    expect(watchedDirs()).toEqual([])
    expect(watchCalls()).toEqual([root])
  })
})

// what inotify reports for a real folder; the other platforms never take this path
describe.skipIf(hostPlatform.value !== 'linux')('PerFolderWatch on a real folder', () => {
  it('reports a file written into a nested folder', async () => {
    folder('a', 'b')
    const watcher = perFolder()
    await eventually(() => expect(watcher.size).toBe(3))

    file('a', 'b', 'note.md')

    await eventually(() => expect(seen).toContain(join('a', 'b', 'note.md')))
  })

  it('follows a folder created after the start', async () => {
    folder('first')
    const watcher = perFolder()
    await eventually(() => expect(watcher.size).toBe(2))
    await settle()

    folder('late', 'inner')
    await eventually(() => expect(watcher.size).toBe(4))
    file('late', 'inner', 'note.md')

    await eventually(() => expect(seen).toContain(join('late', 'inner', 'note.md')))
  })

  it('stops watching a removed folder', async () => {
    folder('a', 'b')
    const watcher = perFolder()
    await eventually(() => expect(watcher.size).toBe(3))

    rmSync(at('a'), { recursive: true })

    await eventually(() => expect(watcher.size).toBe(1))
  })

  it('follows a renamed folder', async () => {
    folder('a', 'b')
    const watcher = perFolder()
    await eventually(() => expect(watcher.size).toBe(3))

    renameSync(at('a'), at('z'))
    await eventually(() => {
      expect(seen).toContain(join('z', 'b'))
      expect(watcher.size).toBe(3)
    })
    file('z', 'b', 'note.md')

    await eventually(() => expect(seen).toContain(join('z', 'b', 'note.md')))
    expect(seen).not.toContain(join('a', 'b', 'note.md'))
  })

  it('drops the folders below a removed root and raises no error', async () => {
    folder('a', 'b')
    const watcher = perFolder()
    const errors: unknown[] = []
    watcher.on('error', (err) => errors.push(err))
    await eventually(() => expect(watcher.size).toBe(3))

    rmSync(root, { recursive: true })

    await eventually(() => expect(watcher.size).toBe(1))
    expect(errors).toEqual([])
  })

  it('throws when the root does not exist', () => {
    expect(() => perFolder(at('missing'))).toThrow()
  })
})

describe('FolderWatcher', () => {
  let changes: string[][]
  function folderWatcher(dir = root): FolderWatcher {
    const watcher = new FolderWatcher(dir, (dirs) => changes.push(dirs))
    open.push(watcher)
    return watcher
  }

  beforeEach(() => {
    changes = []
  })

  it.each(['darwin', 'win32'] as const)('keeps the one recursive watch on %s', (platform) => {
    usePlatform(platform)
    scriptWatch()
    folder('a', 'b')

    const watcher = folderWatcher()

    expect(watcher.active).toBe(true)
    expect(vi.mocked(watch).mock.calls).toEqual([
      [root, { recursive: true, persistent: false }, expect.any(Function)],
    ])
  })

  it('watches folder by folder on Linux', async () => {
    usePlatform('linux')
    scriptWatch()
    folder('a', 'b')
    folder('.git', 'objects')

    const watcher = folderWatcher()

    expect(watcher.active).toBe(true)
    await eventually(() =>
      expect(watchedDirs()).toEqual([
        root,
        at('.git'),
        at('.git', 'objects'),
        at('a'),
        at('a', 'b'),
      ]),
    )
    for (const [, options] of vi.mocked(watch).mock.calls) {
      expect(options).toEqual({ persistent: false })
    }
  })

  it('reports the changed folders on Linux in one debounced call', async () => {
    usePlatform('linux')
    scriptWatch()
    folder('a', 'b')
    folderWatcher()
    await eventually(() => expect(watchedDirs()).toHaveLength(3))

    scripted.get(at('a', 'b'))!.fire('change', 'note.md')
    scripted.get(at('a', 'b'))!.fire('change', 'note.md')
    scripted.get(at('a'))!.fire('rename', 'gone.md')

    await eventually(() => expect(changes).toHaveLength(1))
    expect(changes[0]).toEqual([at('a', 'b'), at('a')])
  })

  it('goes inactive on Linux when the watch on the root fails, and closes every folder', async () => {
    usePlatform('linux')
    scriptWatch()
    folder('a')
    const watcher = folderWatcher()
    await eventually(() => expect(watchedDirs()).toHaveLength(2))

    scripted.get(root)!.emit('error', new Error('root failed'))

    expect(watcher.active).toBe(false)
    expect(watchedDirs()).toEqual([])
  })

  it.each(['linux', 'darwin'] as const)(
    'is inactive on %s when the root does not exist',
    (platform) => {
      usePlatform(platform)

      expect(folderWatcher(at('missing')).active).toBe(false)
    },
  )

  it.skipIf(hostPlatform.value !== 'linux')(
    'reports a file written below the root on a Linux host',
    async () => {
      folder('a', 'b')
      folderWatcher()
      await eventually(() => expect(watchCalls()).toHaveLength(3))

      file('a', 'b', 'note.md')

      await eventually(() => expect(changes.flat()).toContain(at('a', 'b')))
    },
  )
})
