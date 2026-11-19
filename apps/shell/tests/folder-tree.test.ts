import { createHash } from 'node:crypto'
import {
  chmodSync,
  constants,
  copyFileSync,
  cpSync,
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  openSync,
  readdirSync,
  readFileSync,
  realpathSync,
  renameSync,
  rmSync,
  symlinkSync,
  writeFileSync,
  type PathLike,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, parse, relative } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  createFolder,
  describeRoot,
  isHiddenEntry,
  isInsideRoot,
  isSelfOrDescendant,
  listFolder,
  moveOnDisk,
  movePathsInto,
  pathsUnder,
  rebasePath,
  renameFolder,
  uniqueNameIn,
} from '../src/main/folder-tree'

vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    chmodSync: vi.fn(actual.chmodSync),
    copyFileSync: vi.fn(actual.copyFileSync),
    cpSync: vi.fn(actual.cpSync),
    lstatSync: vi.fn(actual.lstatSync),
    openSync: vi.fn(actual.openSync),
    readdirSync: vi.fn(actual.readdirSync),
    renameSync: vi.fn(actual.renameSync),
    rmSync: vi.fn(actual.rmSync),
  }
})

const actualFs = await vi.importActual<typeof import('node:fs')>('node:fs')

const errors = {
  badArgs: 'badArgs',
  badName: 'badName',
  missing: 'missing',
  exists: 'exists',
  failed: 'failed',
}

let root: string

function touch(rel: string, content = 'x'): string {
  const path = join(root, rel)
  mkdirSync(join(path, '..'), { recursive: true })
  writeFileSync(path, content)
  return path
}

beforeEach(() => {
  root = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-folders-')))
})

afterEach(() => {
  rmSync(root, { recursive: true, force: true })
})

describe('listFolder', () => {
  it('lists sub-folders first (natural order) and supported files newest first, skipping junk', () => {
    mkdirSync(join(root, 'Clients'))
    mkdirSync(join(root, 'archive 10'))
    mkdirSync(join(root, 'archive 2'))
    mkdirSync(join(root, '.hidden'))
    mkdirSync(join(root, 'node_modules'))
    touch('report.docx')
    touch('notes.md')
    touch('image.png')
    touch('~$report.docx')
    touch('.DS_Store')
    touch('Thumbs.db')

    const listing = listFolder(root, new Set([join(root, 'notes.md')]))
    expect(listing.folders.map((f) => f.name)).toEqual(['archive 2', 'archive 10', 'Clients'])
    expect(listing.files.map((f) => f.name).sort()).toEqual(['notes.md', 'report.docx'])
    expect(listing.files.find((f) => f.name === 'notes.md')?.starred).toBe(true)
    expect(listing.folders.every((f) => f.hasSubfolders === false)).toBe(true)
  })

  it('hides the Markdown assets folder only when it holds the asset manifest', () => {
    mkdirSync(join(root, 'assets'))
    expect(isHiddenEntry(root, 'assets', true)).toBe(false)
    touch('assets/.genoffice-assets.json', '{}')
    expect(isHiddenEntry(root, 'assets', true)).toBe(true)
    expect(listFolder(root, new Set()).folders).toEqual([])
  })

  it('reports whether a folder has visible sub-folders', () => {
    mkdirSync(join(root, 'a', 'b'), { recursive: true })
    mkdirSync(join(root, 'c', '.git'), { recursive: true })
    const listing = listFolder(root, new Set())
    expect(listing.folders.find((f) => f.name === 'a')?.hasSubfolders).toBe(true)
    expect(listing.folders.find((f) => f.name === 'c')?.hasSubfolders).toBe(false)
  })

  it('returns an empty listing for a directory that vanished', () => {
    const gone = join(root, 'gone')
    expect(listFolder(gone, new Set())).toEqual({
      dir: gone,
      folders: [],
      files: [],
      missing: true,
    })
  })

  /**
   * The Home tree only lists what the shell can open, and TREE_FILE_EXTENSIONS
   * is kept to mirror the open-dialog filter. A .txt left out of it is invisible
   * on Home, so a file the open dialog happily offers cannot be clicked into
   * from the folder pane — the same drift, one list further along.
   */
  it('lists the text app extensions alongside markdown', () => {
    touch('notes.md')
    touch('notes.txt')
    touch('data.json')
    touch('photo.png')

    expect(
      listFolder(root, new Set())
        .files.map((f) => f.name)
        .sort(),
    ).toEqual(['data.json', 'notes.md', 'notes.txt'])
  })

  /**
   * The ODF imports come in a document / template pair per app. Half a pair
   * in TREE_FILE_EXTENSIONS means the open dialog offers a file the folder
   * pane refuses to show, for one app only.
   */
  it('lists both halves of every ODF import pair', () => {
    for (const extension of ['odt', 'ott', 'ods', 'ots', 'odp', 'otp']) {
      touch(`doc.${extension}`)
    }

    expect(
      listFolder(root, new Set())
        .files.map((f) => f.name)
        .sort(),
    ).toEqual(['doc.odp', 'doc.ods', 'doc.odt', 'doc.otp', 'doc.ots', 'doc.ott'])
  })
})

describe('isInsideRoot', () => {
  it('accepts the root, descendants and not-yet-existing children; rejects siblings and parents', () => {
    mkdirSync(join(root, 'sub'))
    expect(isInsideRoot(root, root)).toBe(true)
    expect(isInsideRoot(root, join(root, 'sub'))).toBe(true)
    expect(isInsideRoot(root, join(root, 'sub', 'new', 'deeper'))).toBe(true)
    expect(isInsideRoot(root, join(root, '..'))).toBe(false)
    expect(isInsideRoot(root, `${root}-sibling`)).toBe(false)
  })

  it('rejects a symlink inside the root that points outside it', () => {
    const outside = mkdtempSync(join(tmpdir(), 'genoffice-outside-'))
    try {
      symlinkSync(outside, join(root, 'escape'))
      expect(isInsideRoot(root, join(root, 'escape'))).toBe(false)
      expect(isInsideRoot(root, join(root, 'escape', 'child.docx'))).toBe(false)
    } finally {
      rmSync(outside, { recursive: true, force: true })
    }
  })
})

describe('createFolder / renameFolder', () => {
  it('creates, refuses bad names and duplicates', () => {
    expect(createFolder(root, 'Clients', errors)).toEqual({ ok: true, path: join(root, 'Clients') })
    expect(createFolder(root, 'Clients', errors)).toEqual({ ok: false, error: 'exists' })
    expect(createFolder(root, 'bad/name', errors)).toEqual({ ok: false, error: 'badName' })
    expect(createFolder(root, '   ', errors)).toEqual({ ok: false, error: 'badName' })
    expect(createFolder(join(root, 'nope'), 'x', errors)).toEqual({ ok: false, error: 'missing' })
  })

  it('renames in place and blocks a clash with a different existing folder', () => {
    mkdirSync(join(root, 'a'))
    mkdirSync(join(root, 'b'))
    expect(renameFolder(join(root, 'a'), 'b', errors)).toEqual({ ok: false, error: 'exists' })
    expect(renameFolder(join(root, 'a'), 'c', errors)).toEqual({ ok: true, path: join(root, 'c') })
    expect(renameFolder(join(root, 'c'), 'c', errors)).toEqual({ ok: true, path: join(root, 'c') })
    expect(renameFolder(join(root, 'zzz'), 'q', errors)).toEqual({ ok: false, error: 'missing' })
  })
})

describe('pathsUnder', () => {
  it('keeps only the candidates below the folder, at any depth, once each', () => {
    const dir = join(root, 'src')
    const deep = join(dir, 'a', 'b', 'deep.md')
    const direct = join(dir, 'report.docx')
    expect(
      pathsUnder(dir, [
        deep,
        direct,
        direct,
        dir,
        join(root, 'src-other', 'x.docx'),
        join(root, 'report.docx'),
        join(dir, '..', 'outside.md'),
      ]),
    ).toEqual([deep, direct])
  })

  it('never touches the disk', () => {
    const dir = join(root, 'never-created')
    expect(pathsUnder(dir, [join(dir, 'x.docx')])).toEqual([join(dir, 'x.docx')])
    expect(existsSync(dir)).toBe(false)
  })
})

describe('rebasePath', () => {
  it('leaves a path that did not live under the moved folder alone', () => {
    const outside = touch('notes.md')
    const sibling = touch('src2/deep.md')
    expect(rebasePath(outside, join(root, 'src'), join(root, 'dest'))).toBe(outside)
    expect(rebasePath(sibling, join(root, 'src'), join(root, 'dest'))).toBe(sibling)
  })
})

describe('movePathsInto', () => {
  it('moves files and folders, reporting old → new paths', () => {
    const file = touch('report.docx')
    mkdirSync(join(root, 'src'))
    const nested = touch('src/deep.md')
    mkdirSync(join(root, 'dest'))
    const result = movePathsInto([file, join(root, 'src')], join(root, 'dest'), 'ask', errors, {
      replaceExisting: () => ({ commit: () => {}, rollback: () => {} }),
    })
    expect(result.conflicts).toEqual([])
    expect(result.failed).toEqual([])
    expect(result.moved).toEqual([
      { from: file, to: join(root, 'dest', 'report.docx') },
      { from: join(root, 'src'), to: join(root, 'dest', 'src') },
    ])
    expect(existsSync(join(root, 'dest', 'src', 'deep.md'))).toBe(true)
    expect(rebasePath(nested, join(root, 'src'), join(root, 'dest', 'src'))).toBe(
      join(root, 'dest', 'src', 'deep.md'),
    )
  })

  it("with policy 'ask' leaves collisions alone and reports them", () => {
    const file = touch('report.docx', 'new')
    touch('dest/report.docx', 'old')
    const result = movePathsInto([file], join(root, 'dest'), 'ask', errors, {
      replaceExisting: () => ({ commit: () => {}, rollback: () => {} }),
    })
    expect(result.moved).toEqual([])
    expect(result.conflicts).toEqual([file])
  })

  it("'keepBoth' renames the incoming item; 'replace' hands the old one to the caller first", () => {
    const a = touch('report.docx', 'new')
    touch('dest/report.docx', 'old')
    const kept = movePathsInto([a], join(root, 'dest'), 'keepBoth', errors, {
      replaceExisting: () => ({ commit: () => {}, rollback: () => {} }),
    })
    expect(kept.moved).toEqual([{ from: a, to: join(root, 'dest', 'report (2).docx') }])

    const b = touch('report.docx', 'newer')
    const replaced: string[] = []
    const committed: string[] = []
    const result = movePathsInto([b], join(root, 'dest'), 'replace', errors, {
      replaceExisting: (path) => {
        replaced.push(path)
        rmSync(path)
        return { commit: () => committed.push(path), rollback: () => {} }
      },
    })
    expect(replaced).toEqual([join(root, 'dest', 'report.docx')])
    expect(committed).toEqual([join(root, 'dest', 'report.docx')])
    expect(result.moved).toEqual([{ from: b, to: join(root, 'dest', 'report.docx') }])
  })

  it("'replace' puts the displaced target back when the move itself fails", () => {
    const src = touch('report.docx', 'new')
    const target = touch('dest/report.docx', 'old')
    const parked = join(root, 'dest', '.parked')
    let rolledBack = false
    const result = movePathsInto([src], join(root, 'dest'), 'replace', errors, {
      replaceExisting: (path) => {
        renameSync(path, parked)
        // the source vanishes under us → moveOnDisk fails after the target was parked
        rmSync(src)
        return {
          commit: () => {
            throw new Error('must not commit a failed move')
          },
          rollback: () => {
            rolledBack = true
            renameSync(parked, path)
          },
        }
      },
    })
    expect(rolledBack).toBe(true)
    expect(result.moved).toEqual([])
    expect(result.failed).toHaveLength(1)
    expect(existsSync(target)).toBe(true)
    expect(readFileSync(target, 'utf8')).toBe('old')
  })

  it('refuses to move a folder into itself or a descendant and skips no-op moves', () => {
    mkdirSync(join(root, 'a', 'b'), { recursive: true })
    const file = touch('a/x.md')
    const result = movePathsInto([join(root, 'a'), file], join(root, 'a', 'b'), 'ask', errors, {
      replaceExisting: () => ({ commit: () => {}, rollback: () => {} }),
    })
    expect(result.failed).toEqual([{ path: join(root, 'a'), error: 'badArgs' }])
    expect(result.moved).toEqual([{ from: file, to: join(root, 'a', 'b', 'x.md') }])
    const noop = movePathsInto(
      [join(root, 'a', 'b', 'x.md')],
      join(root, 'a', 'b'),
      'ask',
      errors,
      {
        replaceExisting: () => ({ commit: () => {}, rollback: () => {} }),
      },
    )
    expect(noop).toEqual({ moved: [], conflicts: [], failed: [] })
    expect(isSelfOrDescendant(join(root, 'a'), join(root, 'a', 'b'))).toBe(true)
    expect(isSelfOrDescendant(join(root, 'a'), join(root, 'ab'))).toBe(false)
  })

  it('reports a missing source and a missing target', () => {
    mkdirSync(join(root, 'dest'))
    const ghost = join(root, 'ghost.docx')
    expect(
      movePathsInto([ghost], join(root, 'dest'), 'ask', errors, { replaceExisting: () => {} })
        .failed,
    ).toEqual([{ path: ghost, error: 'missing' }])
    const file = touch('real.docx')
    expect(
      movePathsInto([file], join(root, 'nowhere'), 'ask', errors, { replaceExisting: () => {} })
        .failed,
    ).toEqual([{ path: file, error: 'missing' }])
  })
})

describe('moveOnDisk across volumes', () => {
  const onWindows = process.platform === 'win32'
  const posixOnly = it.skipIf(onWindows)
  let from: string
  let volume: string

  const errnoError = (code: string) => Object.assign(new Error(`${code}: injected`), { code })

  /**
   * What libuv's copy does on POSIX once it fails after opening the destination:
   * an exclusive copy refuses a file that is already there; any other copy has
   * created or truncated the destination by then and removes it on the way out.
   */
  const failLikeLibuv =
    (code: string) =>
    (_source: unknown, target: unknown, mode?: number): never => {
      const path = String(target)
      if (((mode ?? 0) & constants.COPYFILE_EXCL) !== 0 && actualFs.existsSync(path))
        throw errnoError('EEXIST')
      actualFs.writeFileSync(path, '')
      actualFs.rmSync(path)
      throw errnoError(code)
    }

  const crossVolume = () =>
    vi.mocked(renameSync).mockImplementationOnce(() => {
      throw errnoError('EXDEV')
    })

  /** A mount without permission support; Windows copies through CopyFileW, which works there. */
  const refusePlatformCopy = () => {
    if (!onWindows) vi.mocked(copyFileSync).mockImplementation(failLikeLibuv('ENOTSUP'))
  }

  const refuseChmod = () =>
    vi.mocked(chmodSync).mockImplementation(() => {
      throw errnoError('ENOTSUP')
    })

  /** The `nth` file copy fails for good: the platform copy and the plain-write retry both. */
  const failFileCopy = (nth: number) => {
    let files = 0
    let failing = ''
    vi.mocked(copyFileSync).mockImplementation((source, target, mode) => {
      files += 1
      if (files === nth) failing = String(target)
      else if (onWindows) return actualFs.copyFileSync(source, target, mode)
      return failLikeLibuv('ENOTSUP')(source, target, mode)
    })
    vi.mocked(openSync).mockImplementation((...args: Parameters<typeof actualFs.openSync>) => {
      if (args[0] === failing && args[1] === 'wx') throw errnoError('EIO')
      return actualFs.openSync(...args)
    })
    return () => files
  }

  /** Every entry below `dir` (`dir` itself when it is a file): kind and content hash, by relative path. */
  function snapshot(dir: string): Record<string, string> {
    const out: Record<string, string> = {}
    const walk = (path: string, rel: string) => {
      const stats = actualFs.lstatSync(path)
      if (stats.isSymbolicLink()) out[rel] = `link:${actualFs.readlinkSync(path)}`
      else if (stats.isFile())
        out[rel] = createHash('sha256').update(actualFs.readFileSync(path)).digest('hex')
      else {
        out[rel] = 'dir'
        for (const name of actualFs.readdirSync(path).sort())
          walk(join(path, name), rel ? `${rel}/${name}` : name)
      }
    }
    walk(dir, '')
    return out
  }

  function makeTree(dir: string): void {
    actualFs.mkdirSync(join(dir, 'sub', 'deeper'), { recursive: true })
    actualFs.writeFileSync(join(dir, 'a.txt'), 'alpha')
    actualFs.writeFileSync(join(dir, 'sub', 'b.txt'), 'beta')
    // more than one copy buffer
    actualFs.writeFileSync(
      join(dir, 'sub', 'deeper', 'c.bin'),
      Buffer.alloc(1536 * 1024, 0).map((_, i) => i % 251),
    )
    actualFs.writeFileSync(join(dir, 'sub', 'deeper', 'd.txt'), 'delta')
  }

  const modeOf = (path: string) => actualFs.statSync(path).mode & 0o7777

  beforeEach(() => {
    // the outer hooks also go through the mocked functions
    vi.clearAllMocks()
    from = join(root, 'local')
    volume = join(root, 'share')
    actualFs.mkdirSync(from)
    actualFs.mkdirSync(volume)
    // not this move's: it must survive every outcome
    actualFs.writeFileSync(join(volume, 'keep.txt'), 'theirs')
  })

  afterEach(() => {
    vi.resetAllMocks()
    // read-only folders the tests made, so the outer clean-up can remove them
    const unlock = (dir: string) => {
      actualFs.chmodSync(dir, 0o700)
      for (const entry of actualFs.readdirSync(dir, { withFileTypes: true }))
        if (entry.isDirectory()) unlock(join(dir, entry.name))
    }
    unlock(root)
  })

  posixOnly('moves a lone file when the platform copy fails', () => {
    const file = join(from, 'report.docx')
    actualFs.writeFileSync(file, 'report bytes')
    const to = join(volume, 'report.docx')
    crossVolume()
    refusePlatformCopy()

    moveOnDisk(file, to)

    expect(readFileSync(to, 'utf8')).toBe('report bytes')
    expect(existsSync(file)).toBe(false)
    expect(copyFileSync).toHaveBeenCalledWith(file, to, constants.COPYFILE_EXCL)
  })

  posixOnly('moves a folder tree when both the platform copy and chmod are refused', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    const before = snapshot(folder)
    const to = join(volume, 'project')
    crossVolume()
    refusePlatformCopy()
    refuseChmod()
    let atRemoval: Record<string, string> = {}
    vi.mocked(rmSync).mockImplementation((path, options) => {
      if (path === folder) atRemoval = snapshot(to)
      actualFs.rmSync(path, options)
    })

    moveOnDisk(folder, to)

    expect(cpSync).not.toHaveBeenCalled()
    expect(chmodSync).toHaveBeenCalled()
    expect(atRemoval).toEqual(before)
    expect(snapshot(to)).toEqual(before)
    expect(existsSync(folder)).toBe(false)
    expect(readFileSync(join(volume, 'keep.txt'), 'utf8')).toBe('theirs')
  })

  it('removes only what it created when the copy fails part-way, and keeps the source', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    const before = snapshot(folder)
    const sharedBefore = snapshot(volume)
    const to = join(volume, 'project')
    crossVolume()
    const filesCopied = failFileCopy(3)

    expect(() => moveOnDisk(folder, to)).toThrow()

    expect(filesCopied()).toBe(3)
    expect(snapshot(volume)).toEqual(sharedBefore)
    expect(snapshot(folder)).toEqual(before)
    // undone entry by entry, never by a recursive delete of the destination
    expect(rmSync).not.toHaveBeenCalled()
  })

  it('refuses a file already at the destination and keeps its bytes', () => {
    const file = join(from, 'report.docx')
    actualFs.writeFileSync(file, 'mine')
    const to = join(volume, 'keep.txt')
    crossVolume()
    refusePlatformCopy()

    expect(() => moveOnDisk(file, to)).toThrow(expect.objectContaining({ code: 'EEXIST' }))

    expect(readFileSync(to, 'utf8')).toBe('theirs')
    expect(readFileSync(file, 'utf8')).toBe('mine')
  })

  it('refuses a folder already at the destination and leaves what is in it alone', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    const to = join(volume, 'project')
    actualFs.mkdirSync(to)
    actualFs.writeFileSync(join(to, 'a.txt'), 'theirs')
    actualFs.writeFileSync(join(to, 'other.txt'), 'theirs too')
    const sharedBefore = snapshot(volume)
    const before = snapshot(folder)
    crossVolume()
    refusePlatformCopy()

    expect(() => moveOnDisk(folder, to)).toThrow(expect.objectContaining({ code: 'EEXIST' }))

    expect(snapshot(volume)).toEqual(sharedBefore)
    expect(snapshot(folder)).toEqual(before)
  })

  it('keeps the complete destination when removing the source fails', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    const before = snapshot(folder)
    const to = join(volume, 'project')
    crossVolume()
    refusePlatformCopy()
    vi.mocked(rmSync).mockImplementation((path, options) => {
      if (path === folder) throw errnoError('EBUSY')
      actualFs.rmSync(path, options)
    })

    expect(() => moveOnDisk(folder, to)).toThrow(expect.objectContaining({ code: 'EBUSY' }))

    expect(snapshot(to)).toEqual(before)
    expect(snapshot(folder)).toEqual(before)
  })

  posixOnly('gives each folder the mode of its source folder, and each file its own', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    // without the owner's write bit; empty, so the source can still be removed
    actualFs.mkdirSync(join(folder, 'sub', 'empty'))
    actualFs.chmodSync(join(folder, 'sub', 'empty'), 0o500)
    actualFs.chmodSync(join(folder, 'a.txt'), 0o640)
    actualFs.chmodSync(join(folder, 'sub'), 0o711)
    actualFs.chmodSync(folder, 0o750)
    const shared = join(folder, 'shared')
    actualFs.mkdirSync(shared)
    actualFs.chmodSync(shared, 0o2775)
    // some file systems drop the setgid bit; only one that keeps it can show it carried over
    const keepsSetgid = modeOf(shared) === 0o2775
    if (!keepsSetgid) actualFs.chmodSync(shared, 0o775)
    const to = join(volume, 'project')
    crossVolume()

    moveOnDisk(folder, to)

    expect(modeOf(to)).toBe(0o750)
    expect(modeOf(join(to, 'sub'))).toBe(0o711)
    expect(modeOf(join(to, 'sub', 'empty'))).toBe(0o500)
    expect(modeOf(join(to, 'shared'))).toBe(keepsSetgid ? 0o2775 : 0o775)
    expect(modeOf(join(to, 'a.txt'))).toBe(0o640)
    expect(existsSync(folder)).toBe(false)
  })

  posixOnly('keeps a new folder closed to everyone else until the copy is complete', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    actualFs.chmodSync(join(folder, 'sub'), 0o777)
    const to = join(volume, 'project')
    crossVolume()
    const openToOthers: number[] = []
    vi.mocked(copyFileSync).mockImplementation((source, target, mode) => {
      openToOthers.push(modeOf(dirname(String(target))) & 0o077)
      return actualFs.copyFileSync(source, target, mode)
    })

    moveOnDisk(folder, to)

    expect(openToOthers).toEqual([0, 0, 0, 0])
    expect(modeOf(join(to, 'sub'))).toBe(0o777)
  })

  posixOnly('copies a read-only folder and the files inside it', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    const readOnly = join(folder, 'sub', 'locked')
    actualFs.mkdirSync(readOnly)
    actualFs.writeFileSync(join(readOnly, 'e.txt'), 'epsilon')
    actualFs.writeFileSync(join(readOnly, 'f.txt'), 'phi')
    actualFs.chmodSync(readOnly, 0o555)
    const before = snapshot(folder)
    const to = join(volume, 'project')
    crossVolume()
    refusePlatformCopy()

    let failure: unknown
    try {
      moveOnDisk(folder, to)
    } catch (err) {
      failure = err
    }

    expect(snapshot(to)).toEqual(before)
    expect(modeOf(join(to, 'sub', 'locked'))).toBe(0o555)
    // Emptying a read-only source folder is refused to anyone but root, on a
    // local disk too; the complete copy at the destination stays.
    if (process.getuid?.() === 0) expect(failure).toBeUndefined()
    else expect(failure).toEqual(expect.objectContaining({ code: 'EACCES' }))
  })

  posixOnly('rolls back a part-way failure inside a read-only folder', () => {
    const folder = join(from, 'project')
    actualFs.mkdirSync(folder)
    const readOnly = join(folder, 'locked')
    actualFs.mkdirSync(readOnly)
    actualFs.writeFileSync(join(readOnly, 'e.txt'), 'epsilon')
    actualFs.writeFileSync(join(readOnly, 'f.txt'), 'phi')
    actualFs.chmodSync(readOnly, 0o555)
    const before = snapshot(folder)
    const sharedBefore = snapshot(volume)
    const to = join(volume, 'project')
    crossVolume()
    // the second file of the read-only folder, after the first is already in it
    const filesCopied = failFileCopy(2)

    expect(() => moveOnDisk(folder, to)).toThrow()

    expect(filesCopied()).toBe(2)
    expect(snapshot(volume)).toEqual(sharedBefore)
    expect(snapshot(folder)).toEqual(before)
  })

  posixOnly('removes a link it made when the copy fails after it', () => {
    const folder = join(from, 'project')
    actualFs.mkdirSync(folder)
    actualFs.writeFileSync(join(folder, 'b.txt'), 'beta')
    symlinkSync('b.txt', join(folder, 'a-link'))
    const before = snapshot(folder)
    const sharedBefore = snapshot(volume)
    const to = join(volume, 'project')
    crossVolume()
    // the link first, then the file whose copy fails
    vi.mocked(readdirSync).mockImplementation(((path: PathLike) =>
      actualFs.readdirSync(path).sort()) as typeof readdirSync)
    const filesCopied = failFileCopy(1)

    expect(() => moveOnDisk(folder, to)).toThrow()

    expect(filesCopied()).toBe(1)
    expect(cpSync).toHaveBeenCalledTimes(1)
    expect(snapshot(volume)).toEqual(sharedBefore)
    expect(snapshot(folder)).toEqual(before)
  })

  posixOnly('leaves files alone that it reaches through a link put where its folder was', () => {
    const folder = join(from, 'project')
    actualFs.mkdirSync(join(folder, 'sub'), { recursive: true })
    actualFs.writeFileSync(join(folder, 'sub', 'one.txt'), 'one')
    actualFs.writeFileSync(join(folder, 'sub', 'two.txt'), 'two')
    const before = snapshot(folder)
    // someone else's folder, holding files of the same names
    const theirs = join(root, 'theirs')
    actualFs.mkdirSync(theirs)
    actualFs.writeFileSync(join(theirs, 'one.txt'), 'their one')
    actualFs.writeFileSync(join(theirs, 'two.txt'), 'their two')
    const theirsBefore = snapshot(theirs)
    const to = join(volume, 'project')
    crossVolume()
    let files = 0
    vi.mocked(copyFileSync).mockImplementation((source, target, mode) => {
      files += 1
      if (files === 1) return actualFs.copyFileSync(source, target, mode)
      // the first file is in the new folder by now: swap that folder for a link
      actualFs.renameSync(join(to, 'sub'), join(to, 'sub.moved'))
      actualFs.symlinkSync(theirs, join(to, 'sub'))
      throw errnoError('EEXIST')
    })

    expect(() => moveOnDisk(folder, to)).toThrow(expect.objectContaining({ code: 'EEXIST' }))

    expect(files).toBe(2)
    expect(snapshot(theirs)).toEqual(theirsBefore)
    expect(snapshot(folder)).toEqual(before)
  })

  posixOnly('gives no mode to what a link put where its folder was points at', () => {
    const folder = join(from, 'project')
    actualFs.mkdirSync(join(folder, 'outer', 'inner'), { recursive: true })
    actualFs.chmodSync(join(folder, 'outer'), 0o755)
    const theirs = join(root, 'theirs')
    actualFs.mkdirSync(theirs)
    actualFs.chmodSync(theirs, 0o700)
    const to = join(volume, 'project')
    crossVolume()
    // modes go deepest first: `inner`, then `outer`, then the top folder
    vi.mocked(chmodSync).mockImplementationOnce(() => {
      actualFs.renameSync(join(to, 'outer'), join(to, 'outer.moved'))
      actualFs.symlinkSync(theirs, join(to, 'outer'))
    })

    moveOnDisk(folder, to)

    expect(modeOf(theirs)).toBe(0o700)
    expect(existsSync(folder)).toBe(false)
    expect(lstatSync(join(to, 'outer.moved', 'inner')).isDirectory()).toBe(true)
  })

  posixOnly('undoes a failed copy into a folder that is reached through a link', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    const before = snapshot(folder)
    const sharedBefore = snapshot(volume)
    const linked = join(root, 'linked')
    symlinkSync(volume, linked)
    crossVolume()
    const filesCopied = failFileCopy(3)

    expect(() => moveOnDisk(folder, join(linked, 'project'))).toThrow()

    expect(filesCopied()).toBe(3)
    expect(snapshot(volume)).toEqual(sharedBefore)
    expect(snapshot(folder)).toEqual(before)
  })

  posixOnly('sets folder modes in a folder that is reached through a link', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    // a mode that no umask leaves a new folder with; empty, so the source can still be removed
    actualFs.mkdirSync(join(folder, 'sub', 'empty'))
    actualFs.chmodSync(join(folder, 'sub', 'empty'), 0o500)
    const linked = join(root, 'linked')
    symlinkSync(volume, linked)
    crossVolume()

    moveOnDisk(folder, join(linked, 'project'))

    expect(modeOf(join(volume, 'project', 'sub', 'empty'))).toBe(0o500)
    expect(existsSync(folder)).toBe(false)
  })

  posixOnly('sets the mode of a folder before its parent closes to the owner', () => {
    const folder = join(from, 'project')
    actualFs.mkdirSync(join(folder, 'inner'), { recursive: true })
    actualFs.chmodSync(join(folder, 'inner'), 0o500)
    // A mount can report a mode it does not enforce: this folder was read
    // although its mode has no search bit.
    vi.mocked(lstatSync).mockImplementation(((path: PathLike) => {
      const stats = actualFs.lstatSync(path)
      if (path === folder) stats.mode = (stats.mode & ~0o777) | 0o600
      return stats
    }) as typeof lstatSync)
    const to = join(volume, 'project')
    crossVolume()

    moveOnDisk(folder, to)

    expect(modeOf(to)).toBe(0o600)
    actualFs.chmodSync(to, 0o700)
    expect(modeOf(join(to, 'inner'))).toBe(0o500)
    expect(existsSync(folder)).toBe(false)
  })

  it('still moves when one chmod of the final mode pass is refused', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    const before = snapshot(folder)
    const to = join(volume, 'project')
    crossVolume()
    refusePlatformCopy()
    vi.mocked(chmodSync).mockImplementationOnce(() => {
      throw errnoError('EPERM')
    })

    moveOnDisk(folder, to)

    expect(chmodSync).toHaveBeenCalledTimes(3)
    expect(rmSync).toHaveBeenCalledWith(folder, { recursive: true, force: true })
    expect(existsSync(folder)).toBe(false)
    expect(snapshot(to)).toEqual(before)
  })

  posixOnly('carries a symbolic link inside the folder over as a link', () => {
    const folder = join(from, 'project')
    makeTree(folder)
    symlinkSync('a.txt', join(folder, 'link.txt'))
    const to = join(volume, 'project')
    crossVolume()
    refusePlatformCopy()

    moveOnDisk(folder, to)

    expect(lstatSync(join(to, 'link.txt')).isSymbolicLink()).toBe(true)
    expect(readFileSync(join(to, 'a.txt'), 'utf8')).toBe('alpha')
    expect(existsSync(folder)).toBe(false)
  })
})

describe('helpers', () => {
  it('uniqueNameIn counts up from (2)', () => {
    touch('a.md')
    touch('a (2).md')
    expect(uniqueNameIn(root, 'a.md')).toBe('a (3).md')
    expect(uniqueNameIn(root, 'b.md')).toBe('b.md')
  })

  it('uniqueNameIn appends the counter to a dotted folder name whole', () => {
    // a folder has no extension: 'v1.0' must not become 'v1 (2).0'
    mkdirSync(join(root, 'v1.0'))
    expect(uniqueNameIn(root, 'v1.0', true)).toBe('v1.0 (2)')
    mkdirSync(join(root, 'v1.0 (2)'))
    expect(uniqueNameIn(root, 'v1.0', true)).toBe('v1.0 (3)')
    // a file keeps the extension split
    touch('report.docx')
    expect(uniqueNameIn(root, 'report.docx')).toBe('report (2).docx')
  })

  it("'keepBoth' appends the counter to a dotted folder name whole", () => {
    mkdirSync(join(root, 'src', 'v1.0'), { recursive: true })
    mkdirSync(join(root, 'dest', 'v1.0'), { recursive: true })
    const moved = movePathsInto(
      [join(root, 'src', 'v1.0')],
      join(root, 'dest'),
      'keepBoth',
      errors,
      {
        replaceExisting: () => ({ commit: () => {}, rollback: () => {} }),
      },
    )
    expect(moved.moved).toEqual([
      { from: join(root, 'src', 'v1.0'), to: join(root, 'dest', 'v1.0 (2)') },
    ])
  })

  it('describeRoot creates a missing root and reports it usable', () => {
    const fresh = join(root, 'GenOffice')
    expect(describeRoot(fresh)).toEqual({
      path: fresh,
      name: 'GenOffice',
      usable: true,
      readable: true,
      removable: false,
    })
  })

  it('describeRoot reports a path blocked by a file as unusable', () => {
    const blocked = touch('blocked')
    expect(describeRoot(blocked).usable).toBe(false)
  })
})

describe('filesystem roots', () => {
  const fsRoot = parse(tmpdir()).root

  it('treats descendants as inside a root that already ends with a separator', () => {
    expect(isInsideRoot(fsRoot, fsRoot)).toBe(true)
    expect(isInsideRoot(fsRoot, tmpdir())).toBe(true)

    expect(pathsUnder(fsRoot, [fsRoot, tmpdir()])).toEqual([tmpdir()])

    expect(isSelfOrDescendant(fsRoot, tmpdir())).toBe(true)
  })

  it('rebases descendants from a filesystem root without doubling the separator', () => {
    const destination = join(root, 'rebased')

    expect(rebasePath(tmpdir(), fsRoot, destination)).toBe(
      join(destination, relative(fsRoot, tmpdir())),
    )
  })
})
