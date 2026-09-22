import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  renameSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  createFolder,
  describeRoot,
  isHiddenEntry,
  isInsideRoot,
  isSelfOrDescendant,
  listFolder,
  movePathsInto,
  pathsUnder,
  rebasePath,
  renameFolder,
  uniqueNameIn,
} from '../src/main/folder-tree'

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

describe('helpers', () => {
  it('uniqueNameIn counts up from (2)', () => {
    touch('a.md')
    touch('a (2).md')
    expect(uniqueNameIn(root, 'a.md')).toBe('a (3).md')
    expect(uniqueNameIn(root, 'b.md')).toBe('b.md')
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
