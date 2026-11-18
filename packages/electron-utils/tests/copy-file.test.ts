import { randomBytes } from 'node:crypto'
import {
  chmodSync,
  closeSync,
  constants,
  copyFileSync,
  existsSync,
  linkSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  openSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  symlinkSync,
  unlinkSync,
  writeFileSync,
  writeSync,
} from 'node:fs'
import { copyFile, open, stat, unlink, type FileHandle } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  copyFileExclusive,
  copyFileExclusiveSync,
  copyFileFresh,
  copyFileInPlace,
} from '../src/copy-file'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    copyFile: vi.fn(actual.copyFile),
    open: vi.fn(actual.open),
    stat: vi.fn(actual.stat),
    unlink: vi.fn(actual.unlink),
  }
})

vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    closeSync: vi.fn(actual.closeSync),
    copyFileSync: vi.fn(actual.copyFileSync),
    openSync: vi.fn(actual.openSync),
    unlinkSync: vi.fn(actual.unlinkSync),
    writeSync: vi.fn(actual.writeSync),
  }
})

const actualFs = await vi.importActual<typeof import('node:fs')>('node:fs')
const actualPromises = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises')

const hostPlatform = Object.getOwnPropertyDescriptor(process, 'platform')!
const onWindowsHost = hostPlatform.value === 'win32'
const usePlatform = (platform: NodeJS.Platform) =>
  Object.defineProperty(process, 'platform', { ...hostPlatform, value: platform })

/** 2.5 MiB: more than one copy buffer, and not a whole number of them */
const BIG_BYTES = Math.floor(2.5 * 1024 * 1024)

let dir = ''
const workDir = () => {
  dir = mkdtempSync(join(tmpdir(), 'eu-cf-'))
  return dir
}

afterEach(() => {
  Object.defineProperty(process, 'platform', hostPlatform)
  vi.resetAllMocks()
  if (dir) rmSync(dir, { recursive: true, force: true })
  dir = ''
})

const errnoError = (code: string) => Object.assign(new Error(`${code}: injected`), { code })

/**
 * What libuv's copy does on POSIX once it fails after opening the destination:
 * an exclusive copy refuses a file that is already there; any other copy has
 * created or truncated the destination by then and removes it on the way out.
 */
function failLikeLibuv(code: string) {
  return (_source: unknown, target: unknown, mode?: number): never => {
    const path = String(target)
    if (((mode ?? 0) & constants.COPYFILE_EXCL) !== 0 && existsSync(path))
      throw errnoError('EEXIST')
    writeFileSync(path, '')
    rmSync(path)
    throw errnoError(code)
  }
}

/** Run `patch` on every handle the code under test opens for writing `path`. */
function patchWriteHandle(path: string, patch: (handle: FileHandle) => void): void {
  vi.mocked(open).mockImplementation(async (...args) => {
    const handle = await actualPromises.open(...(args as Parameters<typeof actualPromises.open>))
    if (args[0] === path && args[1] !== 'r') patch(handle)
    return handle
  })
}

/** The descriptors the code under test opens for writing `path`. */
function watchWriteDescriptors(path: string): Set<number> {
  const descriptors = new Set<number>()
  vi.mocked(openSync).mockImplementation((...args) => {
    const descriptor = actualFs.openSync(...(args as Parameters<typeof actualFs.openSync>))
    if (args[0] === path && args[1] !== 'r') descriptors.add(descriptor)
    return descriptor
  })
  return descriptors
}

interface Variant {
  toString(): string
  copy(source: string, target: string): Promise<void>
  failPlatformCopy(code: string): void
  /** the platform copy fails, and a file appears at the target before the fallback opens it */
  failPlatformCopyThenRace(code: string, theirs: string): void
  platformCopyCalls(): unknown[][]
  openCalls(): unknown[][]
  unlinksOf(path: string): number
  /** every write to `path` lands at most `limit` bytes */
  shortWrites(path: string, limit: number): void
  /** the second write to `path` fails */
  failSecondWrite(path: string, code: string): void
  /** closing `path` fails, after the descriptor is released */
  failClose(path: string, code: string): void
}

const asyncVariant: Variant = {
  toString: () => 'copyFileExclusive',
  copy: (source, target) => copyFileExclusive(source, target),
  failPlatformCopy: (code) => {
    vi.mocked(copyFile).mockImplementation(async (source, target, mode) =>
      failLikeLibuv(code)(source, target, mode),
    )
  },
  failPlatformCopyThenRace: (code, theirs) => {
    vi.mocked(copyFile).mockImplementation(async (_source, target) => {
      writeFileSync(String(target), theirs)
      throw errnoError(code)
    })
  },
  platformCopyCalls: () => vi.mocked(copyFile).mock.calls,
  openCalls: () => vi.mocked(open).mock.calls,
  unlinksOf: (path) => vi.mocked(unlink).mock.calls.filter(([removed]) => removed === path).length,
  shortWrites: (path, limit) => {
    patchWriteHandle(path, (handle) => {
      const write = handle.write.bind(handle)
      Object.assign(handle, {
        write: (buffer: Buffer, offset: number, length: number, position: null) =>
          write(buffer, offset, Math.min(length, limit), position),
      })
    })
  },
  failSecondWrite: (path, code) => {
    patchWriteHandle(path, (handle) => {
      const write = handle.write.bind(handle)
      let calls = 0
      Object.assign(handle, {
        write: (buffer: Buffer, offset: number, length: number, position: null) => {
          calls += 1
          if (calls === 2) return Promise.reject(errnoError(code))
          return write(buffer, offset, length, position)
        },
      })
    })
  },
  failClose: (path, code) => {
    patchWriteHandle(path, (handle) => {
      const close = handle.close.bind(handle)
      Object.assign(handle, {
        close: async () => {
          await close()
          throw errnoError(code)
        },
      })
    })
  },
}

const syncVariant: Variant = {
  toString: () => 'copyFileExclusiveSync',
  copy: async (source, target) => copyFileExclusiveSync(source, target),
  failPlatformCopy: (code) => {
    vi.mocked(copyFileSync).mockImplementation((source, target, mode) =>
      failLikeLibuv(code)(source, target, mode),
    )
  },
  failPlatformCopyThenRace: (code, theirs) => {
    vi.mocked(copyFileSync).mockImplementation((_source, target) => {
      writeFileSync(String(target), theirs)
      throw errnoError(code)
    })
  },
  platformCopyCalls: () => vi.mocked(copyFileSync).mock.calls,
  openCalls: () => vi.mocked(openSync).mock.calls,
  unlinksOf: (path) =>
    vi.mocked(unlinkSync).mock.calls.filter(([removed]) => removed === path).length,
  shortWrites: (path, limit) => {
    const descriptors = watchWriteDescriptors(path)
    vi.mocked(writeSync).mockImplementation(((
      descriptor: number,
      buffer: Buffer,
      offset: number,
      length: number,
      position: null,
    ) =>
      actualFs.writeSync(
        descriptor,
        buffer,
        offset,
        descriptors.has(descriptor) ? Math.min(length, limit) : length,
        position,
      )) as typeof writeSync)
  },
  failSecondWrite: (path, code) => {
    const descriptors = watchWriteDescriptors(path)
    let calls = 0
    vi.mocked(writeSync).mockImplementation(((
      descriptor: number,
      buffer: Buffer,
      offset: number,
      length: number,
      position: null,
    ) => {
      if (descriptors.has(descriptor)) {
        calls += 1
        if (calls === 2) throw errnoError(code)
      }
      return actualFs.writeSync(descriptor, buffer, offset, length, position)
    }) as typeof writeSync)
  },
  failClose: (path, code) => {
    const descriptors = watchWriteDescriptors(path)
    vi.mocked(closeSync).mockImplementation((descriptor) => {
      actualFs.closeSync(descriptor)
      if (descriptors.has(descriptor)) throw errnoError(code)
    })
  },
}

describe.each([asyncVariant, syncVariant])('%s', (variant) => {
  it('uses the platform copy alone when it works', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    writeFileSync(source, 'bytes')
    await variant.copy(source, target)
    expect(readFileSync(target, 'utf8')).toBe('bytes')
    expect(variant.platformCopyCalls()).toEqual([[source, target, constants.COPYFILE_EXCL]])
    expect(variant.openCalls()).toEqual([])
  })

  // A mount without permission support fails the copy at fchmod (ENOTSUP from
  // Node's own libuv, EACCES from a distribution's), and libuv then removes
  // what it created. No list of codes decides the fallback.
  it.each(['ENOTSUP', 'EACCES', 'EIO', 'EPERM'])(
    'copies through plain reads and writes when the platform copy fails with %s',
    async (code) => {
      usePlatform('linux')
      const source = join(workDir(), 'a.bin')
      const target = join(dir, 'b.bin')
      const data = randomBytes(BIG_BYTES)
      writeFileSync(source, data)
      variant.failPlatformCopy(code)
      await variant.copy(source, target)
      expect(readFileSync(target).equals(data)).toBe(true)
      expect(readdirSync(dir).sort()).toEqual(['a.bin', 'b.bin'])
    },
  )

  it('refuses an existing target and leaves it untouched', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    writeFileSync(source, 'mine')
    writeFileSync(target, 'theirs')
    await expect(variant.copy(source, target)).rejects.toMatchObject({ code: 'EEXIST' })
    expect(readFileSync(target, 'utf8')).toBe('theirs')
    expect(variant.openCalls()).toEqual([])
    expect(variant.unlinksOf(target)).toBe(0)
  })

  it('refuses an existing target on a mount where the platform copy never works', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    writeFileSync(source, 'mine')
    writeFileSync(target, 'theirs')
    variant.failPlatformCopy('ENOTSUP')
    await expect(variant.copy(source, target)).rejects.toMatchObject({ code: 'EEXIST' })
    expect(readFileSync(target, 'utf8')).toBe('theirs')
    expect(variant.openCalls()).toEqual([])
    expect(variant.unlinksOf(target)).toBe(0)
  })

  // EEXIST is kept for a file that was there before the call: a caller that
  // tries the next name on EEXIST would otherwise leave a file behind per name.
  it('reports the platform copy error, and touches nothing, when the target appeared after it failed', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    writeFileSync(source, 'mine')
    variant.failPlatformCopyThenRace('ENOTSUP', 'theirs')
    await expect(variant.copy(source, target)).rejects.toMatchObject({ code: 'ENOTSUP' })
    expect(readFileSync(target, 'utf8')).toBe('theirs')
    expect(variant.unlinksOf(target)).toBe(0)
  })

  it('removes the target it created and reports its own error when a write fails part-way', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    writeFileSync(source, randomBytes(BIG_BYTES))
    variant.failPlatformCopy('ENOTSUP')
    variant.failSecondWrite(target, 'EIO')
    await expect(variant.copy(source, target)).rejects.toMatchObject({ code: 'EIO' })
    expect(readdirSync(dir)).toEqual(['a.bin'])
    expect(variant.unlinksOf(target)).toBe(1)
  })

  // gvfs sends the bytes to the server when the file is closed.
  it('treats a failed close as a failed copy and removes the target', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    writeFileSync(source, 'bytes')
    variant.failPlatformCopy('ENOTSUP')
    variant.failClose(target, 'EIO')
    await expect(variant.copy(source, target)).rejects.toMatchObject({ code: 'EIO' })
    expect(readdirSync(dir)).toEqual(['a.bin'])
    expect(variant.unlinksOf(target)).toBe(1)
  })

  it('finishes a copy whose writes land short', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    const data = randomBytes(BIG_BYTES)
    writeFileSync(source, data)
    variant.failPlatformCopy('ENOTSUP')
    variant.shortWrites(target, 700_001)
    await variant.copy(source, target)
    expect(readFileSync(target).equals(data)).toBe(true)
  })

  it('reports a missing source and creates no target', async () => {
    usePlatform('linux')
    const target = join(workDir(), 'b.bin')
    await expect(variant.copy(join(dir, 'nope.bin'), target)).rejects.toMatchObject({
      code: 'ENOENT',
    })
    expect(readdirSync(dir)).toEqual([])
  })

  // Reading a pipe or a device has no end, and a folder cannot be read.
  it('refuses a source that is not a regular file before it creates anything', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'folder')
    const target = join(dir, 'b.bin')
    mkdirSync(source)
    variant.failPlatformCopy('ENOTSUP')
    await expect(variant.copy(source, target)).rejects.toMatchObject({ code: 'EISDIR' })
    expect(readdirSync(dir)).toEqual(['folder'])
    expect(variant.openCalls().filter(([opened]) => opened === target)).toEqual([])
  })

  it.skipIf(onWindowsHost).each([
    ['600', 0o600],
    ['640', 0o640],
    ['755', 0o755],
  ])('gives the fallback copy of a %s file no wider permissions', async (_octal, mode) => {
    usePlatform('linux')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    writeFileSync(source, 'bytes')
    chmodSync(source, mode)
    variant.failPlatformCopy('ENOTSUP')
    await variant.copy(source, target)
    expect(statSync(target).mode & 0o777 & ~mode).toBe(0)
  })

  // CopyFileW has no fchmod step and does not remove its target when it fails.
  it('runs the platform copy alone on Windows', async () => {
    usePlatform('win32')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    writeFileSync(source, 'bytes')
    await variant.copy(source, target)
    expect(readFileSync(target, 'utf8')).toBe('bytes')
    expect(variant.platformCopyCalls()).toEqual([[source, target, constants.COPYFILE_EXCL]])
    expect(variant.openCalls()).toEqual([])
  })

  it('reports the platform copy error as is on Windows', async () => {
    usePlatform('win32')
    const source = join(workDir(), 'a.bin')
    const target = join(dir, 'b.bin')
    writeFileSync(source, 'bytes')
    variant.failPlatformCopy('ENOTSUP')
    await expect(variant.copy(source, target)).rejects.toMatchObject({ code: 'ENOTSUP' })
    expect(variant.openCalls()).toEqual([])
    expect(existsSync(target)).toBe(false)
  })
})

describe('copyFileInPlace', () => {
  // A failed platform copy has truncated the target and then unlinks it, so
  // the step that writes over an existing document never calls it.
  it('replaces the contents without the platform copy', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin')
    const data = randomBytes(BIG_BYTES)
    writeFileSync(source, data)
    writeFileSync(target, randomBytes(BIG_BYTES + 4096))
    await copyFileInPlace(source, target)
    expect(readFileSync(target).equals(data)).toBe(true)
    expect(vi.mocked(copyFile)).not.toHaveBeenCalled()
    expect(vi.mocked(unlink)).not.toHaveBeenCalled()
  })

  it('creates a target that is not there yet', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin')
    writeFileSync(source, 'new')
    await copyFileInPlace(source, target)
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(vi.mocked(copyFile)).not.toHaveBeenCalled()
  })

  it('never removes the target, even when the write fails', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin')
    writeFileSync(source, randomBytes(BIG_BYTES))
    writeFileSync(target, 'old')
    asyncVariant.failSecondWrite(target, 'EIO')
    await expect(copyFileInPlace(source, target)).rejects.toMatchObject({ code: 'EIO' })
    expect(existsSync(target)).toBe(true)
    expect(vi.mocked(unlink)).not.toHaveBeenCalled()
    expect(vi.mocked(copyFile)).not.toHaveBeenCalled()
  })

  it('leaves the target as it was when it cannot be opened', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin')
    writeFileSync(source, 'new')
    writeFileSync(target, 'old')
    vi.mocked(open).mockImplementation(async (...args) => {
      if (args[0] === target) throw errnoError('EACCES')
      return actualPromises.open(...(args as Parameters<typeof actualPromises.open>))
    })
    await expect(copyFileInPlace(source, target)).rejects.toMatchObject({ code: 'EACCES' })
    expect(readFileSync(target, 'utf8')).toBe('old')
    expect(vi.mocked(unlink)).not.toHaveBeenCalled()
  })

  // Opening a file for writing truncates it: copying a file onto itself would empty it.
  it('writes nothing when both paths are the same path', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'doc.bin')
    writeFileSync(source, 'keep')
    await copyFileInPlace(source, source)
    expect(readFileSync(source, 'utf8')).toBe('keep')
    expect(vi.mocked(open)).not.toHaveBeenCalled()
  })

  it.skipIf(onWindowsHost)(
    'writes nothing when the target is a hard link to the source',
    async () => {
      usePlatform('linux')
      const source = join(workDir(), 'doc.bin')
      const target = join(dir, 'link.bin')
      writeFileSync(source, 'keep')
      linkSync(source, target)
      await copyFileInPlace(source, target)
      expect(readFileSync(source, 'utf8')).toBe('keep')
      expect(vi.mocked(open)).not.toHaveBeenCalled()
    },
  )

  it.skipIf(onWindowsHost)(
    'writes nothing when the target is a symbolic link to the source',
    async () => {
      usePlatform('linux')
      const source = join(workDir(), 'doc.bin')
      const target = join(dir, 'link.bin')
      writeFileSync(source, 'keep')
      symlinkSync(source, target)
      await copyFileInPlace(source, target)
      expect(readFileSync(source, 'utf8')).toBe('keep')
      expect(vi.mocked(open)).not.toHaveBeenCalled()
    },
  )

  /** Every stat reports file number 0, as a mount that has none does. */
  const reportNoFileNumber = () => {
    vi.mocked(stat).mockImplementation((async (path: string, options: { bigint: true }) =>
      Object.assign(await actualPromises.stat(path, options), { ino: 0n })) as typeof stat)
  }

  // Two different files both report 0 there. Taking them for one file would
  // skip the write and lose the new bytes.
  it('writes when the mount reports no file number', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin')
    writeFileSync(source, 'new')
    writeFileSync(target, 'old')
    reportNoFileNumber()
    await copyFileInPlace(source, target)
    expect(readFileSync(target, 'utf8')).toBe('new')
  })

  // A mount can hand two files the same number. One file has one size, so
  // two sizes mean two files.
  it('writes when two files of different sizes share a file number', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin')
    writeFileSync(source, 'new and longer')
    writeFileSync(target, 'old')
    vi.mocked(stat).mockImplementation((async (path: string, options: { bigint: true }) =>
      Object.assign(await actualPromises.stat(path, options), {
        ino: 42n,
        dev: 7n,
      })) as typeof stat)
    await copyFileInPlace(source, target)
    expect(readFileSync(target, 'utf8')).toBe('new and longer')
  })

  it.skipIf(onWindowsHost)(
    'still recognises a symbolic link to the source when the mount reports no file number',
    async () => {
      usePlatform('linux')
      const source = join(workDir(), 'doc.bin')
      const target = join(dir, 'link.bin')
      writeFileSync(source, 'keep')
      symlinkSync(source, target)
      reportNoFileNumber()
      await copyFileInPlace(source, target)
      expect(readFileSync(source, 'utf8')).toBe('keep')
      expect(vi.mocked(open)).not.toHaveBeenCalled()
    },
  )

  // CopyFileW leaves its target alone when it fails, so Windows keeps the call it always made.
  it('runs the platform copy alone on Windows', async () => {
    usePlatform('win32')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin')
    writeFileSync(source, 'new')
    writeFileSync(target, 'old')
    await copyFileInPlace(source, target)
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(vi.mocked(copyFile).mock.calls).toEqual([[source, target]])
    expect(vi.mocked(open)).not.toHaveBeenCalled()
  })

  it('reports the platform copy error as is on Windows', async () => {
    usePlatform('win32')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin')
    writeFileSync(source, 'new')
    writeFileSync(target, 'old')
    vi.mocked(copyFile).mockRejectedValue(errnoError('EBUSY'))
    await expect(copyFileInPlace(source, target)).rejects.toMatchObject({ code: 'EBUSY' })
    expect(readFileSync(target, 'utf8')).toBe('old')
    expect(vi.mocked(open)).not.toHaveBeenCalled()
    expect(vi.mocked(unlink)).not.toHaveBeenCalled()
  })
})

describe('copyFileFresh', () => {
  it('creates the copy when nothing holds the name', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin.unsaved.tmp')
    writeFileSync(source, 'new')
    await copyFileFresh(source, target)
    expect(readFileSync(target, 'utf8')).toBe('new')
  })

  it('replaces the file that holds the name', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin.unsaved.tmp')
    writeFileSync(source, 'new')
    writeFileSync(target, 'an earlier copy, longer than the new one')
    await copyFileFresh(source, target)
    expect(readFileSync(target, 'utf8')).toBe('new')
  })

  it('replaces the file that holds the name where the platform copy never works', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin.unsaved.tmp')
    writeFileSync(source, 'new')
    writeFileSync(target, 'an earlier copy')
    asyncVariant.failPlatformCopy('ENOTSUP')
    await copyFileFresh(source, target)
    expect(readFileSync(target, 'utf8')).toBe('new')
  })

  // In a shared folder someone else can put a link under the name first.
  // Writing through it would hand the bytes to whatever it points at.
  it.skipIf(onWindowsHost)('replaces a symbolic link instead of writing through it', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const elsewhere = join(dir, 'elsewhere.txt')
    const target = join(dir, 'doc.bin.unsaved.tmp')
    writeFileSync(source, 'new')
    writeFileSync(elsewhere, 'not ours to change')
    symlinkSync(elsewhere, target)
    await copyFileFresh(source, target)
    expect(readFileSync(elsewhere, 'utf8')).toBe('not ours to change')
    expect(lstatSync(target).isFile()).toBe(true)
    expect(readFileSync(target, 'utf8')).toBe('new')
  })

  it('writes nothing when the name cannot be cleared', async () => {
    usePlatform('linux')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin.unsaved.tmp')
    writeFileSync(source, 'new')
    writeFileSync(target, 'someone else holds this')
    vi.mocked(unlink).mockRejectedValue(errnoError('EPERM'))
    await expect(copyFileFresh(source, target)).rejects.toMatchObject({ code: 'EPERM' })
    expect(readFileSync(target, 'utf8')).toBe('someone else holds this')
    expect(vi.mocked(copyFile)).not.toHaveBeenCalled()
    expect(vi.mocked(open)).not.toHaveBeenCalled()
  })

  it('runs the platform copy alone on Windows', async () => {
    usePlatform('win32')
    const source = join(workDir(), 'new.bin')
    const target = join(dir, 'doc.bin.unsaved.tmp')
    writeFileSync(source, 'new')
    writeFileSync(target, 'an earlier copy')
    await copyFileFresh(source, target)
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(vi.mocked(copyFile).mock.calls).toEqual([[source, target]])
    expect(vi.mocked(unlink)).not.toHaveBeenCalled()
    expect(vi.mocked(open)).not.toHaveBeenCalled()
  })
})
