/**
 * Saving promotes a same-directory temp file over the target with rename.
 * Windows AV/indexer/cloud-sync locks make that rename fail transiently with
 * EPERM/EACCES/EBUSY (user report: "EPERM: operation not permitted, rename
 * .tmp.xlsx → …"), so the promotion retries, falls back to an in-place copy,
 * and surfaces a stable, localizable message when the target stays locked.
 */
import { constants, existsSync } from 'node:fs'
import {
  copyFile,
  mkdir,
  mkdtemp,
  open,
  readFile,
  readdir,
  rename,
  rm,
  stat,
  truncate,
  unlink,
  writeFile,
} from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { promoteFileAtomically } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    copyFile: vi.fn(actual.copyFile),
    open: vi.fn(actual.open),
    rename: vi.fn(actual.rename),
    stat: vi.fn(actual.stat),
    unlink: vi.fn(actual.unlink),
  }
})
const copyFileMock = vi.mocked(copyFile)
const openMock = vi.mocked(open)
const renameMock = vi.mocked(rename)
const statMock = vi.mocked(stat)
const unlinkMock = vi.mocked(unlink)

const scratches: string[] = []
const actualCopyFile = copyFileMock.getMockImplementation()!
const actualRename = renameMock.getMockImplementation()!
const actualOpen = openMock.getMockImplementation()!
const actualStat = statMock.getMockImplementation()!
const actualUnlink = unlinkMock.getMockImplementation()!

const hostPlatform = Object.getOwnPropertyDescriptor(process, 'platform')!
const usePlatform = (platform: NodeJS.Platform) =>
  Object.defineProperty(process, 'platform', { ...hostPlatform, value: platform })

const busy = () => Object.assign(new Error('EBUSY: resource busy'), { code: 'EBUSY' })

/** Directory chmod does not prevent rename on Windows or when running as root. */
function blockRename(temporary: string, target: string): void {
  renameMock.mockImplementation(async (from, to) => {
    if (String(from) === temporary && String(to) === target) throw busy()
    return actualRename(from, to)
  })
}

async function scratchDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-promote-test-'))
  scratches.push(dir)
  return dir
}

afterEach(async () => {
  Object.defineProperty(process, 'platform', hostPlatform)
  vi.restoreAllMocks()
  openMock.mockReset()
  openMock.mockImplementation(actualOpen)
  statMock.mockReset()
  statMock.mockImplementation(actualStat)
  unlinkMock.mockReset()
  unlinkMock.mockImplementation(actualUnlink)
  copyFileMock.mockReset()
  copyFileMock.mockImplementation(actualCopyFile)
  renameMock.mockReset()
  renameMock.mockImplementation(actualRename)
  for (const dir of scratches.splice(0)) {
    await rm(dir, { recursive: true, force: true })
  }
})

describe('promoteFileAtomically', () => {
  it('replaces the target and removes the temp file', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    await promoteFileAtomically(temporary, target)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
    await expect(stat(temporary)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('falls back to an in-place copy when only the rename is blocked', async () => {
    // The Windows path, which keeps the platform copy for the in-place step.
    usePlatform('win32')
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    blockRename(temporary, target)
    await promoteFileAtomically(temporary, target)
    expect(renameMock.mock.calls.length).toBeGreaterThan(1)
    expect(copyFileMock).toHaveBeenCalledWith(temporary, target)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
    await expect(stat(temporary)).rejects.toMatchObject({ code: 'ENOENT' })
  }, 15_000)

  it('reports a persistently locked target with a stable localizable message', async () => {
    // The Windows path, which keeps the platform copy for the in-place step.
    usePlatform('win32')
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    // rename AND in-place copy both refused — the Excel-holds-the-file case
    blockRename(temporary, target)
    copyFileMock.mockImplementation(async (src, dest, mode) => {
      if (String(src) === temporary && String(dest) === target) throw busy()
      return actualCopyFile(src, dest, mode)
    })
    await expect(promoteFileAtomically(temporary, target)).rejects.toThrow(
      'The save target is locked by another program',
    )
    // the finished bytes survive the failure for the caller's cleanup/retry
    expect(await readFile(temporary, 'utf8')).toBe('new-bytes')
    expect(await readFile(target, 'utf8')).toBe('old-bytes')
  }, 15_000)

  it('restores the target when the in-place copy dies after truncating it', async () => {
    // The Windows path, which keeps the platform copy for the in-place step.
    usePlatform('win32')
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    blockRename(temporary, target)
    // the backup copy runs for real; the copy over the target truncates it
    // and then fails the way a lock acquired mid-write does
    copyFileMock.mockImplementation(async (src, dest, mode) => {
      if (String(src) !== temporary) return actualCopyFile(src, dest, mode)
      await truncate(String(dest))
      throw busy()
    })
    await expect(promoteFileAtomically(temporary, target)).rejects.toThrow(
      'The save target is locked by another program',
    )
    expect(await readFile(target, 'utf8')).toBe('old-bytes')
    expect(await readFile(temporary, 'utf8')).toBe('new-bytes')
  }, 15_000)

  it('names the surviving backup when the target cannot be restored either', async () => {
    // The Windows path, which keeps the platform copy for the in-place step.
    usePlatform('win32')
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    blockRename(temporary, target)
    copyFileMock.mockImplementation(async (src, dest, mode) => {
      if (String(src) === temporary) {
        await truncate(String(dest))
        throw busy()
      }
      if (dirname(String(dest)) === locked) throw busy()
      return actualCopyFile(src, dest, mode)
    })
    const failure = await promoteFileAtomically(temporary, target).catch((error: Error) => error)
    expect(failure?.message).toContain('preserved at: ')
    const survivor = failure!.message.split('preserved at: ')[1] ?? ''
    // The locked directory also refuses the recovered copy, so the tmp backup stays.
    expect(survivor.startsWith(tmpdir())).toBe(true)
    expect(basename(survivor)).toMatch(/^book\.recovered-[0-9a-f-]+\.xlsx$/)
    expect(await readFile(survivor, 'utf8')).toBe('old-bytes')
    await rm(survivor, { force: true })
  }, 15_000)

  it('propagates non-retryable errors untouched', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.missing.tmp.xlsx')
    await expect(promoteFileAtomically(temporary, join(dir, 'book.xlsx'))).rejects.toMatchObject({
      code: 'ENOENT',
    })
  })
})

/**
 * A mount without permission support (gvfs SMB, WebDAV) refuses the platform
 * copy into it the way libuv fails there on POSIX: an exclusive copy refuses a
 * file that is already there; any other copy has created or truncated the
 * destination by then and removes it on the way out. The backup in tmpdir is
 * on a local disk and copies normally.
 */
function platformCopyFailsLikeLibuvIn(share: string): void {
  copyFileMock.mockImplementation(async (src, dest, mode) => {
    const path = String(dest)
    if (dirname(path) !== share) return actualCopyFile(src, dest, mode)
    if (((mode ?? 0) & constants.COPYFILE_EXCL) !== 0 && existsSync(path))
      throw Object.assign(new Error('EEXIST: file already exists'), { code: 'EEXIST' })
    await writeFile(path, '')
    await actualUnlink(path)
    throw Object.assign(new Error('ENOTSUP: operation not supported'), { code: 'ENOTSUP' })
  })
}

function refuseRename(temporary: string, target: string, code: string): void {
  renameMock.mockImplementation(async (from, to) => {
    if (String(from) === temporary && String(to) === target)
      throw Object.assign(new Error(`${code}: injected`), { code })
    return actualRename(from, to)
  })
}

/**
 * The first copy over `target` dies on its second write, after it has
 * truncated the workbook and written part of the new bytes. With
 * `refuseRestore`, the next copy over it (the restore) cannot open it at all.
 */
function inPlaceCopyDies(target: string, refuseRestore = false): void {
  let opens = 0
  openMock.mockImplementation(async (...args: Parameters<typeof open>) => {
    if (String(args[0]) !== target || args[1] !== 'w') return actualOpen(...args)
    opens += 1
    if (opens === 2 && refuseRestore) {
      throw Object.assign(new Error('EACCES: injected'), { code: 'EACCES' })
    }
    const handle = await actualOpen(...args)
    if (opens !== 1) return handle
    const write = handle.write.bind(handle) as (...a: unknown[]) => Promise<unknown>
    let writes = 0
    handle.write = (async (...a: unknown[]) => {
      writes += 1
      if (writes === 2) throw Object.assign(new Error('EIO: injected'), { code: 'EIO' })
      return write(...a)
    }) as typeof handle.write
    return handle
  })
}

/** More than one copy buffer, so the in-place copy writes at least twice. */
const largeNewBytes = Buffer.alloc(1536 * 1024, 7)

describe('promoteFileAtomically on a mount that refuses the platform copy', () => {
  it('lands the save in place and never removes the workbook', async () => {
    usePlatform('linux')
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    refuseRename(temporary, target, 'EPERM')
    platformCopyFailsLikeLibuvIn(dir)
    await promoteFileAtomically(temporary, target)
    expect(renameMock).toHaveBeenCalledTimes(5)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
    expect(unlinkMock.mock.calls.filter(([path]) => String(path) === target)).toEqual([])
    expect(copyFileMock.mock.calls.filter(([, dest]) => String(dest) === target)).toEqual([])
    expect(await readdir(dir)).toEqual(['book.xlsx'])
  }, 15_000)

  it('restores the workbook from the backup when the in-place copy dies part-way', async () => {
    usePlatform('linux')
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, largeNewBytes)
    await writeFile(target, 'old-bytes')
    refuseRename(temporary, target, 'EPERM')
    platformCopyFailsLikeLibuvIn(dir)
    inPlaceCopyDies(target)
    await expect(promoteFileAtomically(temporary, target)).rejects.toThrow('EIO')
    expect(await readFile(target, 'utf8')).toBe('old-bytes')
    expect(await readFile(temporary)).toEqual(largeNewBytes)
    expect(unlinkMock.mock.calls.filter(([path]) => String(path) === target)).toEqual([])
  }, 15_000)

  it('names the recovered copy beside the workbook when the restore fails too', async () => {
    usePlatform('linux')
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, largeNewBytes)
    await writeFile(target, 'old-bytes')
    refuseRename(temporary, target, 'EPERM')
    platformCopyFailsLikeLibuvIn(dir)
    inPlaceCopyDies(target, true)
    const failure = await promoteFileAtomically(temporary, target).catch((error: Error) => error)
    expect(failure?.message).toContain('preserved at: ')
    const survivor = failure!.message.split('preserved at: ')[1] ?? ''
    expect(dirname(survivor)).toBe(dir)
    expect(basename(survivor)).toMatch(/^book\.recovered-\d+\.xlsx$/)
    expect(await readFile(survivor, 'utf8')).toBe('old-bytes')
    expect(await readFile(temporary)).toEqual(largeNewBytes)
  }, 15_000)

  it('keeps a file that already holds the recovered name and reports the tmpdir backup', async () => {
    usePlatform('linux')
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, largeNewBytes)
    await writeFile(target, 'old-bytes')
    vi.spyOn(Date, 'now').mockReturnValue(1_700_000_000_000)
    const taken = join(dir, 'book.recovered-1700000000000.xlsx')
    await writeFile(taken, 'someone else')
    refuseRename(temporary, target, 'EPERM')
    platformCopyFailsLikeLibuvIn(dir)
    inPlaceCopyDies(target, true)
    const failure = await promoteFileAtomically(temporary, target).catch((error: Error) => error)
    const survivor = failure!.message.split('preserved at: ')[1] ?? ''
    expect(survivor.startsWith(tmpdir())).toBe(true)
    expect(await readFile(survivor, 'utf8')).toBe('old-bytes')
    expect(await readFile(taken, 'utf8')).toBe('someone else')
    await rm(survivor, { force: true })
  }, 15_000)

  // A FUSE mount may keep answering stat from a cache: the truncated workbook
  // then still looks untouched, and only the backup can bring it back.
  it('restores the workbook even when stat still reports its old size and time', async () => {
    usePlatform('linux')
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, largeNewBytes)
    await writeFile(target, 'old-bytes')
    const cached = await actualStat(target)
    statMock.mockImplementation((async (path: string, options?: unknown) =>
      String(path) === target ? cached : actualStat(path, options as never)) as typeof stat)
    refuseRename(temporary, target, 'EPERM')
    platformCopyFailsLikeLibuvIn(dir)
    inPlaceCopyDies(target)
    await expect(promoteFileAtomically(temporary, target)).rejects.toThrow('EIO')
    expect(await readFile(target, 'utf8')).toBe('old-bytes')
    expect(await readFile(temporary)).toEqual(largeNewBytes)
  }, 15_000)

  it('reports a workbook it cannot open for writing as locked', async () => {
    usePlatform('linux')
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    refuseRename(temporary, target, 'EPERM')
    platformCopyFailsLikeLibuvIn(dir)
    openMock.mockImplementation(async (...args: Parameters<typeof open>) => {
      if (String(args[0]) === target && args[1] === 'w') throw busy()
      return actualOpen(...args)
    })
    await expect(promoteFileAtomically(temporary, target)).rejects.toThrow(
      'The save target is locked by another program',
    )
    expect(await readFile(target, 'utf8')).toBe('old-bytes')
    expect(await readFile(temporary, 'utf8')).toBe('new-bytes')
  }, 15_000)
})
