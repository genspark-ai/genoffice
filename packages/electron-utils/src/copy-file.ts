/**
 * File copies that still work on mounts where the platform copy does not.
 *
 * Node's copyFile sets the destination's mode (fchmod) before it copies the
 * data. A mount without permission support — GNOME gvfs SMB and WebDAV, for
 * one — refuses that call, so the copy fails although plain reads and writes
 * work there. On POSIX the failed copy also unlinks the destination, so a file
 * that was already there is gone afterwards.
 *
 * Windows copies through CopyFileW, which has neither problem. There the
 * platform copy stays the only step.
 */
import {
  closeSync,
  constants,
  copyFileSync,
  fstatSync,
  openSync,
  readSync,
  unlinkSync,
  writeSync,
} from 'node:fs'
import { copyFile, open, realpath, stat, unlink } from 'node:fs/promises'

/** One buffer of this size carries a whole copy, so a multi-hundred-MB file is never read into RAM. */
const COPY_CHUNK_BYTES = 1024 * 1024

const errorCode = (error: unknown) => (error as NodeJS.ErrnoException).code ?? ''

/** Asked on every call rather than once at import, so both branches can run in one process. */
const platformCopyOnly = () => process.platform === 'win32'

const stalledWrite = () =>
  Object.assign(new Error('EIO: the copy stopped making progress'), { code: 'EIO' })

/**
 * A pipe or a device has no end to read up to, and a folder cannot be read at
 * all: only a regular file is copied by reading it.
 */
const notRegularFile = (isDirectory: boolean) =>
  isDirectory
    ? Object.assign(new Error('EISDIR: illegal operation on a directory, copy'), { code: 'EISDIR' })
    : Object.assign(new Error('EINVAL: not a regular file, copy'), { code: 'EINVAL' })

/**
 * Copy `source` to `target`, which must not exist yet. An existing `target` is
 * left untouched and the call rejects with `EEXIST`.
 *
 * The platform copy runs first. Outside Windows, when it fails for any other
 * reason the bytes are read and written once more through plain file handles,
 * and that attempt's error is the one reported.
 *
 * `EEXIST` therefore always means the file was there before the call. When
 * only the second attempt finds the name taken (the failed platform copy
 * could not remove what it had created, or someone created the file in
 * between), the first error is reported instead, so a caller that tries the
 * next name on `EEXIST` does not walk on leaving a file behind each time.
 */
export async function copyFileExclusive(source: string, target: string): Promise<void> {
  let platformError: unknown
  try {
    await copyFile(source, target, constants.COPYFILE_EXCL)
    return
  } catch (error) {
    if (platformCopyOnly() || errorCode(error) === 'EEXIST') throw error
    platformError = error
  }
  try {
    await writeCopy(source, target, 'wx')
  } catch (error) {
    throw errorCode(error) === 'EEXIST' ? platformError : error
  }
}

/** copyFileExclusive for callers that cannot await: it blocks until the copy is done. */
export function copyFileExclusiveSync(source: string, target: string): void {
  let platformError: unknown
  try {
    copyFileSync(source, target, constants.COPYFILE_EXCL)
    return
  } catch (error) {
    if (platformCopyOnly() || errorCode(error) === 'EEXIST') throw error
    platformError = error
  }
  try {
    writeNewCopySync(source, target)
  } catch (error) {
    throw errorCode(error) === 'EEXIST' ? platformError : error
  }
}

/**
 * Replace the contents of `target` (created when missing) with those of
 * `source`, in place. Not atomic: a failure part-way leaves `target`
 * incomplete, so the caller must still hold the new bytes in another file. It
 * is the step after a refused rename.
 *
 * Outside Windows the platform copy is never tried here: by the time it fails
 * it has truncated `target`, and it then unlinks it.
 */
export async function copyFileInPlace(source: string, target: string): Promise<void> {
  if (platformCopyOnly()) {
    await copyFile(source, target)
    return
  }
  // Opening a file for writing truncates it, which would empty the source as well.
  if (await isSameFile(source, target)) return
  await writeCopy(source, target, 'w')
}

/**
 * Leave a copy of `source` under `target`, a name the caller owns (a rescue
 * copy). Whatever holds that name is replaced, not written into: in a shared
 * folder it may be a link, or a file someone else put there, and neither may
 * receive the bytes.
 *
 * Windows keeps the single platform copy over the earlier file.
 */
export async function copyFileFresh(source: string, target: string): Promise<void> {
  if (platformCopyOnly()) {
    await copyFile(source, target)
    return
  }
  await unlink(target).catch((error) => {
    if (errorCode(error) !== 'ENOENT') throw error
  })
  await copyFileExclusive(source, target)
}

/**
 * True only when both paths are known to name one file. A mount that reports
 * no file number (0) cannot tell, and neither can a path that does not resolve
 * (the target may not exist yet): both count as different files, because
 * skipping the write would lose the new bytes. For the same reason a shared
 * file number alone is not enough: one file also has one size and one
 * modification time, and a mount that hands two files the same number rarely
 * agrees on those too.
 */
async function isSameFile(source: string, target: string): Promise<boolean> {
  try {
    if ((await realpath(source)) === (await realpath(target))) return true
  } catch {
    // compared by file number below
  }
  try {
    const [from, to] = await Promise.all([
      stat(source, { bigint: true }),
      stat(target, { bigint: true }),
    ])
    return (
      from.ino !== 0n &&
      from.dev !== 0n &&
      from.ino === to.ino &&
      from.dev === to.dev &&
      from.size === to.size &&
      from.mtimeNs === to.mtimeNs
    )
  } catch {
    return false
  }
}

/**
 * Read `source` and write it to `target` front to back. `'wx'` creates the
 * target and refuses an existing one; `'w'` truncates or creates it. `'r+'`
 * and `'a'` are not used: gvfs WebDAV refuses both.
 */
async function writeCopy(source: string, target: string, flags: 'w' | 'wx'): Promise<void> {
  const input = await open(source, 'r')
  try {
    const stats = await input.stat()
    if (!stats.isFile()) throw notRegularFile(stats.isDirectory())
    // The mode goes in at creation; a separate chmod is what these mounts refuse.
    const output = await open(target, flags, stats.mode & 0o777)
    let closed = false
    try {
      const buffer = Buffer.allocUnsafe(COPY_CHUNK_BYTES)
      for (;;) {
        const { bytesRead } = await input.read(buffer, 0, buffer.length, null)
        if (bytesRead === 0) break
        let written = 0
        while (written < bytesRead) {
          const { bytesWritten } = await output.write(buffer, written, bytesRead - written, null)
          if (bytesWritten === 0) throw stalledWrite()
          written += bytesWritten
        }
      }
      closed = true
      // gvfs hands the bytes to the server on close, so a failed close is a failed copy.
      await output.close()
    } catch (error) {
      if (!closed) await output.close().catch(() => {})
      // Only a file this call created is removed; `'w'` may hold someone's document.
      if (flags === 'wx') await unlink(target).catch(() => {})
      throw error
    }
  } finally {
    await input.close().catch(() => {})
  }
}

/** writeCopy with `'wx'`, blocking. */
function writeNewCopySync(source: string, target: string): void {
  const input = openSync(source, 'r')
  try {
    const stats = fstatSync(input)
    if (!stats.isFile()) throw notRegularFile(stats.isDirectory())
    const output = openSync(target, 'wx', stats.mode & 0o777)
    let closed = false
    try {
      const buffer = Buffer.allocUnsafe(COPY_CHUNK_BYTES)
      for (;;) {
        const bytesRead = readSync(input, buffer, 0, buffer.length, null)
        if (bytesRead === 0) break
        let written = 0
        while (written < bytesRead) {
          const bytesWritten = writeSync(output, buffer, written, bytesRead - written, null)
          if (bytesWritten === 0) throw stalledWrite()
          written += bytesWritten
        }
      }
      closed = true
      closeSync(output)
    } catch (error) {
      if (!closed) {
        try {
          closeSync(output)
        } catch {
          // the copy already failed
        }
      }
      try {
        unlinkSync(target)
      } catch {
        // the copy already failed
      }
      throw error
    }
  } finally {
    try {
      closeSync(input)
    } catch {
      // nothing was written through this descriptor
    }
  }
}
