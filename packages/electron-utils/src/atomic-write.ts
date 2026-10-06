import { randomBytes } from 'node:crypto'
import {
  closeSync,
  fsyncSync,
  mkdirSync,
  openSync,
  renameSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { copyFile, open, rename, unlink, writeFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

/** Transient Windows codes: antivirus/indexer briefly locks the rename target. */
const RETRYABLE_RENAME_CODES = new Set(['EPERM', 'EACCES', 'EBUSY'])
const RENAME_RETRY_DELAYS_MS = [50, 100, 200, 400]

/**
 * Codes a refused fsync may surface with inside cloud-sync folders, under AV
 * locks or on filesystems without fsync. The bytes are already written, so a
 * refused flush only weakens power-loss durability and must not fail the save.
 */
const TOLERATED_SYNC_CODES = new Set(['EPERM', 'EACCES', 'EBUSY', 'EINVAL', 'ENOSYS'])

const errorCode = (error: unknown) => (error as NodeJS.ErrnoException).code ?? ''
const isRetryableRename = (error: unknown) => RETRYABLE_RENAME_CODES.has(errorCode(error))
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))
const sleepSync = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms)

const tempPathBeside = (filePath: string) =>
  join(dirname(filePath), `.${basename(filePath)}.${randomBytes(6).toString('hex')}.tmp`)

async function syncFileBestEffort(path: string): Promise<void> {
  let handle
  try {
    handle = await open(path, 'r+')
  } catch (error) {
    if (TOLERATED_SYNC_CODES.has(errorCode(error))) return
    throw error
  }
  try {
    await handle.sync()
  } catch (error) {
    if (!TOLERATED_SYNC_CODES.has(errorCode(error))) throw error
  } finally {
    await handle.close()
  }
}

function syncFileBestEffortSync(path: string): void {
  let descriptor: number
  try {
    descriptor = openSync(path, 'r+')
  } catch (error) {
    if (TOLERATED_SYNC_CODES.has(errorCode(error))) return
    throw error
  }
  try {
    fsyncSync(descriptor)
  } catch (error) {
    if (!TOLERATED_SYNC_CODES.has(errorCode(error))) throw error
  } finally {
    closeSync(descriptor)
  }
}

/**
 * Same-dir temp file, best-effort fsync, then rename, so neither a crash
 * mid-write nor a power loss right after the rename can leave the target
 * truncated. Rename-over-existing fails transiently on Windows under
 * Defender/indexer locks (retry with backoff), and permanently on network
 * mounts — a gvfs SMB share refuses it with EEXIST, outside the retry set,
 * which made every Docs/Sheets save fail where a plain write succeeded
 * (#1877). Either way the save publishes through an in-place fallback write:
 * losing atomicity for that one save beats failing a save the previous plain
 * writeFileSync would have completed.
 */
export async function atomicWriteFile(filePath: string, data: Buffer): Promise<void> {
  const tmp = tempPathBeside(filePath)
  try {
    await writeFile(tmp, data)
    await syncFileBestEffort(tmp)
  } catch (error) {
    await unlink(tmp).catch(() => {})
    throw error
  }
  try {
    await renameWithRetry(tmp, filePath)
    return
  } catch {
    // any publish failure lands on the in-place fallback below
  }
  // Keep the completed temp until the fallback lands: if that write fails
  // or the process dies, the new bytes still exist somewhere on disk.
  await writeFile(filePath, data)
  await unlink(tmp).catch(() => {})
}

/** Publish a completed temp file over the target, retrying the transient
 *  Windows rename locks; any other failure propagates to the caller, whose
 *  in-place fallback publishes the save without the atomic replace. */
async function renameWithRetry(tmp: string, filePath: string): Promise<void> {
  for (let attempt = 0; ; attempt += 1) {
    try {
      await rename(tmp, filePath)
      return
    } catch (error) {
      if (!isRetryableRename(error) || attempt >= RENAME_RETRY_DELAYS_MS.length) throw error
      await sleep(RENAME_RETRY_DELAYS_MS[attempt] ?? 0)
    }
  }
}

/**
 * Same temp-file, best-effort fsync and rename sequence as atomicWriteFile, for
 * a copy of an existing file. The copy itself stays in the kernel, so a
 * multi-hundred-MB PDF is never read into memory; only the flush and the
 * publish are shared, so the shell does not need a second implementation.
 * A refused publish — transient Windows lock or a network mount that does not
 * rename over existing files — falls back to an in-place copy, like
 * atomicWriteFile.
 */
export async function atomicCopyFile(source: string, filePath: string): Promise<void> {
  const tmp = tempPathBeside(filePath)
  try {
    await copyFile(source, tmp)
    await syncFileBestEffort(tmp)
  } catch (error) {
    await unlink(tmp).catch(() => {})
    throw error
  }
  try {
    await renameWithRetry(tmp, filePath)
    return
  } catch {
    // any publish failure lands on the in-place fallback below
  }
  // Keep the completed temp until the fallback lands, as atomicWriteFile does.
  await copyFile(source, filePath)
  await unlink(tmp).catch(() => {})
}

/**
 * Atomic replacement for small userData JSON state files (recents, stars, AI
 * settings). Same retry + in-place fallback as atomicWriteFile; the retry
 * sleeps block the caller, which is acceptable for a few KB under a transient
 * Windows lock.
 */
export function writeJsonAtomic(filePath: string, value: unknown): void {
  mkdirSync(dirname(filePath), { recursive: true })
  const json = JSON.stringify(value, null, 2)
  const tmp = tempPathBeside(filePath)
  const removeTemp = () => {
    try {
      unlinkSync(tmp)
    } catch {
      /* temp never created */
    }
  }
  try {
    writeFileSync(tmp, json, { encoding: 'utf8', flag: 'wx' })
    syncFileBestEffortSync(tmp)
    for (let attempt = 0; ; attempt += 1) {
      try {
        renameSync(tmp, filePath)
        return
      } catch (error) {
        if (!isRetryableRename(error) || attempt >= RENAME_RETRY_DELAYS_MS.length) throw error
        sleepSync(RENAME_RETRY_DELAYS_MS[attempt] ?? 0)
      }
    }
  } catch (error) {
    if (isRetryableRename(error)) {
      writeFileSync(filePath, json, 'utf8')
      removeTemp()
      return
    }
    removeTemp()
    throw error
  }
}
