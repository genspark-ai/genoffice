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
 * Same-dir temp file, best-effort fsync, then rename. Rename-over-existing fails
 * transiently on Windows (Defender/indexer locks: retry) and permanently on some
 * network mounts (gvfs SMB refuses with EEXIST, genoffice#1877); either way the
 * save falls back to an in-place write rather than failing.
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
    // fall through to the in-place write
  }
  // Keep the completed temp until the fallback lands: if that write fails
  // or the process dies, the new bytes still exist somewhere on disk.
  await writeFile(filePath, data)
  await unlink(tmp).catch(() => {})
}

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

/** copyFile stays in the kernel, so a multi-hundred-MB file is never read into RAM */
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
    // fall through to the in-place write
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
