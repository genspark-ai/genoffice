import { randomBytes } from 'node:crypto'
import {
  closeSync,
  chmodSync,
  existsSync,
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
const TOLERATED_SYNC_CODES = new Set(['EPERM', 'EACCES', 'EBUSY', 'EINVAL', 'ENOSYS', 'ENOTSUP'])

const errorCode = (error: unknown) => (error as NodeJS.ErrnoException).code ?? ''
const isRetryableRename = (error: unknown) => RETRYABLE_RENAME_CODES.has(errorCode(error))
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))
const sleepSync = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms)

// No leading dot: the macOS SMB client marks dot-files hidden and the flag
// survives the rename onto the real document (genoffice#2145).
const tempPathBeside = (filePath: string) =>
  join(dirname(filePath), `${basename(filePath)}.${randomBytes(6).toString('hex')}.tmp`)

/**
 * Where a failed save parks the new bytes: beside the target, under ONE
 * stable name.
 *
 * The staging temp carries a random suffix — two writers of the same target
 * must not share it — but the rescue copy must not: a save that keeps
 * failing (the document is open in another program, and the autosave retries
 * it every 30 s) would otherwise leave a new file beside the document on
 * every attempt, and the user's folder fills with them (genoffice#2567).
 */
const rescuePathFor = (filePath: string) =>
  join(dirname(filePath), `${basename(filePath)}.unsaved.tmp`)

/**
 * Park the new bytes beside the target, replacing the copy an earlier failed
 * attempt left. copyFile (not rename) because the destination may already
 * exist — the previous attempt's copy — and a rename onto an existing file is
 * refused on Windows. It stays in the kernel, so a multi-hundred-MB file is
 * never read into RAM.
 *
 * If the copy does not land, the staging temp still holds the bytes: they
 * are on disk either way.
 */
async function keepRescueCopy(filePath: string, tmp: string): Promise<void> {
  const rescue = rescuePathFor(filePath)
  try {
    await copyFile(tmp, rescue)
    // The staging temp is already on disk; the copy must be too before it
    // becomes the only holder of the new bytes.
    await syncFileBestEffort(rescue)
  } catch {
    /* the staging temp still holds the bytes */
    return
  }
  await unlink(tmp).catch(() => {})
}

/** The save landed: an earlier attempt's rescue copy is stale. */
async function clearRescueCopy(filePath: string): Promise<void> {
  await unlink(rescuePathFor(filePath)).catch(() => {})
}

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
    await clearRescueCopy(filePath)
    return
  } catch {
    // fall through to the in-place write
  }
  // Keep the completed temp until the fallback lands: if that write fails
  // or the process dies, the new bytes still exist somewhere on disk.
  try {
    await writeFile(filePath, data)
  } catch (error) {
    // The target could not be written at all: park the new bytes beside it
    // under one stable name, so a save that keeps failing replaces the same
    // file instead of leaving a new one on every attempt (genoffice#2567).
    await keepRescueCopy(filePath, tmp)
    throw error
  }
  await unlink(tmp).catch(() => {})
  await clearRescueCopy(filePath)
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
    await clearRescueCopy(filePath)
    return
  } catch {
    // fall through to the in-place write
  }
  // Keep the completed temp until the fallback lands, as atomicWriteFile does.
  try {
    await copyFile(source, filePath)
  } catch (error) {
    await keepRescueCopy(filePath, tmp)
    throw error
  }
  await unlink(tmp).catch(() => {})
  await clearRescueCopy(filePath)
}

/**
 * The same contract for output that is produced as a stream rather than held
 * in one Buffer: `write` fills the sibling temp path (a large deck streams out
 * of JSZip chunk by chunk and never fits atomicWriteFile's single buffer), and
 * only a temp that `write` completed is synced and renamed over the target. A
 * writer that throws part-way leaves the target untouched, so a save that dies
 * mid-stream keeps the previous file on disk instead of a truncated one
 * (genoffice#2123).
 */
export async function atomicWriteVia(
  filePath: string,
  write: (tempPath: string) => Promise<void>,
): Promise<void> {
  const tmp = tempPathBeside(filePath)
  try {
    await write(tmp)
    await syncFileBestEffort(tmp)
  } catch (error) {
    await unlink(tmp).catch(() => {})
    throw error
  }
  try {
    await renameWithRetry(tmp, filePath)
    await clearRescueCopy(filePath)
    return
  } catch {
    // fall through to the in-place copy
  }
  // Keep the completed temp until the fallback lands, as atomicWriteFile does.
  try {
    await copyFile(tmp, filePath)
  } catch (error) {
    await keepRescueCopy(filePath, tmp)
    throw error
  }
  await unlink(tmp).catch(() => {})
  await clearRescueCopy(filePath)
}

/**
 * Atomic replacement for small userData JSON state files (recents, stars, AI
 * settings). Same retry + in-place fallback as atomicWriteFile; the retry
 * sleeps block the caller, which is acceptable for a few KB under a transient
 * Windows lock.
 *
 * `mode` restricts the file's permissions (e.g. 0o600 for the AI settings
 * file, whose BYOK keys must not be readable by other local accounts): it
 * applies at creation so the rename lands the restricted file, and the
 * in-place fallback passes it too. Ignored on platforms without POSIX modes.
 */
export function writeJsonAtomic(filePath: string, value: unknown, mode?: number): void {
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
    writeFileSync(tmp, json, { encoding: 'utf8', flag: 'wx', ...(mode ? { mode } : {}) })
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
      writeFileSync(filePath, json, { encoding: 'utf8', ...(mode ? { mode } : {}) })
      removeTemp()
      return
    }
    removeTemp()
    throw error
  }
}

/**
 * Best-effort owner-only permissions for a credentials file that may predate
 * the 0600 write (a file saved by an older build stays world-readable until
 * the next save). Called from the settings readers, which run at every app
 * start, so one launch heals the mode. No-op on platforms without POSIX modes.
 */
export function restrictFileToOwner(filePath: string): void {
  try {
    if (existsSync(filePath)) chmodSync(filePath, 0o600)
  } catch {
    /* a locked or unusual filesystem must not break the read that follows */
  }
}
