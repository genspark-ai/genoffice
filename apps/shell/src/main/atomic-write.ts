import { randomBytes } from 'node:crypto'
import { open, rename, unlink, writeFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

export { atomicCopyFile } from '@genoffice/electron-utils/atomic-write'

const RETRYABLE_RENAME_CODES = new Set(['EPERM', 'EACCES', 'EBUSY'])
// cloud-sync folders, AV locks and fsync-less filesystems refuse the flush after a
// successful write; that only weakens power-loss durability, so the save proceeds
const TOLERATED_SYNC_CODES = new Set(['EPERM', 'EACCES', 'EBUSY', 'EINVAL', 'ENOSYS'])
const RENAME_RETRIES = 4
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

/// Not the electron-utils helper: that one falls back to an in-place write, this
/// one must leave the destination untouched on failure (export-atomicity.test.ts).
export function atomicWriteFile(filePath: string, data: Uint8Array): Promise<void> {
  return viaTemp(filePath, (tmp) => writeFile(tmp, data))
}

async function viaTemp(filePath: string, fill: (tmp: string) => Promise<void>): Promise<void> {
  const tmp = join(
    dirname(filePath),
    `.${basename(filePath)}.${randomBytes(6).toString('hex')}.tmp`,
  )
  try {
    await fill(tmp)
    await fsyncPath(tmp)
    for (let attempt = 0; ; attempt += 1) {
      try {
        await rename(tmp, filePath)
        // the rename is only durable once the directory entry is flushed; opening
        // a directory fails on Windows, where NTFS journals the rename itself
        if (process.platform !== 'win32') await fsyncPath(dirname(filePath), true)
        return
      } catch (error) {
        const code = (error as NodeJS.ErrnoException).code ?? ''
        if (!RETRYABLE_RENAME_CODES.has(code) || attempt >= RENAME_RETRIES) throw error
        await sleep(50 * 2 ** attempt)
      }
    }
  } catch (error) {
    await unlink(tmp).catch(() => {})
    throw error
  }
}

async function fsyncPath(path: string, tolerateAll = false): Promise<void> {
  const tolerated = (error: unknown) =>
    tolerateAll || TOLERATED_SYNC_CODES.has((error as NodeJS.ErrnoException).code ?? '')
  let handle
  try {
    handle = await open(path, 'r')
  } catch (error) {
    if (tolerated(error)) return
    throw error
  }
  try {
    await handle.sync()
  } catch (error) {
    if (!tolerated(error)) throw error
  } finally {
    await handle.close()
  }
}
