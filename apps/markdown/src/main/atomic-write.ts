import { randomBytes } from 'node:crypto'
import { rename, unlink, writeFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

/** Transient Windows codes: antivirus/indexer briefly locks the rename target. */
const RETRYABLE_RENAME_CODES = new Set(['EPERM', 'EACCES', 'EBUSY'])
const RENAME_RETRIES = 4

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/**
 * Same-dir temp file + rename, so a crash mid-write can't truncate the target.
 * Rename-over-existing fails transiently on Windows under Defender/indexer
 * locks: retry with backoff, then fall back to an in-place write.
 * Mirrors apps/docs/src/main/atomic-write.ts (candidate for electron-utils).
 */
export async function atomicWriteFile(filePath: string, data: Buffer): Promise<void> {
  const tmp = join(
    dirname(filePath),
    `.${basename(filePath)}.${randomBytes(6).toString('hex')}.tmp`,
  )
  try {
    await writeFile(tmp, data)
    for (let attempt = 0; ; attempt++) {
      try {
        await rename(tmp, filePath)
        return
      } catch (err) {
        const code = (err as NodeJS.ErrnoException).code ?? ''
        if (!RETRYABLE_RENAME_CODES.has(code) || attempt >= RENAME_RETRIES) throw err
        await sleep(50 * 2 ** attempt)
      }
    }
  } catch (err) {
    try {
      await unlink(tmp)
    } catch {
      /* tmp never created */
    }
    if (RETRYABLE_RENAME_CODES.has((err as NodeJS.ErrnoException).code ?? '')) {
      await writeFile(filePath, data)
      return
    }
    throw err
  }
}
