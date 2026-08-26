import { randomBytes } from 'node:crypto'
import { rename, unlink, writeFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

/** Transient Windows codes: antivirus/indexer briefly locks the rename target. */
const RETRYABLE_RENAME_CODES = new Set(['EPERM', 'EACCES', 'EBUSY'])
const RENAME_RETRIES = 4

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

/**
 * Same-directory temp file + rename prevents observers from reading a partial
 * export. Windows rename-over-existing can fail transiently, so retry with
 * backoff and finally fall back to the prior in-place write behavior.
 */
export async function atomicWriteFile(filePath: string, data: Buffer): Promise<void> {
  const tmp = join(
    dirname(filePath),
    `.${basename(filePath)}.${randomBytes(6).toString('hex')}.tmp`,
  )
  try {
    await writeFile(tmp, data)
    for (let attempt = 0; ; attempt += 1) {
      try {
        await rename(tmp, filePath)
        return
      } catch (error) {
        const code = (error as NodeJS.ErrnoException).code ?? ''
        if (!RETRYABLE_RENAME_CODES.has(code) || attempt >= RENAME_RETRIES) throw error
        await sleep(50 * 2 ** attempt)
      }
    }
  } catch (error) {
    const retryable = RETRYABLE_RENAME_CODES.has((error as NodeJS.ErrnoException).code ?? '')
    if (retryable) {
      // Preserve the completed temp until the non-atomic fallback succeeds.
      // If that write fails or the process exits, the new bytes still exist.
      await writeFile(filePath, data)
      await unlink(tmp).catch(() => {})
      return
    }
    await unlink(tmp).catch(() => {})
    throw error
  }
}
