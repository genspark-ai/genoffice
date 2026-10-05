import { randomBytes } from 'node:crypto'
import { rename, unlink, writeFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

/// The shell's copy path was a second implementation of the same temp-file and
/// rename dance, and the only atomic write in the tree that did not flush
/// before publishing - a power loss right after the rename could leave a copied
/// PDF or docx truncated. atomicCopyFile now lives beside atomicWriteFile in
/// electron-utils, which already fsyncs the temp file and deliberately tolerates
/// a refused flush (EPERM/EINVAL/ENOSYS) on cloud-sync and AV-locked folders.
/// The copy itself stays in the kernel there, so a multi-hundred-MB file is
/// still never read into RAM.
export { atomicCopyFile } from '@genoffice/electron-utils/atomic-write'

const RETRYABLE_RENAME_CODES = new Set(['EPERM', 'EACCES', 'EBUSY'])
const RENAME_RETRIES = 4
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

/// Deliberately not the shared electron-utils helper: that one falls back to an
/// in-place destination write when the publish keeps failing, whereas this one
/// propagates the error and leaves the destination untouched, which
/// apps/shell/tests/export-atomicity.test.ts pins.
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
    await unlink(tmp).catch(() => {})
    throw error
  }
}
