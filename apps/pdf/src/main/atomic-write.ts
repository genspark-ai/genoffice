import { rename, rm, writeFile } from 'node:fs/promises'

/**
 * Write PDF bytes beside the destination, then atomically replace it.
 * Any failed write or rename removes the partially written temp file.
 */
export async function writePdfAtomically(targetPath: string, bytes: Uint8Array): Promise<void> {
  const tmp = `${targetPath}.gensave-${process.pid}.tmp`
  try {
    await writeFile(tmp, bytes)
    await rename(tmp, targetPath)
  } catch (err) {
    await rm(tmp, { force: true })
    throw err
  }
}
