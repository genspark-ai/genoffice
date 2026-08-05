import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync, chmodSync } from 'node:fs'
import { dirname } from 'node:path'

/** Parse a JSON file without allowing a corrupt settings file to crash startup. */
export function readJsonFile<T>(filePath: string): T | undefined {
  try {
    if (!existsSync(filePath)) return undefined
    return JSON.parse(readFileSync(filePath, 'utf8')) as T
  } catch {
    return undefined
  }
}

/**
 * Write a JSON document with a same-directory temporary file followed by rename.
 * A process interruption therefore leaves either the old document or the new one.
 */
export function writeAtomicJson(filePath: string, value: unknown): void {
  mkdirSync(dirname(filePath), { recursive: true })
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`
  try {
    writeFileSync(tempPath, JSON.stringify(value, null, 2) + '\n', {
      encoding: 'utf8',
      mode: 0o600,
    })
    chmodSync(tempPath, 0o600)
    renameSync(tempPath, filePath)
    chmodSync(filePath, 0o600)
  } catch (error) {
    try {
      // The temporary file is uniquely named and can only contain this write.
      // unlinkSync is intentionally avoided here so a failed cleanup cannot mask
      // the original error; a later startup may safely ignore the .tmp file.
      chmodSync(tempPath, 0o600)
    } catch {
      // Best-effort permissions on a failed write.
    }
    throw error
  }
}
