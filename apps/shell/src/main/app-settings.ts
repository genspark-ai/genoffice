import { randomUUID } from 'node:crypto'
import { readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'

/**
 * Helpers for userData/app-settings.json — a flat JSON object holding
 * cross-module app preferences (UI language, first-run onboarding flag).
 * Editor modules read the same file, so the shape must stay a plain object.
 */

export function readAppSettings(path: string): Record<string, unknown> {
  try {
    const raw: unknown = JSON.parse(readFileSync(path, 'utf8'))
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
      return raw as Record<string, unknown>
    }
  } catch {
    // missing or corrupt file: treat as empty settings
  }
  return {}
}

/**
 * Persist several related settings as one atomic replacement. The temporary
 * file lives beside the destination so rename never crosses filesystems.
 */
export function writeAppSettings(path: string, updates: Record<string, unknown>): void {
  const settings = { ...readAppSettings(path), ...updates }
  const tempPath = `${path}.${process.pid}.${randomUUID()}.tmp`
  try {
    writeFileSync(tempPath, JSON.stringify(settings, null, 2), {
      encoding: 'utf8',
      flag: 'wx',
      flush: true,
    })
    renameSync(tempPath, path)
  } catch (error) {
    try {
      unlinkSync(tempPath)
    } catch {
      // The write may have failed before the temporary file was created.
    }
    throw error
  }
}

export function writeAppSetting(path: string, key: string, value: unknown): void {
  writeAppSettings(path, { [key]: value })
}
