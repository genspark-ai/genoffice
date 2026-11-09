/// The largest .csv/.tsv Sheets converts on open (genoffice#2505). The
/// conversion streams rows into the zip archive and reports progress, so
/// files well past the old hardcoded 32 MiB cap import safely; the cap
/// below is the memory guard for the decoded text (held whole), not for
/// the workbook. Stored under `csvImportMaxBytes` (bytes) in
/// userData/app-settings.json, read fresh on every open so raising it
/// takes effect without a restart. Keep DEFAULT_CSV_IMPORT_MAX_BYTES in
/// sync with the renderer-side copy in xlsx-gateway's csv-import, which
/// shares the same default for Data → From Text/CSV.
import { readFileSync } from 'node:fs'

import { writeJsonAtomic } from './atomic-write'

export const CSV_IMPORT_MAX_BYTES_KEY = 'csvImportMaxBytes'

/** default cap: 128 MiB, four times the old hardcoded ceiling */
export const DEFAULT_CSV_IMPORT_MAX_BYTES = 128 * 1024 * 1024

/** never let a typo'd setting turn into a 1-byte cap or an unbounded read */
export const MIN_CSV_IMPORT_MAX_BYTES = 1 * 1024 * 1024

/** hard ceiling: decoding ~512 MiB of text alone can OOM low-memory machines */
export const MAX_CSV_IMPORT_MAX_BYTES = 512 * 1024 * 1024

/** clamp an arbitrary value to a usable cap; anything unusable → the default */
export function sanitizeCsvImportMaxBytes(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return DEFAULT_CSV_IMPORT_MAX_BYTES
  return Math.min(MAX_CSV_IMPORT_MAX_BYTES, Math.max(MIN_CSV_IMPORT_MAX_BYTES, Math.trunc(value)))
}

function readSettings(settingsPath: string): Record<string, unknown> {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) return raw as Record<string, unknown>
  } catch {
    /* missing or corrupt: fall through to the default */
  }
  return {}
}

/** the configured CSV import cap from app-settings.json, or the default */
export function readCsvImportMaxBytes(settingsPath: string): number {
  return sanitizeCsvImportMaxBytes(readSettings(settingsPath)[CSV_IMPORT_MAX_BYTES_KEY])
}

/** persist a cap (validated/clamped) alongside the other app settings */
export function writeCsvImportMaxBytes(settingsPath: string, bytes: number): number {
  const settings = readSettings(settingsPath)
  const sanitized = sanitizeCsvImportMaxBytes(bytes)
  writeJsonAtomic(settingsPath, { ...settings, [CSV_IMPORT_MAX_BYTES_KEY]: sanitized })
  return sanitized
}

/** human-readable size for rejection messages: whole MB, e.g. "128 MB" */
export function formatCsvImportBytes(bytes: number): string {
  return `${Math.round(bytes / (1024 * 1024))} MB`
}

/** true when a file of `fileSize` bytes may be converted; exactly at the cap is allowed */
export function csvImportAllowed(fileSize: number, maxBytes = DEFAULT_CSV_IMPORT_MAX_BYTES): boolean {
  return fileSize <= maxBytes
}
