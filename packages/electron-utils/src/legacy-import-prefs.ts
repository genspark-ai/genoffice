/// "Convert without asking next time" on the legacy-format import dialogs
/// (.doc/.odt/.hwp in Docs, .ods in Sheets, .odp in Slides). Stored per
/// format under `legacyImportSkipConfirm` in userData/app-settings.json; the
/// shell's Settings → General switch clears or sets every format at once.
/// Only the confirmation is skipped: the Save As step that picks the converted
/// copy's destination always runs.
import { readFileSync } from 'node:fs'
import { writeJsonAtomic } from './atomic-write'

export const LEGACY_IMPORT_SKIP_KEY = 'legacyImportSkipConfirm'

export const LEGACY_IMPORT_KINDS = ['word97', 'odt', 'hwp', 'ods', 'odp'] as const
export type LegacyImportKind = (typeof LEGACY_IMPORT_KINDS)[number]

function readSettings(settingsPath: string): Record<string, unknown> {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) return raw as Record<string, unknown>
  } catch {
    /* missing or corrupt: no preferences */
  }
  return {}
}

function skipMap(settings: Record<string, unknown>): Partial<Record<LegacyImportKind, boolean>> {
  const raw = settings[LEGACY_IMPORT_SKIP_KEY]
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {}
  const out: Partial<Record<LegacyImportKind, boolean>> = {}
  for (const kind of LEGACY_IMPORT_KINDS) {
    if ((raw as Record<string, unknown>)[kind] === true) out[kind] = true
  }
  return out
}

export function readLegacyImportSkip(settingsPath: string, kind: LegacyImportKind): boolean {
  return skipMap(readSettings(settingsPath))[kind] === true
}

export function writeLegacyImportSkip(
  settingsPath: string,
  kind: LegacyImportKind,
  skip: boolean,
): void {
  const settings = readSettings(settingsPath)
  const map = skipMap(settings)
  if (skip) map[kind] = true
  else delete map[kind]
  writeJsonAtomic(settingsPath, { ...settings, [LEGACY_IMPORT_SKIP_KEY]: map })
}

/** true while no format is set to skip, i.e. every legacy import still asks */
export function readLegacyImportPrompt(settingsPath: string): boolean {
  return Object.keys(skipMap(readSettings(settingsPath))).length === 0
}

/** on: clear every skip flag; off: skip the confirmation for every format */
export function writeLegacyImportPrompt(settingsPath: string, prompt: boolean): void {
  const settings = readSettings(settingsPath)
  const map: Partial<Record<LegacyImportKind, boolean>> = {}
  if (!prompt) for (const kind of LEGACY_IMPORT_KINDS) map[kind] = true
  writeJsonAtomic(settingsPath, { ...settings, [LEGACY_IMPORT_SKIP_KEY]: map })
}
