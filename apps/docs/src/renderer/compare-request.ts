/** What the Compare Documents / Combine Documents dialog asks for. */
import type { CompareGranularity, ParsedDocFull } from '@genoffice/docx-engine'

/** one side of a comparison: the document open in this tab, or a picked file */
export type CompareSide =
  | { kind: 'current' }
  | { kind: 'file'; name: string; parsed: ParsedDocFull; bytes: Uint8Array; author: string }

export const COMPARE_SETTING_KEYS = [
  'moves',
  'comments',
  'formatting',
  'caseChanges',
  'whiteSpace',
  'tables',
  'headersFooters',
  'notes',
  'textBoxes',
  'fields',
] as const

export type CompareSettingKey = (typeof COMPARE_SETTING_KEYS)[number]
export type CompareDialogSettings = Record<CompareSettingKey, boolean>

/** where the changes are shown: written into the open document, or opened as a new one */
export type CompareTarget = 'original' | 'revised' | 'new'

export interface CompareRequest {
  mode: 'compare' | 'combine'
  original: CompareSide
  revised: CompareSide
  /** the name every difference is labelled with */
  author: string
  /** combine: the name the original's tracked changes take when they carry none */
  originalAuthor?: string
  settings: CompareDialogSettings
  granularity: CompareGranularity
  target: CompareTarget
}

export const ALL_COMPARE_SETTINGS: CompareDialogSettings = {
  moves: true,
  comments: true,
  formatting: true,
  caseChanges: true,
  whiteSpace: true,
  tables: true,
  headersFooters: true,
  notes: true,
  textBoxes: true,
  fields: true,
}

/** written into the open document, only its body is compared: these need the new-document comparison */
const NEW_DOCUMENT_ONLY: ReadonlySet<CompareSettingKey> = new Set([
  'moves',
  'comments',
  'headersFooters',
  'notes',
  'textBoxes',
  'fields',
])

export const settingAvailable = (key: CompareSettingKey, target: CompareTarget): boolean =>
  target === 'new' || !NEW_DOCUMENT_ONLY.has(key)

/** changes can be written into a document only when it is the one open here, and editable */
export function targetAvailable(
  target: CompareTarget,
  original: CompareSide | null,
  revised: CompareSide | null,
  canWrite = true,
): boolean {
  if (target === 'new') return true
  return canWrite && (target === 'original' ? original : revised)?.kind === 'current'
}

/** the settings as the comparison runs them: one the target cannot honour is off */
export function effectiveSettings(
  settings: CompareDialogSettings,
  target: CompareTarget,
): CompareDialogSettings {
  const out = { ...settings }
  for (const key of COMPARE_SETTING_KEYS) if (!settingAvailable(key, target)) out[key] = false
  return out
}

/** the original's own changes can be named only when it is read from its file, not edited in place */
export const originalLabelAvailable = (
  mode: CompareRequest['mode'],
  original: CompareSide | null,
  target: CompareTarget,
): boolean =>
  mode === 'combine' && !!original && !(original.kind === 'current' && target === 'original')
