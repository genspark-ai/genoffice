/**
 * Scanned-PDF reader (Antigravity) settings, status and IPC channels (fork-only).
 * Persisted in app-settings.json under `agyOcr`; ON by default (it only runs while the
 * Antigravity CLI is signed in, on AC power and when idle). It sends page images of the user's
 * scanned documents to Google through their Antigravity account.
 *
 * Automatic OCR keeps 80/60/40/20/20 percent of the five-hour bucket by elapsed hour. The weekly
 * reserve follows the provider's seven-day cycle: a 12-point daily allowance leaves
 * 88/76/64/52/40/28/16 percent by cycle day, with a 16-percent minimum. Actual automatic usage
 * is measured around calls and persisted per quota group and provider-cycle day. Unknown or
 * delayed usage pauses additional automatic calls until it can be accounted for.
 * Manual OCR explicitly bypasses app reserves and daily caps. Automatic Unlimited is a separate
 * opt-in that bypasses these limits too; provider limits, cancellation and power policy remain.
 */
import {
  AGY_OCR_DEFAULT_MODEL,
  OCR_DEFAULT_FIVE_HOUR_FLOOR_END,
  OCR_DEFAULT_FIVE_HOUR_FLOOR_START,
  OCR_DEFAULT_MAX_PAGES_PER_FILE,
  OCR_DEFAULT_PAGES_PER_CALL,
  OCR_DEFAULT_WEEKLY_DROP_PER_DAY,
  OCR_DEFAULT_WEEKLY_FIRST_DAY_FLOOR,
  OCR_DEFAULT_WEEKLY_MIN_FLOOR,
  OCR_MAX_PAGES_PER_CALL,
  OCR_MAX_PAGES_PER_FILE,
  OCR_MAX_PDFS_PER_DAY,
  validFiveHourFloors,
  validWeeklyFloors,
  type RankedOcrModel,
} from '@genoffice/ai-provider/agy-ocr'

export const AGY_OCR_SETTINGS_KEY = 'agyOcr'

/** 'auto' = Vision on macOS, then TrueOCR and Tesseract. Explicit Tesseract runs only Tesseract. */
export type LocalOcrEngineChoice = 'auto' | 'apple-vision' | 'trueocr-vie' | 'tesseract-vie'

/** Escalation backend for pages that local OCR could not read confidently. */
export type OcrEscalationBackend = 'auto' | 'agy' | 'local-vlm'

/**
 * On-device "light index" OCR that runs BEFORE the cloud reader: the first `lightPages` pages of each
 * scanned PDF / image are read locally (enough to find the file), and only pages the local engine
 * is unsure about are offered to Antigravity. Identity papers and `sensitive` files stay local.
 */
export interface LocalOcrSettings {
  enabled: boolean
  /** leading pages of each file read locally in the first, quick pass (1..5) */
  lightPages: number
  /**
   * Pages of each file read locally in all, 1..50. The pages after `lightPages` are read later by a
   * slow background pass that only runs when the quick pass has nothing left and the computer is on
   * the charger, and never beyond the "PDF pages to index" limit (10 pages on a computer with 4 GB of
   * memory or less); a value at or below `lightPages` turns that pass off.
   */
  maxPages: number
  engine: LocalOcrEngineChoice
}

/**
 * The most pages of one file whose OCR text is put in the index (storage-budget.ts maxOcrPagesPerFile):
 * text of later pages would be read and never found.
 */
export const LOCAL_OCR_MAX_PAGES = 50

export const DEFAULT_LOCAL_OCR_SETTINGS: LocalOcrSettings = {
  enabled: true,
  lightPages: 2,
  // the same as the indexer's default PDF page limit (chunks.ts DEFAULT_PDF_PAGES): a book's contents
  maxPages: 30,
  engine: 'auto',
}

export interface AgyOcrSettings {
  /** Automatic OCR allowance in percentage points; manual OCR does not use this budget. */
  autoWeeklyDailyBudgetPercent?: number
  autoUnlimited?: boolean
  enabled: boolean
  model: string
  /** optional cap on distinct PDFs per local day; 0 = unlimited */
  maxPdfsPerDay: number
  maxPagesPerFile: number
  pagesPerCall: number
  /** Legacy reserve fields retained for compatibility; the current policy uses the cycle budget. */
  weeklyFirstDayFloor: number
  weeklyDropPerDay: number
  weeklyMinFloor: number
  ignoreWeekly: boolean
  /** Legacy linear floors; current automatic OCR uses the fixed hourly staircase. */
  fiveHourFloorStart: number
  fiveHourFloorEnd: number
  ignoreFiveHour: boolean
  onlyOnAC: boolean
  onlyWhenIdle: boolean
  /** local OCR pass; missing in settings saved before it existed (defaults apply) */
  localOcr?: LocalOcrSettings
  /** Escalation backend when local OCR needs VLM help; undefined = 'auto' (agy then local LFM3-VL-2B). */
  escalationBackend?: OcrEscalationBackend
}

export const DEFAULT_AGY_OCR_SETTINGS: AgyOcrSettings = {
  autoWeeklyDailyBudgetPercent: 12,
  autoUnlimited: false,
  enabled: true,
  model: AGY_OCR_DEFAULT_MODEL,
  maxPdfsPerDay: 0,
  maxPagesPerFile: OCR_DEFAULT_MAX_PAGES_PER_FILE,
  pagesPerCall: OCR_DEFAULT_PAGES_PER_CALL,
  weeklyFirstDayFloor: OCR_DEFAULT_WEEKLY_FIRST_DAY_FLOOR,
  weeklyDropPerDay: OCR_DEFAULT_WEEKLY_DROP_PER_DAY,
  weeklyMinFloor: OCR_DEFAULT_WEEKLY_MIN_FLOOR,
  ignoreWeekly: false,
  fiveHourFloorStart: OCR_DEFAULT_FIVE_HOUR_FLOOR_START,
  fiveHourFloorEnd: OCR_DEFAULT_FIVE_HOUR_FLOOR_END,
  ignoreFiveHour: false,
  onlyOnAC: true,
  onlyWhenIdle: true,
  localOcr: DEFAULT_LOCAL_OCR_SETTINGS,
}

const LOCAL_ENGINES: readonly LocalOcrEngineChoice[] = [
  'auto',
  'apple-vision',
  'trueocr-vie',
  'tesseract-vie',
]

const ESCALATION_BACKENDS: readonly OcrEscalationBackend[] = ['auto', 'agy', 'local-vlm']

/** The local OCR settings with defaults filled in; every field validated (stored input is untrusted). */
export function mergeLocalOcrSettings(
  base: LocalOcrSettings | undefined,
  patch: unknown,
): LocalOcrSettings {
  const from = base ?? DEFAULT_LOCAL_OCR_SETTINGS
  const p = (patch && typeof patch === 'object' ? patch : {}) as Record<string, unknown>
  return {
    enabled: typeof p.enabled === 'boolean' ? p.enabled : from.enabled,
    lightPages: int(p.lightPages, 1, 5, from.lightPages),
    // settings stored before this field existed have none
    maxPages: int(
      p.maxPages,
      1,
      LOCAL_OCR_MAX_PAGES,
      from.maxPages ?? DEFAULT_LOCAL_OCR_SETTINGS.maxPages,
    ),
    engine: LOCAL_ENGINES.includes(p.engine as LocalOcrEngineChoice)
      ? (p.engine as LocalOcrEngineChoice)
      : (from.engine ?? DEFAULT_LOCAL_OCR_SETTINGS.engine),
  }
}

// same shape agy-cli accepts (isSafeAgyModelId); duplicated here because agy-cli is Node-only
const MODEL_ID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,99}$/

function int(value: unknown, min: number, max: number, fallback: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return fallback
  return Math.min(max, Math.max(min, Math.round(value)))
}

/** Apply a partial update (renderer input is untrusted) on top of `base`; every field is validated. */
export function mergeAgyOcrSettings(base: AgyOcrSettings, patch: unknown): AgyOcrSettings {
  const p = (patch && typeof patch === 'object' ? patch : {}) as Record<string, unknown>
  // the three weekly numbers are valid only together (0 <= min <= first <= 100); else the old ones stay
  const first = int(p.weeklyFirstDayFloor, 0, 100, base.weeklyFirstDayFloor)
  const drop = int(p.weeklyDropPerDay, 0, 100, base.weeklyDropPerDay)
  const min = int(p.weeklyMinFloor, 0, 100, base.weeklyMinFloor)
  const weeklyOk = validWeeklyFloors(first, drop, min)
  const start = int(p.fiveHourFloorStart, 0, 100, base.fiveHourFloorStart)
  const end = int(p.fiveHourFloorEnd, 0, 100, base.fiveHourFloorEnd)
  const fiveOk = validFiveHourFloors(start, end)
  const escalationBackend = ESCALATION_BACKENDS.includes(
    p.escalationBackend as OcrEscalationBackend,
  )
    ? (p.escalationBackend as OcrEscalationBackend)
    : base.escalationBackend
  return {
    autoWeeklyDailyBudgetPercent: int(
      p.autoWeeklyDailyBudgetPercent,
      0,
      100,
      base.autoWeeklyDailyBudgetPercent ?? 12,
    ),
    autoUnlimited:
      typeof p.autoUnlimited === 'boolean' ? p.autoUnlimited : (base.autoUnlimited ?? false),
    enabled: typeof p.enabled === 'boolean' ? p.enabled : base.enabled,
    model:
      typeof p.model === 'string' && MODEL_ID.test(p.model.trim()) ? p.model.trim() : base.model,
    maxPdfsPerDay: int(p.maxPdfsPerDay, 0, OCR_MAX_PDFS_PER_DAY, base.maxPdfsPerDay),
    maxPagesPerFile: int(p.maxPagesPerFile, 0, OCR_MAX_PAGES_PER_FILE, base.maxPagesPerFile),
    pagesPerCall: int(p.pagesPerCall, 1, OCR_MAX_PAGES_PER_CALL, base.pagesPerCall),
    weeklyFirstDayFloor: weeklyOk ? first : base.weeklyFirstDayFloor,
    weeklyDropPerDay: weeklyOk ? drop : base.weeklyDropPerDay,
    weeklyMinFloor: weeklyOk ? min : base.weeklyMinFloor,
    ignoreWeekly: typeof p.ignoreWeekly === 'boolean' ? p.ignoreWeekly : base.ignoreWeekly,
    fiveHourFloorStart: fiveOk ? start : base.fiveHourFloorStart,
    fiveHourFloorEnd: fiveOk ? end : base.fiveHourFloorEnd,
    ignoreFiveHour: typeof p.ignoreFiveHour === 'boolean' ? p.ignoreFiveHour : base.ignoreFiveHour,
    onlyOnAC: typeof p.onlyOnAC === 'boolean' ? p.onlyOnAC : base.onlyOnAC,
    onlyWhenIdle: typeof p.onlyWhenIdle === 'boolean' ? p.onlyWhenIdle : base.onlyWhenIdle,
    localOcr: mergeLocalOcrSettings(base.localOcr, p.localOcr),
    ...(escalationBackend && escalationBackend !== 'auto' ? { escalationBackend } : {}),
  }
}

/** Settings as stored in app-settings.json; anything missing or invalid falls back to the defaults (OFF). */
export function agyOcrSettingsFrom(stored: Record<string, unknown>): AgyOcrSettings {
  return mergeAgyOcrSettings(DEFAULT_AGY_OCR_SETTINGS, stored[AGY_OCR_SETTINGS_KEY])
}

/** One quota bucket as the UI shows it, in percent. */
export interface AgyOcrBucketLive {
  /** percent left, as last read */
  percent: number
  /** the reserve in force now */
  floor: number
  /** percent needed to start a batch (floor + margin) */
  startAt: number
  /** weekly: tomorrow's floor and when it takes effect (epoch ms) */
  nextFloor?: number
  nextFloorAt?: number
  /** the bucket is switched off by the user ("Ignore this limit") */
  ignored: boolean
}

/** What the reader is doing / waiting for, in terms the UI can phrase. */
export type AgyOcrActivity =
  | { kind: 'budget-pending' }
  | { kind: 'budget-blocked'; window: '5h' | 'weekly'; until: number }
  | { kind: 'off' }
  /** reading now, or allowed and due at the next check */
  | { kind: 'working' }
  /** enabled; the quota has not been read yet */
  | { kind: 'checking' }
  | { kind: 'nothing' }
  | { kind: 'gate'; why: string }
  /** the optional PDFs-per-day cap is used up; resumes tomorrow */
  | { kind: 'cap' }
  /** quota / rate limit / sign-in error stopped today's run */
  | { kind: 'halted'; message: string }
  | { kind: 'backoff'; until: number }
  | {
      kind: 'quota-blocked'
      window: '5h' | 'weekly'
      /** percent left */
      percent: number
      floor: number
      startAt: number
      /** below the floor itself (false: just inside the margin above it) */
      belowFloor: boolean
      /** when the floor schedule lets work resume; absent = only when the quota refills */
      clearsAt?: number
      resetAt?: number
    }
  | { kind: 'quota-unreadable' }
  | { kind: 'quota-unknown-group'; group?: string }

export interface AgyOcrStatus {
  autoBudget?: {
    fiveHour: { spent: number; limit: number; resetAt?: number }
    weeklyDaily: { spent: number; limit: number; resetAt: number }
    pending: boolean
  }
  settings: AgyOcrSettings
  /** local calendar day the counters belong to */
  day: string
  pdfsToday: number
  pagesToday: number
  tokensToday: { input: number; output: number; thinking: number }
  callsToday: number
  /** scanned PDFs that still have pages to read */
  filesWaiting: number
  running: boolean
  queuedDocuments?: number
  queuedDocumentIds?: number[]
  progress?: { done: number; total: number }
  currentFile?: string
  currentPath?: string
  stage?: 'quota' | 'rendering' | 'recognizing' | 'indexing'
  /** latest quota reading for the model's group (undefined = never read) */
  quota?: {
    group?: string
    readAt: number
    weekly?: AgyOcrBucketLive
    fiveHour?: AgyOcrBucketLive
  }
  lastResult?: { at: number; pages: number; file: string }
  lastError?: { at: number; message: string; path?: string }
  activity: AgyOcrActivity
  /** Effective escalation backend currently selected ('agy' | 'local-vlm' | 'none'). */
  activeBackend?: 'agy' | 'local-vlm' | 'none'
}

export const AGY_OCR_CHANNELS = {
  getState: 'agy-ocr:get-state',
  setSettings: 'agy-ocr:set-settings',
  listModels: 'agy-ocr:list-models',
  readNow: 'agy-ocr:read-now',
  enqueue: 'agy-ocr:enqueue',
  cancel: 'agy-ocr:cancel',
  refreshQuota: 'agy-ocr:refresh-quota',
  cancelDocuments: 'agy-ocr:cancel-documents',
} as const

export interface AgyOcrModelList {
  models: RankedOcrModel[]
  error?: string
}

export interface AgyOcrReadNowResult {
  ok: boolean
  pages?: number
  /** 'busy' | 'unavailable' | 'not-pdf' | 'nothing-to-read' | plain-language reason */
  error?: string
}

/** Renderer-facing methods, merged into HomeApi via ForkHomeApi. */
export interface AgyOcrApi {
  getAgyOcrStatus(): Promise<AgyOcrStatus | null>
  setAgyOcrSettings(patch: Partial<AgyOcrSettings>): Promise<AgyOcrSettings | null>
  listAgyOcrModels(): Promise<AgyOcrModelList>
  /** Read one scanned PDF now, ignoring the quota floors; `confirmed` must be true (explicit consent). */
  readScannedPdfWithAgy(documentId: number, confirmed: boolean): Promise<AgyOcrReadNowResult>
  /** Admit up to 200 PDFs without waiting for OCR; respects the configured quota reserves. */
  enqueueScannedPdfsWithAgy?(
    documentIds: number[],
    confirmed: boolean,
  ): Promise<{ queued: number; skipped: number; error?: string }>
  cancelAgyOcr?(): Promise<boolean>
  refreshAgyOcrQuota?(): Promise<AgyOcrStatus | null>
  cancelScannedPdfsWithAgy?(documentIds: number[]): Promise<number>
}
