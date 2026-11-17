/**
 * The scanned-PDF reader: while the user's Antigravity quota is plentiful it sends page images of
 * not-yet-readable PDFs to Antigravity, stores the transcription and lets the normal indexing
 * pipeline chunk and embed it (see ocr-sidecar.ts for how text re-enters the index). Everything
 * that touches the outside world is injected, so the whole policy (quota hysteresis, gating,
 * ordering, batching, failure and backoff rules) is tested without agy, PDFium, Electron or a
 * real clock.
 *
 * Rules that keep it safe to leave on:
 *  - quota-paced: `/usage` is read before and after automatic calls. Provider-cycle reserves
 *    are 80/60/40/20/20 by five-hour window hour and 88/76/64/52/40/28/16 by weekly cycle day
 *    at the default 12-point allowance. A two-point start margin avoids oscillation. Measured
 *    daily spending and unfinished usage baselines survive restart; delayed usage blocks new
 *    automatic calls. Explicit Unlimited and manual OCR bypass app quota policy;
 *  - nothing runs on battery / while the user is active / while indexing is paused (settings);
 *  - quota, rate-limit, sign-in and missing-CLI errors end the day's run (no retries today);
 *  - any other failure backs off exponentially (10 min, 20 min, 40 min ...) and stops the run;
 *  - a file that fails 3 times, or cannot be read at all (password, damaged), is marked
 *    non-retryable with a plain reason and is never tried again by the scheduler;
 *  - optional safety caps: PDFs per day (0 = unlimited), pages per file, pages per call.
 */
import { basename } from 'node:path'
import { clampPdfPages, DEFAULT_PDF_PAGES } from './chunks'
import { folderPauses } from './folder-pause'
import {
  OCR_MAX_CALLS_PER_RUN,
  ocrPageLimit,
  OCR_MAX_FILE_ATTEMPTS,
  OCR_TICK_MS,
  classifyAgyOcrError,
  agyUsageGroupName,
  decideQuota,
  haltsTheDay,
  localDateKey,
  ocrBackoffMs,
  orderOcrCandidates,
  parseAgyOcrOutput,
  pdfCapReached,
  planOcrBatch,
  type AgyUsageReading,
  type OcrCandidate,
  type OcrErrorKind,
  type QuotaBucketView,
  type QuotaRules,
  type TimezoneOffset,
} from '@genoffice/ai-provider/agy-ocr'
import type {
  AgyOcrActivity,
  AgyOcrBucketLive,
  AgyOcrReadNowResult,
  AgyOcrSettings,
  AgyOcrStatus,
} from '../../shared/fork/agy-ocr'
import type { OcrRenderRequest, OcrRenderResult } from './agy-ocr-render'
import type { OcrStateStore } from './agy-ocr-state'
import type { OcrDocRow, OcrFileMeta, OcrPageText } from './ocr-sidecar'
import {
  reconcileOcrBudget,
  reserveOcrBudget,
  ocrWindowReserves,
  stabilizeOcrQuotaSnapshot,
  type OcrQuotaSnapshot,
} from './ocr-auto-budget'

/** user idle time (seconds) that counts as "idle"; the indexing policy uses the same figure */
export const OCR_IDLE_SECONDS = 120

// ---- gating -------------------------------------------------------------------------------

export interface OcrPolicyView {
  paused: boolean
  onBattery: boolean
  /** on battery: 3 = 80% or more, 2 = 50-79% or unknown, 1 = below 50% (see the indexing policy) */
  batteryBand?: number
}

/** On battery, reading is allowed from this band up (half a charge or more). */
const OCR_MIN_BATTERY_BAND = 2

export type OcrGateReason = 'indexing-paused' | 'on-battery' | 'not-idle' | 'no-power-info'

/** May the reader start (or continue) a call right now? */
export function evaluateOcrGate(input: {
  settings: Pick<AgyOcrSettings, 'onlyOnAC' | 'onlyWhenIdle'>
  /** latest indexing policy (battery / low memory / lock / thermal); null before the first reading */
  policy: OcrPolicyView | null
  /** seconds since the last keyboard or mouse input; null when unknown */
  idleSeconds: number | null
}): { ok: true } | { ok: false; reason: OcrGateReason } {
  const { settings, policy, idleSeconds } = input
  if (!policy) return { ok: false, reason: 'no-power-info' }
  if (policy.paused) return { ok: false, reason: 'indexing-paused' }
  // "only on AC" still lets a battery with a good charge work; below half it waits for the charger
  if (settings.onlyOnAC && policy.onBattery && !((policy.batteryBand ?? 0) >= OCR_MIN_BATTERY_BAND))
    return { ok: false, reason: 'on-battery' }
  if (settings.onlyWhenIdle && !(idleSeconds !== null && idleSeconds >= OCR_IDLE_SECONDS))
    return { ok: false, reason: 'not-idle' }
  return { ok: true }
}

/** The quota floors the settings describe; an ignored bucket never blocks. */
export function quotaRulesOf(settings: AgyOcrSettings): QuotaRules {
  return {
    weekly: {
      firstDayFloor: settings.weeklyFirstDayFloor,
      dropPerDay: settings.weeklyDropPerDay,
      minFloor: settings.weeklyMinFloor,
      ignore: settings.ignoreWeekly,
    },
    fiveHour: {
      floorStart: settings.fiveHourFloorStart,
      floorEnd: settings.fiveHourFloorEnd,
      ignore: settings.ignoreFiveHour,
    },
  }
}

// ---- dependencies ---------------------------------------------------------------------------

export interface OcrSavePagesResult {
  ok: boolean
  savedCount?: number
  error?: string
}

export interface OcrJobHost {
  /** document memory is switched on (nothing is read while it is off) */
  isEnabled(): boolean
  /**
   * `manual`: the person asked for this file, so pages the local pass accepted are not skipped
   * and nothing waits for the local pass (automatic runs only ever see escalated pages).
   */
  candidates(maxPagesPerFile: number, options?: { manual?: boolean }): OcrDocRow[]
  documentById(id: number): { id: number; path: string } | null
  pagesDone(path: string, mtimeMs: number, sizeBytes: number): number[]
  /** null = the index process did not answer in time */
  render(path: string, request: OcrRenderRequest): Promise<OcrRenderResult | null>
  savePages(
    path: string,
    meta: OcrFileMeta,
    pages: readonly OcrPageText[],
  ): Promise<OcrSavePagesResult | void> | OcrSavePagesResult | void
  /** queue the document for re-extraction so the new OCR text becomes chunks and vectors */
  onSourcesRecovered?(clear: (paths: string[]) => boolean): () => void
  recordSourceFailure?(
    path: string,
    meta: { mtimeMs: number; sizeBytes: number },
    code: string,
    detail?: string,
  ): void
  reindex(path: string): void
  /**
   * Put the new OCR text into the index right now and resolve once the file is searchable: for a
   * read the person asked for, which must not wait in the line behind every other file.
   */
  reindexNow?(path: string): Promise<void>
}

export interface OcrRecognizeInput {
  model: string
  /** 1-based PDF page numbers, in image order */
  pages: number[]
  images: Uint8Array[]
  signal: AbortSignal
}

export interface OcrJobDeps {
  settings(): AgyOcrSettings
  /** how many pages of a PDF the indexer reads: nothing past it is worth a transcription */
  pdfPageLimit?(): number
  host: OcrJobHost
  state: OcrStateStore
  /** one agy call for several page images; throws agy's own error message on failure */
  recognize(input: OcrRecognizeInput): Promise<{
    text: string
    usage?: { inputTokens?: number; outputTokens?: number; thinkingTokens?: number }
  }>
  /** Optional local VLM fallback recognizer (LFM3-VL-2B) when agy is unavailable. */
  recognizeLocal?: OcrJobDeps['recognize']
  /** Resolves which escalation backend handles low-confidence OCR pages. */
  resolveBackend?: () => 'agy' | 'local-vlm' | 'none'
  /** `agy /usage` (free); null when it cannot be read. Must not throw. */
  readUsage(): Promise<AgyUsageReading | null>
  policy(): OcrPolicyView | null
  idleSeconds(): number | null
  /**
   * The local "light index" pass (local-ocr/). Started from every scheduler tick and when it is switched
   * on, independently of the cloud reader's own switch; it is single-flight and applies its own gates.
   */
  localPass?(): Promise<unknown>
  now(): number
  timezoneOffset: TimezoneOffset
  /** repeating timer; returns the cancel function */
  every(callback: () => void, ms: number): () => void
}

/** What a call to processFile tells the loop that runs it. */
interface FileOutcome {
  pages: number
  /** end the whole run (halt, backoff, quota, gate closed, aborted) */
  stop: boolean
}

const WAITING_CACHE_MS = 60_000
const NO_WORK_RECHECK_MS = 60 * 60_000
/** a blocked quota is not asked again before its window refreshes (capped), see nextCheckAt */
const MAX_QUOTA_WAIT_MS = 6 * 60 * 60_000
const OCR_STAGE_TIMEOUT_MS = 120_000
const OCR_CALL_TIMEOUT_MS = 180_000
const OCR_MAX_MANUAL_QUEUE = 2000

/** A broken worker/provider must not hold the OCR queue indefinitely. */
async function boundedOcr<T>(
  work: Promise<T>,
  ms: number,
  signal?: AbortSignal,
  onTimeout?: () => void,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const finish = (callback: () => void) => {
      clearTimeout(timer)
      signal?.removeEventListener('abort', aborted)
      callback()
    }
    const aborted = () =>
      finish(() => reject(Object.assign(new Error('OCR cancelled'), { name: 'AbortError' })))
    const timer = setTimeout(
      () =>
        finish(() => {
          reject(new Error('OCR step timed out; retry later'))
          onTimeout?.()
        }),
      ms,
    )
    signal?.addEventListener('abort', aborted, { once: true })
    if (signal?.aborted) aborted()
    work.then(
      (value) => finish(() => resolve(value)),
      (error) => finish(() => reject(error)),
    )
  })
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

const RENDER_REASONS: Record<string, string> = {
  password: 'The PDF is password-protected',
  corrupt: 'The PDF is damaged',
  unsupported: 'The PDF cannot be opened',
  'too-large': 'The PDF is too large to read with OCR',
  missing: 'The file is no longer there',
}

export class AgyOcrJob {
  private cancelTick: (() => void) | null = null
  private running = false
  private abort: AbortController | null = null
  private stopped = false
  /** do not look for work before this time (nothing to do, or the quota is blocked until a reset) */
  private nextCheckAt = 0
  private waiting: { at: number; count: number } | null = null
  private callsThisRun = 0
  /** manual reads wait for each other here, so a batch of "read now" never meets "busy" */
  private manualChain: Promise<unknown> = Promise.resolve()
  private idleWaiters: Array<() => void> = []
  private queuedIds = new Set<number>()
  private drainingQueue = false
  private currentFile?: string
  private currentPath?: string
  private stage?: 'quota' | 'rendering' | 'recognizing' | 'indexing'
  private pendingManual = 0
  private cancelEpoch = 0
  private usageReading: Promise<AgyUsageReading | null> | null = null
  private quotaSnapshot: OcrQuotaSnapshot | null = null
  private automaticCallActive = false
  private currentProgress?: { done: number; total: number }

  private activeBackend(): 'agy' | 'local-vlm' | 'none' {
    return this.deps.resolveBackend?.() ?? 'agy'
  }

  private usesAutoBudget(settings: AgyOcrSettings): boolean {
    if (this.activeBackend() === 'local-vlm') return false
    return (
      settings.autoUnlimited !== true && typeof settings.autoWeeklyDailyBudgetPercent === 'number'
    )
  }

  constructor(private readonly deps: OcrJobDeps) {
    for (const item of deps.state.get().manualQueue ?? []) {
      if (deps.host.documentById(item.id)?.path === item.path) this.queuedIds.add(item.id)
    }
  }

  private persistQueue(): boolean {
    this.deps.state.update((data) => {
      data.manualQueue = [...this.queuedIds].flatMap((id) => {
        const document = this.deps.host.documentById(id)
        return document ? [{ id, path: document.path }] : []
      })
    })
    return !this.deps.state.hasPersistenceFailure()
  }

  private recoverReindex(): void {
    this.flushReindex(new Set(this.deps.state.get().pendingReindexPaths ?? []))
  }

  private clearReindexJournal(path: string): void {
    this.deps.state.update((data) => {
      data.pendingReindexPaths = (data.pendingReindexPaths ?? []).filter((item) => item !== path)
    })
  }

  async refreshQuota(): Promise<AgyOcrStatus> {
    await this.quotaAllows(this.workingSettings(), true)
    return this.status()
  }

  // ---- lifecycle ----

  private stopSourceRecovery?: () => void

  start(): void {
    this.stopSourceRecovery ??= this.deps.host.onSourcesRecovered?.((paths) => {
      const saved = this.deps.state.clearMissingSources(paths)
      if (saved) {
        this.nextCheckAt = 0
        setImmediate(() => void this.tick())
      }
      return saved
    })
    for (const item of this.deps.state.get().manualQueue ?? []) {
      if (this.deps.host.documentById(item.id)?.path === item.path) this.queuedIds.add(item.id)
    }
    if (this.cancelTick || this.stopped) return
    this.cancelTick = this.deps.every(() => void this.tick(), OCR_TICK_MS)
    this.recoverReindex()
    if (this.queuedIds.size) void this.drainQueue()
  }

  stop(): void {
    this.stopped = true
    this.stopSourceRecovery?.()
    this.stopSourceRecovery = undefined
    this.cancelTick?.()
    this.cancelTick = null
    this.abort?.abort()
    // Shutdown preserves confirmed manual work for restart; user cancellation clears it.
  }

  /** Settings changed: re-evaluate soon; a re-enabled reader gets a fresh chance (clears a halt). */
  settingsChanged(previous: AgyOcrSettings | null): void {
    const now = this.deps.settings()
    this.waiting = null
    this.nextCheckAt = 0
    const reenabled = !previous?.enabled && now.enabled
    if (reenabled || previous?.model !== now.model)
      this.deps.state.update((data) => {
        data.armed = { fiveHour: false, weekly: false }
        if (reenabled) {
          delete data.halted
          data.backoff = { failures: 0, until: 0 }
        }
      })
    if (previous?.enabled && !now.enabled) this.abort?.abort()
    if (!previous?.localOcr?.enabled && now.localOcr?.enabled) this.startLocalPass()
  }

  // ---- scheduler ----

  /** One cheap check; does real work only when a run is due. Never throws. */
  async tick(): Promise<void> {
    try {
      const settings = this.workingSettings()
      if (this.stopped || this.running) return
      if (!this.deps.host.isEnabled()) return
      this.startLocalPass()
      this.recoverReindex()
      if (this.queuedIds.size) {
        await this.drainQueue()
        return
      }
      if (!settings.enabled) return
      const now = this.deps.now()
      if (now < this.nextCheckAt) return
      const state = this.deps.state.get()
      const today = this.deps.state.today()
      const day = localDateKey(now, this.deps.timezoneOffset)
      if (state.halted && state.halted.day === day) return
      if (state.backoff.until > now) return
      if (!settings.autoUnlimited && pdfCapReached(settings.maxPdfsPerDay, today)) return
      if (!this.gateOpen(settings)) return
      await this.runScheduled(settings)
    } catch (error) {
      this.recordError(errorText(error))
    }
  }

  private startLocalPass(): void {
    try {
      void this.deps.localPass?.().catch((error: unknown) => this.recordError(errorText(error)))
    } catch (error) {
      this.recordError(errorText(error))
    }
  }

  private gateOpen(settings: AgyOcrSettings): boolean {
    return evaluateOcrGate({
      settings,
      policy: this.deps.policy(),
      idleSeconds: this.deps.idleSeconds(),
    }).ok
  }

  /**
   * Read `/usage` (free) and apply the quota floors. True only when a call may start now; an
   * unreadable usage never runs.
   */
  private async quotaAllows(settings: AgyOcrSettings, forceRead = false): Promise<boolean> {
    if (this.activeBackend() === 'local-vlm') return !forceRead || true
    if (settings.autoUnlimited === true && !forceRead) return true
    const budgetMode = this.usesAutoBudget(settings)
    const rules = budgetMode
      ? {
          weekly: { firstDayFloor: 0, dropPerDay: 0, minFloor: 0, ignore: false },
          fiveHour: { floorStart: 0, floorEnd: 0, ignore: false },
          marginPoints: 0,
        }
      : quotaRulesOf(settings)
    // both limits switched off: the quota is not consulted at all (the usage is not even read)
    if (!forceRead && rules.fiveHour.ignore && rules.weekly.ignore) return true
    if (!forceRead) this.stage = 'quota'
    // Settings refresh and the OCR scheduler share one free usage call. A refresh must not
    // overwrite the active file's stage or start another provider command behind its back.
    if (!this.usageReading) {
      this.usageReading = boundedOcr(this.deps.readUsage(), 15_000)
        .catch(() => null)
        .finally(() => {
          this.usageReading = null
        })
    }
    const reading = await boundedOcr(
      this.usageReading,
      16_000,
      forceRead ? undefined : this.abort?.signal,
    ).catch(() => null)
    if (settings.model !== this.deps.settings().model) return false
    const now = this.deps.now()
    let decision = decideQuota({
      reading,
      model: settings.model,
      armed: this.deps.state.get().armed,
      rules,
      now,
    })
    if (forceRead && rules.fiveHour.ignore && rules.weekly.ignore) {
      // Both disabled reserves permit work without consulting usage. A person requesting a
      // refresh still needs the actual bucket values, without enabling either reserve.
      const observed = decideQuota({
        reading,
        model: settings.model,
        armed: this.deps.state.get().armed,
        rules: {
          fiveHour: { ...rules.fiveHour, ignore: false },
          weekly: { ...rules.weekly, ignore: false },
        },
        now,
      })
      if (observed.group) decision.group = observed.group
      if (observed.fiveHour) decision.fiveHour = { ...observed.fiveHour, ignored: true }
      if (observed.weekly) decision.weekly = { ...observed.weekly, ignored: true }
    }
    if (budgetMode) {
      const five = decision.fiveHour
      const week = decision.weekly
      this.quotaSnapshot =
        decision.group && five && week && five.resetAt && five.resetAt > now
          ? {
              group: agyUsageGroupName(settings.model) ?? decision.group,
              fiveHour: five.remaining * 100,
              fiveHourResetAt: five.resetAt,
              weekly: week.remaining * 100,
              ...(week.resetAt ? { weeklyResetAt: week.resetAt } : {}),
            }
          : null
      if (this.quotaSnapshot)
        this.quotaSnapshot = stabilizeOcrQuotaSnapshot(
          this.deps.state.get().autoBudgets?.[this.quotaSnapshot.group],
          this.quotaSnapshot,
          now,
        )
      const reserves = this.quotaSnapshot
        ? ocrWindowReserves(this.quotaSnapshot, now, settings.autoWeeklyDailyBudgetPercent)
        : null
      if (!this.quotaSnapshot || !reserves) {
        decision.kind = 'unreadable'
        decision.run = false
        this.deps.state.update((data) => {
          data.quota = { at: now, decision }
        })
        return false
      }
      const snapshot = this.quotaSnapshot
      decision = decideQuota({
        reading,
        model: settings.model,
        armed: this.deps.state.get().armed,
        now,
        rules: {
          fiveHour: {
            floorStart: reserves.fiveHourReserve,
            floorEnd: reserves.fiveHourReserve,
            ignore: false,
          },
          weekly: {
            firstDayFloor: reserves.weeklyReserve,
            dropPerDay: 0,
            minFloor: reserves.weeklyReserve,
            ignore: false,
          },
          marginPoints: 2,
        },
      })
      if (decision.blocking)
        decision.blocking.clearsAt =
          decision.blocking.window === '5h' ? reserves.fiveHourClearsAt : reserves.weeklyClearsAt
      this.deps.state.update((data) => {
        data.armed = decision.armed
        data.quota = { at: now, decision }
        const budgets = (data.autoBudgets ??= {})
        if (!this.automaticCallActive)
          budgets[snapshot.group] = reconcileOcrBudget(
            budgets[snapshot.group],
            snapshot,
            reserves.dayKey,
            now,
          )
      })
      const account = this.deps.state.get().autoBudgets![snapshot.group]!
      const allowed =
        !account.pending && account.weeklySpent + 1e-6 < settings.autoWeeklyDailyBudgetPercent!
      if (account.pending) {
        this.nextCheckAt = now + OCR_TICK_MS
        return false
      }
      if (!allowed) this.nextCheckAt = reserves.dayResetAt + 1000
      else if (!decision.run)
        this.nextCheckAt = Math.min(reserves.nextFiveHourStepAt, reserves.dayResetAt) + 1000
      return decision.run && allowed && !this.deps.state.hasPersistenceFailure()
    }
    this.deps.state.update((data) => {
      data.armed = decision.armed
      data.quota = { at: now, decision }
    })
    if (decision.kind === 'blocked' && decision.blocking) {
      // the floor only drops on a schedule and quota only grows at a refill: do not ask before then
      const wake = decision.blocking.clearsAt ?? decision.blocking.resetAt
      if (wake !== undefined) this.nextCheckAt = Math.min(wake + 30_000, now + MAX_QUOTA_WAIT_MS)
    }
    return decision.run
  }

  private async runScheduled(settings: AgyOcrSettings): Promise<void> {
    this.running = true
    this.abort = new AbortController()
    this.callsThisRun = 0
    const touched = new Set<string>()
    let quotaBlocked = false
    try {
      const now = this.deps.now()
      const queue = this.eligible(settings.maxPagesPerFile)
      if (queue.length === 0) {
        this.nextCheckAt = now + NO_WORK_RECHECK_MS
        return
      }
      const ordered = orderOcrCandidates(queue, now, settings.maxPagesPerFile)
      for (const candidate of ordered) {
        if (this.abort.signal.aborted || this.callsThisRun >= OCR_MAX_CALLS_PER_RUN) break
        if (
          !settings.autoUnlimited &&
          pdfCapReached(settings.maxPdfsPerDay, this.deps.state.today())
        )
          break
        const outcome = await this.processFile(candidate, settings, { manual: false, touched })
        if (outcome.stop) {
          quotaBlocked = this.deps.state.get().quota?.decision.run === false
          break
        }
      }
      // record where the quota stands after the work (free): the next decision starts from it
      if (this.callsThisRun > 0 && !quotaBlocked && !this.abort.signal.aborted)
        await this.quotaAllows(settings)
      this.deps.state.update((data) => void (data.lastRunAt = this.deps.now()))
    } finally {
      this.flushReindex(touched)
      this.markIdle()
      this.waiting = null
    }
  }

  /** Settings with the "unlimited pages" marker resolved; `status()` keeps the raw value for the UI. */
  private workingSettings(): AgyOcrSettings {
    const settings = this.deps.settings()
    // text read from pages beyond the indexer's page limit is never indexed, so it is never paid
    // for: the reader stops at that page ("unlimited" included, "read now" included)
    const indexed = clampPdfPages(this.deps.pdfPageLimit?.() ?? DEFAULT_PDF_PAGES)
    return {
      ...settings,
      maxPagesPerFile: Math.min(ocrPageLimit(settings.maxPagesPerFile), indexed),
    }
  }

  /** Candidates the scheduler may still try (non-retryable files are out). */
  private eligible(maxPagesPerFile: number): OcrCandidate[] {
    const files = this.deps.state.get().files
    const out: OcrCandidate[] = []
    for (const row of this.deps.host.candidates(maxPagesPerFile)) {
      const failure = files[row.path]
      if (
        failure?.nonRetryable &&
        (!['missing', 'The file is no longer there'].includes(failure.reason ?? '') ||
          this.deps.now() < failure.updatedAt + 24 * 3600_000)
      )
        continue
      out.push({
        path: row.path,
        sizeBytes: row.sizeBytes,
        lastOpenedAt: row.lastOpenedAt,
        mtimeMs: row.mtimeMs,
        pagesDone: row.pagesDone,
        ...(row.totalPages !== undefined ? { totalPages: row.totalPages } : {}),
        ...(row.skipPages ? { skipPages: row.skipPages } : {}),
      })
    }
    return out
  }

  // ---- one file ----

  /**
   * Read pages of one file, call after call, until its page limit, a failure, a closed gate or
   * quota, or an abort. Pages are charged to today's counters BEFORE the call (a crash mid-call
   * cannot lead to a free retry) and refunded only when the failure proves nothing was consumed.
   */
  private async processFile(
    candidate: OcrCandidate,
    settings: AgyOcrSettings,
    options: { manual: boolean; touched: Set<string>; respectQuota?: boolean },
  ): Promise<FileOutcome> {
    const { host, state } = this.deps
    const path = candidate.path
    this.currentFile = basename(path)
    this.currentPath = path
    let charged = 0
    let counted = false
    const signal = this.abort?.signal
    while (!signal?.aborted) {
      if (this.callsThisRun >= OCR_MAX_CALLS_PER_RUN) return { pages: charged, stop: true }
      if (!options.manual) {
        if (!this.gateOpen(settings)) return { pages: charged, stop: true }
      }
      if (!options.manual || options.respectQuota) {
        if (!(await this.quotaAllows(settings))) return { pages: charged, stop: true }
        if (signal?.aborted) break
      }
      const done = host.pagesDone(path, candidate.mtimeMs, candidate.sizeBytes)
      // pages that have their own text layer count as read: they are never rendered or sent
      const skip = candidate.skipPages ?? []
      this.stage = 'rendering'
      const rendered = await boundedOcr(
        host.render(path, {
          done: [...done, ...skip],
          maxPages: settings.maxPagesPerFile,
          count: settings.pagesPerCall,
        }),
        OCR_STAGE_TIMEOUT_MS,
        signal,
      )
      if (signal?.aborted) break
      if (!rendered) {
        this.failFile(path, 'The index process did not answer in time')
        return { pages: charged, stop: true }
      }
      if (!rendered.ok) {
        if (['missing', 'offline', 'source-busy', 'source-permission'].includes(rendered.code)) {
          if (rendered.code !== 'offline')
            host.recordSourceFailure?.(
              path,
              { mtimeMs: candidate.mtimeMs, sizeBytes: candidate.sizeBytes },
              rendered.code,
              rendered.detail,
            )
          return { pages: charged, stop: false }
        }
        if (rendered.code === 'render') {
          this.failFile(path, rendered.message)
          return { pages: charged, stop: false }
        }
        this.markNonRetryable(path, RENDER_REASONS[rendered.code] ?? rendered.message)
        return { pages: charged, stop: false }
      }
      state.update((data) => {
        const file = (data.files[path] ??= { attempts: 0, updatedAt: 0 })
        file.totalPages = rendered.totalPages
        file.updatedAt = this.deps.now()
      })
      const doneNow = new Set(
        host.pagesDone(path, rendered.mtimeMs, rendered.sizeBytes).concat(done, skip),
      )
      const total = Math.min(rendered.totalPages, settings.maxPagesPerFile)
      this.currentProgress = { done: [...doneNow].filter((page) => page <= total).length, total }
      const usingLocal = this.activeBackend() === 'local-vlm'
      const effectivePerCall = usingLocal
        ? Math.min(settings.pagesPerCall, 2)
        : settings.pagesPerCall
      const wanted = planOcrBatch({
        totalPages: rendered.totalPages,
        done: doneNow,
        maxPagesPerFile: settings.maxPagesPerFile,
        pagesPerCall: effectivePerCall,
        budget: effectivePerCall,
      })
      const pages = rendered.pages.filter((p) => wanted.includes(p.page))
      if (pages.length === 0) break // nothing left to read in this file
      const numbers = pages.map((p) => p.page)
      if (!options.manual && this.usesAutoBudget(settings)) {
        const snapshot = this.quotaSnapshot
        if (!snapshot) return { pages: charged, stop: true }
        state.update((data) => {
          const budgets = (data.autoBudgets ??= {})
          budgets[snapshot.group] = reserveOcrBudget(budgets[snapshot.group]!, snapshot)
        })
        if (state.hasPersistenceFailure()) {
          this.recordError('Automatic OCR paused: its usage budget could not be saved')
          return { pages: charged, stop: true }
        }
        this.automaticCallActive = true
      }

      if (!counted) {
        counted = true
        state.update((data) => void (data.day.pdfs += 1))
      }
      this.callsThisRun++
      this.charge(pages.length, 1)
      charged += pages.length
      let response: Awaited<ReturnType<OcrJobDeps['recognize']>>
      let parsed: ReturnType<typeof parseAgyOcrOutput>
      let got: OcrPageText[]
      let automaticAllowed = true
      try {
        this.stage = 'recognizing'
        const recognizeFn =
          usingLocal && this.deps.recognizeLocal ? this.deps.recognizeLocal : this.deps.recognize
        response = await boundedOcr(
          recognizeFn({
            model: usingLocal ? 'LFM3-VL-2B-Q4_K_M' : settings.model,
            pages: numbers,
            images: pages.map((p) => p.jpeg),
            signal: signal ?? new AbortController().signal,
          }),
          OCR_CALL_TIMEOUT_MS,
          signal,
          () => this.abort?.abort(),
        )
        // Save paid-for output before awaiting the free quota command. A crash while usage
        // is slow must not lose recognised pages or require sending those images again.
        if (response.usage)
          state.update((data) => {
            data.day.inputTokens += response.usage?.inputTokens ?? 0
            data.day.outputTokens += response.usage?.outputTokens ?? 0
            data.day.thinkingTokens += response.usage?.thinkingTokens ?? 0
          })
        parsed = parseAgyOcrOutput(response.text, numbers)
        got = numbers
          .filter((n) => parsed.pages.has(n))
          .map((n) => ({ page: n, text: parsed.pages.get(n) ?? '' }))
        if (got.length) {
          const saveRes = await host.savePages(
            path,
            {
              hash: rendered.hash,
              mtimeMs: rendered.mtimeMs,
              sizeBytes: rendered.sizeBytes,
              totalPages: rendered.totalPages,
              model: settings.model,
            },
            got,
          )
          const saveOk =
            saveRes === undefined ||
            (typeof saveRes === 'object' && saveRes !== null && (saveRes as any).ok !== false)

          if (!saveOk) {
            const err =
              typeof saveRes === 'object' &&
              saveRes !== null &&
              typeof (saveRes as any).error === 'string'
                ? (saveRes as any).error
                : 'OCR storage quota denied persistence'
            this.failFile(path, err, false)
            return { pages: charged, stop: true }
          }

          state.update((data) => {
            data.pendingReindexPaths = [...new Set([...(data.pendingReindexPaths ?? []), path])]
          })
          options.touched.add(path)
          this.currentProgress = {
            done: Math.min(total, this.currentProgress.done + got.length),
            total,
          }
        }
        if (signal?.aborted) return { pages: charged, stop: true }
      } catch (error) {
        if (error instanceof Error && error.name === 'AbortError') {
          this.refund(pages.length, 1)
          return { pages: charged - pages.length, stop: true }
        }
        const message = errorText(error)
        const kind = classifyAgyOcrError(message)
        if (haltsTheDay(kind)) {
          if (!options.manual && this.usesAutoBudget(settings) && this.quotaSnapshot) {
            const group = this.quotaSnapshot.group
            state.update((data) => {
              const account = data.autoBudgets?.[group]
              if (account) delete account.pending
            })
          }
          this.refund(pages.length, 1) // the service refused: nothing was read or billed
          this.haltDay(kind, message)
          return { pages: charged - pages.length, stop: true }
        }
        this.failFile(path, message)
        return { pages: charged, stop: true }
      } finally {
        if (!options.manual && this.usesAutoBudget(settings)) {
          this.automaticCallActive = false
          // A UI refresh begun before this call completed cannot account for its final cost.
          // Let that free command settle, then take a fresh post-call snapshot.
          if (this.usageReading) await this.usageReading
          automaticAllowed = await this.quotaAllows(settings)
        }
      }

      if (parsed.missing.length > 0) {
        // the answer did not cover every page: keep what we got, count an attempt, stop here
        this.failFile(
          path,
          `The answer was missing page ${parsed.missing.join(', ')}`,
          got.length > 0,
        )
        return { pages: charged, stop: true }
      }
      this.succeed(path, got.length)
      if (!automaticAllowed) return { pages: charged, stop: true }
      candidate = { ...candidate, pagesDone: candidate.pagesDone + got.length }
      // every page up to the file's limit is read: no further call (or usage read) for this file
      if (doneNow.size + got.length >= Math.min(rendered.totalPages, settings.maxPagesPerFile))
        break
    }
    return { pages: charged, stop: false }
  }

  // ---- state transitions ----

  private charge(pages: number, calls: number): void {
    this.deps.state.update((data) => {
      data.day.pages += pages
      data.day.calls += calls
    })
  }

  private refund(pages: number, calls: number): void {
    this.deps.state.update((data) => {
      data.day.pages = Math.max(0, data.day.pages - pages)
      data.day.calls = Math.max(0, data.day.calls - calls)
    })
  }

  private succeed(path: string, pages: number): void {
    this.deps.state.update((data) => {
      const file = (data.files[path] ??= { attempts: 0, updatedAt: 0 })
      file.attempts = 0
      delete file.nonRetryable
      delete file.reason
      file.updatedAt = this.deps.now()
      data.backoff = { failures: 0, until: 0 }
      data.lastResult = { at: this.deps.now(), pages, file: basename(path) }
      delete data.lastError
    })
  }

  /** A failed call: count it against the file and back off globally. Never retried in a loop. */
  private failFile(path: string, message: string, madeProgress = false): void {
    this.deps.state.update((data) => {
      const file = (data.files[path] ??= { attempts: 0, updatedAt: 0 })
      file.attempts += madeProgress ? 0 : 1
      file.updatedAt = this.deps.now()
      if (file.attempts >= OCR_MAX_FILE_ATTEMPTS) {
        file.nonRetryable = true
        file.reason = `Failed ${OCR_MAX_FILE_ATTEMPTS} times: ${message}`.slice(0, 300)
      }
      data.backoff.failures += 1
      data.backoff.until = this.deps.now() + ocrBackoffMs(data.backoff.failures)
      data.lastError = { at: this.deps.now(), message: message.slice(0, 300), path }
    })
  }

  private markNonRetryable(path: string, reason: string): void {
    this.deps.state.update((data) => {
      const total = data.files[path]?.totalPages
      data.files[path] = {
        attempts: data.files[path]?.attempts ?? 0,
        ...(total !== undefined ? { totalPages: total } : {}),
        nonRetryable: true,
        reason,
        updatedAt: this.deps.now(),
      }
    })
  }

  private haltDay(kind: OcrErrorKind, message: string): void {
    const now = this.deps.now()
    this.deps.state.update((data) => {
      data.halted = {
        day: localDateKey(now, this.deps.timezoneOffset),
        kind,
        message: message.slice(0, 300),
        at: now,
      }
      data.lastError = { at: now, message: message.slice(0, 300) }
    })
  }

  private recordError(message: string): void {
    this.deps.state.update(
      (data) => void (data.lastError = { at: this.deps.now(), message: message.slice(0, 300) }),
    )
  }

  /** Queue re-extraction of every file that gained pages (once per file per run). */
  private flushReindex(touched: Set<string>): void {
    for (const path of touched) {
      try {
        this.deps.host.reindex(path)
        this.clearReindexJournal(path)
      } catch (error) {
        this.recordError(errorText(error))
      }
    }
    touched.clear()
    this.waiting = null
  }

  /** A manual read: the pages just read become searchable text now, not when the line gets to it. */
  private async indexNow(touched: Set<string>): Promise<void> {
    const reindexNow = this.deps.host.reindexNow
    if (!reindexNow) return
    for (const path of [...touched]) {
      try {
        this.stage = 'indexing'
        await boundedOcr(
          reindexNow.call(this.deps.host, path),
          OCR_STAGE_TIMEOUT_MS,
          this.abort?.signal,
        )
        touched.delete(path)
        this.clearReindexJournal(path)
      } catch (error) {
        // left in `touched`: it is queued the ordinary way when this read ends
        this.recordError(errorText(error))
      }
    }
  }

  // ---- manual read of one file ----

  /**
   * "Read with Antigravity now": one file, up to maxPagesPerFile, ignoring the quota thresholds,
   * the halt, the backoff and the gate. The caller has already asked the user to confirm. Pages
   * are still counted in today's totals.
   */
  async readNow(documentId: number): Promise<AgyOcrReadNowResult> {
    if (this.stopped) return { ok: false, error: 'unavailable' }
    if (this.pendingManual >= OCR_MAX_MANUAL_QUEUE) return { ok: false, error: 'queue-full' }
    this.pendingManual++
    const epoch = this.cancelEpoch
    const run = () =>
      epoch === this.cancelEpoch
        ? this.readNowOnce(documentId)
        : Promise.resolve({ ok: false, error: 'cancelled' })
    // one manual read at a time, in the order they were asked for
    const turn = this.manualChain.then(run, run)
    this.manualChain = turn.catch(() => undefined)
    return turn.finally(() => {
      this.pendingManual--
    })
  }

  /** Admit confirmed manual work immediately; manual requests bypass automatic sharing limits. */
  enqueue(documentIds: readonly number[]): { queued: number; skipped: number; error?: string } {
    if (this.stopped || !this.deps.host.isEnabled())
      return { queued: 0, skipped: documentIds.length, error: 'unavailable' }
    let queued = 0
    const beforeQueue = new Set(this.queuedIds)
    for (const id of new Set(documentIds)) {
      const document = this.deps.host.documentById(id)
      if (!document || !/\.pdf$/i.test(document.path) || this.queuedIds.has(id)) continue
      if (this.queuedIds.size >= OCR_MAX_MANUAL_QUEUE) break
      this.queuedIds.add(id)
      queued++
    }
    if (queued) {
      this.deps.state.update((data) => {
        // A confirmed new request may retry after the person fixed login/CLI availability.
        // Provider quota/rate limits still hold; automatic work never clears these guards.
        if (data.halted?.kind === 'auth' || data.halted?.kind === 'cli') {
          delete data.halted
          data.backoff = { failures: 0, until: 0 }
        }
        for (const id of this.queuedIds) {
          if (beforeQueue.has(id)) continue
          const path = this.deps.host.documentById(id)?.path
          const file = path ? data.files[path] : undefined
          if (file) {
            file.attempts = 0
            delete file.nonRetryable
            delete file.reason
          }
        }
      })
      if (!this.persistQueue()) {
        this.queuedIds = beforeQueue
        this.persistQueue()
        return { queued: 0, skipped: documentIds.length, error: 'queue-not-saved' }
      }
      this.nextCheckAt = 0
      void this.drainQueue()
    }
    return { queued, skipped: documentIds.length - queued }
  }

  cancel(): boolean {
    const hadWork = this.running || this.queuedIds.size > 0 || this.pendingManual > 0
    this.cancelEpoch++
    this.queuedIds.clear()
    this.persistQueue()
    this.abort?.abort()
    return hadWork
  }

  cancelDocuments(ids: readonly number[]): number {
    let cancelled = 0
    for (const id of new Set(ids)) {
      const removed = this.queuedIds.delete(id)
      if (removed) cancelled++
      if (this.currentPath && this.deps.host.documentById(id)?.path === this.currentPath) {
        this.abort?.abort()
        if (!removed) cancelled++
      }
    }
    this.persistQueue()
    return cancelled
  }

  private async drainQueue(): Promise<void> {
    if (this.drainingQueue) return
    this.drainingQueue = true
    try {
      await this.yieldToManual()
      while (this.queuedIds.size && !this.stopped) {
        const state = this.deps.state.get()
        if (
          state.backoff.until > this.deps.now() ||
          state.halted?.day === localDateKey(this.deps.now(), this.deps.timezoneOffset)
        )
          break
        const id = [...this.queuedIds].find((id) => {
          const path = this.deps.host.documentById(id)?.path
          return !path || !folderPauses.paused(path)
        })
        if (id === undefined) break
        const path = this.deps.host.documentById(id)?.path
        if (path && state.files[path]?.nonRetryable) {
          this.queuedIds.delete(id)
          this.persistQueue()
          continue
        }
        const turn = this.manualChain.then(() => this.readNowOnce(id, true, false))
        this.manualChain = turn.catch(() => undefined)
        const result = await turn
        if (this.stopped) break
        const latest = this.deps.state.get()
        // Exhausted/permanent file failures require a fresh explicit request, not
        // another paid attempt every time the persisted backoff expires.
        if (path && latest.files[path]?.nonRetryable) {
          this.queuedIds.delete(id)
          this.persistQueue()
        }
        if (
          latest.backoff.until > this.deps.now() ||
          latest.halted?.day === localDateKey(this.deps.now(), this.deps.timezoneOffset)
        )
          break
        if (['quota-reserve', 'halted', 'backoff', 'paused'].includes(result.error ?? '')) break
        this.queuedIds.delete(id)
        this.persistQueue()
        if (!result.ok) break
      }
    } finally {
      this.drainingQueue = false
    }
  }

  /** The scheduled run steps aside after its current call: a person's request goes first. */
  private async yieldToManual(): Promise<void> {
    if (!this.running) return
    this.abort?.abort()
    await new Promise<void>((resolve) => this.idleWaiters.push(resolve))
  }

  private markIdle(): void {
    this.running = false
    this.abort = null
    this.currentFile = undefined
    this.currentProgress = undefined
    this.currentPath = undefined
    this.stage = undefined
    for (const wake of this.idleWaiters.splice(0)) wake()
  }

  private async readNowOnce(
    documentId: number,
    override = true,
    resetAttempts = true,
  ): Promise<AgyOcrReadNowResult> {
    if (this.stopped) return { ok: false, error: 'unavailable' }
    await this.yieldToManual()
    if (!this.deps.host.isEnabled()) return { ok: false, error: 'paused' }
    if (!override) {
      const state = this.deps.state.get()
      if (state.halted?.day === localDateKey(this.deps.now(), this.deps.timezoneOffset))
        return { ok: false, error: 'halted' }
      if (state.backoff.until > this.deps.now()) return { ok: false, error: 'backoff' }
    }
    const document = this.deps.host.documentById(documentId)
    if (document && folderPauses.paused(document.path)) return { ok: false, error: 'paused' }
    if (!document || !/\.pdf$/i.test(document.path)) return { ok: false, error: 'not-pdf' }
    const settings = this.workingSettings()
    const row = this.deps.host
      .candidates(settings.maxPagesPerFile, { manual: true })
      .find((candidate) => candidate.path === document.path)
    if (!row) {
      // Every page has been read already, but the file may never have taken the text in (the
      // earlier request to index it was queued and lost): read it into the index now.
      const reindexNow = this.deps.host.reindexNow
      if (!reindexNow) return { ok: false, error: 'nothing-to-read' }
      try {
        await boundedOcr(reindexNow.call(this.deps.host, document.path), OCR_STAGE_TIMEOUT_MS)
        return { ok: true, pages: 0 }
      } catch (error) {
        return { ok: false, error: errorText(error) }
      }
    }
    this.running = true
    this.abort = new AbortController()
    this.callsThisRun = 0
    const touched = new Set<string>()
    const before = this.deps.state.today().pages
    const startedAt = this.deps.now()
    try {
      if (!override && !(await this.quotaAllows(settings)))
        return { ok: false, error: 'quota-reserve' }
      // a deliberate request outranks the earlier verdicts about this file
      if (resetAttempts)
        this.deps.state.update((data) => {
          const file = data.files[document.path]
          if (file) {
            file.attempts = 0
            delete file.nonRetryable
            delete file.reason
          }
        })
      const outcome = await this.processFile(
        {
          path: row.path,
          sizeBytes: row.sizeBytes,
          lastOpenedAt: row.lastOpenedAt,
          mtimeMs: row.mtimeMs,
          pagesDone: row.pagesDone,
          ...(row.totalPages !== undefined ? { totalPages: row.totalPages } : {}),
          ...(row.skipPages ? { skipPages: row.skipPages } : {}),
        },
        settings,
        { manual: true, respectQuota: !override, touched },
      )
      const state = this.deps.state.get()
      const file = state.files[document.path]
      const failure = file?.nonRetryable
        ? file.reason
        : state.lastError && state.lastError.at >= startedAt
          ? state.lastError.message
          : undefined
      const pages = this.deps.state.today().pages - before
      if (this.abort?.signal.aborted) return { ok: false, error: 'cancelled' }
      await this.indexNow(touched)
      if (touched.size === 0 && failure && (outcome.stop || file?.nonRetryable))
        return { ok: false, error: failure }
      return { ok: true, pages: Math.max(pages, outcome.pages) }
    } catch (error) {
      return { ok: false, error: errorText(error) }
    } finally {
      this.flushReindex(touched)
      this.markIdle()
    }
  }

  // ---- status for the UI ----

  status(): AgyOcrStatus {
    const settings = this.deps.settings()
    const state = this.deps.state.get()
    const today = this.deps.state.today()
    const now = this.deps.now()
    const decision = state.quota?.decision
    const live = (view: QuotaBucketView | undefined): AgyOcrBucketLive | undefined =>
      view
        ? {
            percent: Math.round(view.remaining * 1000) / 10,
            floor: Math.round(view.floor * 10) / 10,
            startAt: Math.round(view.startAt * 10) / 10,
            ignored: view.ignored,
            ...(view.nextFloor !== undefined ? { nextFloor: view.nextFloor } : {}),
            ...(view.nextFloorAt !== undefined ? { nextFloorAt: view.nextFloorAt } : {}),
          }
        : undefined
    const weekly = live(decision?.weekly)
    const fiveHour = live(decision?.fiveHour)
    return {
      settings,
      day: today.day,
      pdfsToday: today.pdfs,
      pagesToday: today.pages,
      tokensToday: {
        input: today.inputTokens,
        output: today.outputTokens,
        thinking: today.thinkingTokens,
      },
      callsToday: today.calls,
      filesWaiting: this.filesWaiting(settings, now),
      running: this.running,
      ...(this.usesAutoBudget(settings)
        ? { autoBudget: this.autoBudgetStatus(settings, now) }
        : {}),
      queuedDocuments: this.queuedIds.size,
      queuedDocumentIds: [...this.queuedIds],
      ...(this.currentProgress ? { progress: this.currentProgress } : {}),
      ...(this.currentFile ? { currentFile: this.currentFile } : {}),
      ...(this.currentPath ? { currentPath: this.currentPath } : {}),
      ...(this.stage ? { stage: this.stage } : {}),
      ...(state.quota && decision && (fiveHour || weekly)
        ? {
            quota: {
              ...(decision.group ? { group: decision.group } : {}),
              ...(fiveHour ? { fiveHour } : {}),
              ...(weekly ? { weekly } : {}),
              readAt: state.quota.at,
            },
          }
        : {}),
      ...(state.lastResult ? { lastResult: state.lastResult } : {}),
      ...(state.lastError ? { lastError: state.lastError } : {}),
      activity: this.activity(settings, now),
      ...(this.deps.resolveBackend ? { activeBackend: this.activeBackend() } : {}),
    }
  }

  private filesWaiting(settings: AgyOcrSettings, now: number): number {
    if (this.waiting && now - this.waiting.at < WAITING_CACHE_MS) return this.waiting.count
    let count: number
    try {
      count = this.eligible(settings.maxPagesPerFile).length
    } catch {
      count = this.waiting?.count ?? 0
    }
    this.waiting = { at: now, count }
    return count
  }

  private activity(settings: AgyOcrSettings, now: number): AgyOcrActivity {
    if (this.running) return { kind: 'working' }
    if (!settings.enabled) return { kind: 'off' }
    const state = this.deps.state.get()
    const day = localDateKey(now, this.deps.timezoneOffset)
    if (state.halted?.day === day) return { kind: 'halted', message: state.halted.message }
    if (state.backoff.until > now) return { kind: 'backoff', until: state.backoff.until }
    if (this.usesAutoBudget(settings)) {
      const budget = this.autoBudgetStatus(settings, now)
      if (budget.pending) return { kind: 'budget-pending' }
      if (budget.weeklyDaily.spent >= budget.weeklyDaily.limit)
        return { kind: 'budget-blocked', window: 'weekly', until: budget.weeklyDaily.resetAt }
    }
    if (this.filesWaiting(settings, now) === 0) return { kind: 'nothing' }
    if (!settings.autoUnlimited && pdfCapReached(settings.maxPdfsPerDay, this.deps.state.today()))
      return { kind: 'cap' }
    const gate = evaluateOcrGate({
      settings,
      policy: this.deps.policy(),
      idleSeconds: this.deps.idleSeconds(),
    })
    if (!gate.ok) return { kind: 'gate', why: gate.reason }
    if (this.activeBackend() === 'local-vlm') return { kind: 'working' }
    if (settings.autoUnlimited) return { kind: 'checking' }
    const decision = state.quota?.decision
    if (!decision) return { kind: 'checking' }
    switch (decision.kind) {
      case 'run':
        return { kind: 'checking' }
      case 'unreadable':
        return { kind: 'quota-unreadable' }
      case 'unknown-group':
        return { kind: 'quota-unknown-group', ...(decision.group ? { group: decision.group } : {}) }
      case 'blocked': {
        const block = decision.blocking
        // the schedule or the refill has passed since the verdict: it is stale, the next check decides
        const due = block?.clearsAt ?? block?.resetAt
        if (!block || (due !== undefined && due <= now)) return { kind: 'checking' }
        return {
          kind: 'quota-blocked',
          window: block.window,
          percent: Math.round(block.remaining * 1000) / 10,
          floor: Math.round(block.floor * 10) / 10,
          startAt: Math.round(block.startAt * 10) / 10,
          belowFloor: block.belowFloor,
          ...(block.clearsAt !== undefined ? { clearsAt: block.clearsAt } : {}),
          ...(block.resetAt !== undefined ? { resetAt: block.resetAt } : {}),
        }
      }
    }
  }

  private autoBudgetStatus(
    settings: AgyOcrSettings,
    now: number,
  ): NonNullable<AgyOcrStatus['autoBudget']> {
    const group = agyUsageGroupName(settings.model)
    const account = group ? this.deps.state.get().autoBudgets?.[group] : undefined
    const reserves = this.quotaSnapshot
      ? ocrWindowReserves(this.quotaSnapshot, now, settings.autoWeeklyDailyBudgetPercent)
      : null
    const recordedDayStart = account?.day ? Date.parse(account.day) : NaN
    const recordedDayResetAt = recordedDayStart + 86_400_000
    const sameRecordedDay = now >= recordedDayStart && now < recordedDayResetAt
    return {
      fiveHour: {
        spent: account?.fiveHourSpent ?? 0,
        limit: 80,
        ...(account ? { resetAt: account.fiveHourResetAt } : {}),
      },
      weeklyDaily: {
        spent:
          account?.day === reserves?.dayKey || sameRecordedDay ? (account?.weeklySpent ?? 0) : 0,
        limit: settings.autoWeeklyDailyBudgetPercent ?? 12,
        resetAt: reserves?.dayResetAt ?? (sameRecordedDay ? recordedDayResetAt : 0),
      },
      pending: !!account?.pending,
    }
  }
}
