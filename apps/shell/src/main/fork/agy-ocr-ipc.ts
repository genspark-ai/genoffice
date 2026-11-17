import { app, powerMonitor } from 'electron'
import type { IpcMain } from 'electron'
import { join } from 'node:path'
import { agyDefaultsUsable } from '@genoffice/ai-provider'
import { buildAgyPrompt, listAgyModels, runAgy } from '@genoffice/ai-provider/agy-cli'
import {
  buildAgyOcrPrompt,
  rankAgyOcrModels,
  systemTimezoneOffset,
} from '@genoffice/ai-provider/agy-ocr'
import { readAgyUsage } from '@genoffice/ai-provider/agy-usage'
import { getLocalVlmRuntime } from '@genoffice/ai-provider/local-vlm'
import { recognizeWithLocalVlm } from './local-vlm-ocr'
import { readAppSettings, writeAppSettingThen } from '../app-settings'
import {
  AGY_OCR_CHANNELS,
  AGY_OCR_SETTINGS_KEY,
  DEFAULT_LOCAL_OCR_SETTINGS,
  agyOcrSettingsFrom,
  mergeAgyOcrSettings,
  type AgyOcrModelList,
  type AgyOcrReadNowResult,
  type AgyOcrSettings,
  type AgyOcrStatus,
} from '../../shared/fork/agy-ocr'
import {
  AgyOcrJob,
  evaluateOcrGate,
  type OcrPolicyView,
  type OcrRecognizeInput,
} from '../document-memory/agy-ocr-job'
import { OcrStateStore } from '../document-memory/agy-ocr-state'
import type { DocumentMemoryManager } from '../document-memory/manager'
import { currentIndexingPolicy } from './indexing-policy-bus'

/**
 * `agy usage` is a child process; the OCR job asks for it on every tick. Share one reading for a
 * short window (and one in-flight call) instead of spawning agy each time.
 */
export const OCR_USAGE_CACHE_MS = 60_000
export function cachedUsageReader<T>(
  read: () => Promise<T>,
  now: () => number,
  ttlMs = OCR_USAGE_CACHE_MS,
): () => Promise<T> {
  let value: { at: number; result: T } | undefined
  let inflight: Promise<T> | undefined
  return () => {
    if (value && now() - value.at < ttlMs) return Promise.resolve(value.result)
    inflight ??= read()
      .then((result) => {
        // a failed read (null) is retried on the next call instead of being remembered for a minute
        if (result !== null && result !== undefined) value = { at: now(), result }
        return result
      })
      .finally(() => {
        inflight = undefined
      })
    return inflight
  }
}

export interface AgyOcrDeps {
  ipcMain: Pick<IpcMain, 'handle'>
  /** absolute path of app-settings.json */
  settingsPath: () => string
  /** userData folder (agy-ocr-state.json lives there) */
  userDataPath: () => string
  getDocumentMemory: () => DocumentMemoryManager | null
}

/** First look at the work queue shortly after launch (the scheduler itself ticks every 10 min). */
const FIRST_TICK_DELAY_MS = 2 * 60_000

/** One real agy call carrying several page JPEGs; the staging folder is removed by runAgy. */
async function recognizeWithAgy(input: OcrRecognizeInput) {
  const plan = buildAgyPrompt('', [
    {
      role: 'user',
      text: buildAgyOcrPrompt(input.pages),
      images: input.images.map((bytes) => ({
        mime: 'image/jpeg',
        base64: Buffer.from(bytes).toString('base64'),
      })),
    },
  ])
  const result = await runAgy({
    model: input.model,
    prompt: plan.prompt,
    files: plan.files,
    signal: input.signal,
    task: 'ocr',
  })
  return {
    text: result.text,
    ...(result.usage
      ? {
          usage: {
            ...(result.usage.promptTokenCount !== undefined
              ? { inputTokens: result.usage.promptTokenCount }
              : {}),
            ...(result.usage.candidatesTokenCount !== undefined
              ? { outputTokens: result.usage.candidatesTokenCount }
              : {}),
            ...(result.usage.thoughtsTokenCount !== undefined
              ? { thinkingTokens: result.usage.thoughtsTokenCount }
              : {}),
          },
        }
      : {}),
  }
}

/**
 * Scanned-PDF reader (Antigravity): settings + status IPC and the daily job. OFF unless the user
 * switched it on in Settings (explicit consent: page images go to Google). Mirrors
 * clipboard-suggest-ipc.ts / indexing-mode-ipc.ts for persistence.
 */
export function registerAgyOcr(deps: AgyOcrDeps): void {
  let settings: AgyOcrSettings = agyOcrSettingsFrom(readAppSettings(deps.settingsPath()))
  let job: AgyOcrJob | null = null
  let jobFor: DocumentMemoryManager | null = null

  const ensureJob = (): AgyOcrJob | null => {
    const manager = deps.getDocumentMemory()
    if (!manager) return null
    if (job && jobFor === manager) return job
    job?.stop()
    jobFor = manager
    const policy = (): OcrPolicyView | null => {
      const published = currentIndexingPolicy()
      return published
        ? {
            paused: published.paused,
            onBattery: published.onBattery,
            ...(published.batteryBand ? { batteryBand: published.batteryBand } : {}),
          }
        : null
    }
    const idleSeconds = (): number | null => {
      try {
        const idle = powerMonitor.getSystemIdleTime()
        return Number.isFinite(idle) ? idle : null
      } catch {
        return null
      }
    }
    // the local pass obeys the same idle / AC / pause switches as the cloud reader
    const local = manager.localOcr(
      () => settings.localOcr ?? DEFAULT_LOCAL_OCR_SETTINGS,
      () => {
        const gate = evaluateOcrGate({ settings, policy: policy(), idleSeconds: idleSeconds() })
        return gate.ok ? { ok: true } : { ok: false, reason: gate.reason }
      },
    )
    const resolveBackend = (): 'agy' | 'local-vlm' | 'none' => {
      const pref = settings.escalationBackend ?? 'auto'
      const runtime = getLocalVlmRuntime()
      const localReady = Boolean(runtime?.isRunnable())
      if (pref === 'local-vlm') return localReady ? 'local-vlm' : 'none'
      if (pref === 'agy') return 'agy'
      if (agyDefaultsUsable()) return 'agy'
      return runtime?.shouldFallbackFromAgy() ? 'local-vlm' : 'agy'
    }
    job = new AgyOcrJob({
      settings: () => settings,
      pdfPageLimit: () => manager.getPdfMaxPages(),
      host: local.host,
      state: new OcrStateStore(
        join(deps.userDataPath(), 'agy-ocr-state.json'),
        Date.now,
        systemTimezoneOffset,
      ),
      recognize: recognizeWithAgy,
      recognizeLocal: recognizeWithLocalVlm,
      resolveBackend,
      readUsage: cachedUsageReader(() => readAgyUsage(), Date.now),
      policy,
      idleSeconds,
      localPass: () => local.runner.tick(),
      now: Date.now,
      timezoneOffset: systemTimezoneOffset,
      every: (callback, ms) => {
        const timer = setInterval(callback, ms)
        timer.unref?.()
        return () => clearInterval(timer)
      },
    })
    job.start()
    return job
  }

  const { ipcMain } = deps
  ipcMain.handle(AGY_OCR_CHANNELS.enqueue, (_event, ids: unknown, confirmed: unknown) => {
    if (
      !Array.isArray(ids) ||
      ids.length > 200 ||
      ids.some((id) => !Number.isSafeInteger(id) || id < 1)
    )
      return { queued: 0, skipped: Array.isArray(ids) ? ids.length : 0, error: 'invalid-request' }
    if (confirmed !== true) return { queued: 0, skipped: ids.length, error: 'not-confirmed' }
    return ensureJob()?.enqueue(ids) ?? { queued: 0, skipped: ids.length, error: 'unavailable' }
  })
  ipcMain.handle(AGY_OCR_CHANNELS.cancel, () => ensureJob()?.cancel() ?? false)
  ipcMain.handle(AGY_OCR_CHANNELS.refreshQuota, () => ensureJob()?.refreshQuota() ?? null)
  ipcMain.handle(AGY_OCR_CHANNELS.cancelDocuments, (_event, ids: unknown) => {
    if (
      !Array.isArray(ids) ||
      ids.length > 200 ||
      ids.some((id) => !Number.isSafeInteger(id) || id < 1)
    )
      return 0
    return ensureJob()?.cancelDocuments(ids) ?? 0
  })
  ipcMain.handle(
    AGY_OCR_CHANNELS.getState,
    (): AgyOcrStatus | null => ensureJob()?.status() ?? null,
  )
  ipcMain.handle(AGY_OCR_CHANNELS.setSettings, (_event, patch: unknown): AgyOcrSettings | null => {
    try {
      const next = mergeAgyOcrSettings(settings, patch)
      writeAppSettingThen(deps.settingsPath(), AGY_OCR_SETTINGS_KEY, next, (saved) => {
        const previous = settings
        settings = saved
        ensureJob()?.settingsChanged(previous)
      })
      return settings
    } catch {
      return null
    }
  })
  ipcMain.handle(AGY_OCR_CHANNELS.listModels, async (): Promise<AgyOcrModelList> => {
    const catalog = await listAgyModels(undefined)
    return {
      models: rankAgyOcrModels(catalog.models),
      ...(catalog.error ? { error: catalog.error } : {}),
    }
  })
  ipcMain.handle(
    AGY_OCR_CHANNELS.readNow,
    async (_event, documentId: unknown, confirmed: unknown): Promise<AgyOcrReadNowResult> => {
      if (typeof documentId !== 'number' || !Number.isSafeInteger(documentId) || documentId < 1)
        return { ok: false, error: 'unavailable' }
      // the renderer asks the user first; a call without the confirmation is refused outright
      if (confirmed !== true) return { ok: false, error: 'not-confirmed' }
      const running = ensureJob()
      return running ? running.readNow(documentId) : { ok: false, error: 'unavailable' }
    },
  )

  void app.whenReady().then(() => {
    const timer = setTimeout(() => {
      const running = ensureJob()
      if (running) void running.tick()
    }, FIRST_TICK_DELAY_MS)
    timer.unref?.()
  })
  app.once('will-quit', () => job?.stop())
}
