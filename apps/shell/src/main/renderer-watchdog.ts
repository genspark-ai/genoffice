import { app, dialog, webContents } from 'electron'
import type { BrowserWindow, ProcessMetric, WebContents } from 'electron'
import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'

/**
 * Renderer resource watchdog. A document renderer that keeps a core busy or
 * grows past a couple of gigabytes for minutes on end (a long open paginating
 * chunk by chunk in a background tab, a runaway layout loop) used to go
 * unnoticed until the machine swapped: nothing in the main process looked at
 * renderer CPU or memory, and the `unresponsive` event never fires for a
 * renderer that still answers input. This samples `app.getAppMetrics()`,
 * records diagnostics (metrics plus a short CPU profile of the hot renderer)
 * under userData when a renderer stays hot, and offers to close the document.
 */
export interface WatchdogDocumentInfo {
  kind: string
  title: string
  filePath?: string
}

export type WatchdogStringKey = 'watchdogTitle' | 'watchdogBody' | 'watchdogWait' | 'watchdogClose'

export interface RendererWatchdogOptions {
  /** the document a renderer shows; null for chrome (Home) renderers, which are logged but never prompted */
  describe: (wc: WebContents) => WatchdogDocumentInfo | null
  parentWindow: (wc: WebContents) => BrowserWindow | null
  /** close through the regular close path (save prompt included) */
  closeDocument: (wc: WebContents) => Promise<void> | void
  t: (key: WatchdogStringKey, params?: Record<string, string | number>) => string
  sampleMs?: number
  memoryLimitMB?: number
  cpuLimitPercent?: number
  /** consecutive hot samples before the renderer is reported */
  hotSamplesBeforeReport?: number
  /** minimum gap between two reports of the same renderer */
  reportCooldownMs?: number
  /** duration of the CPU profile captured with a report (0 disables) */
  profileMs?: number
}

const DEFAULTS = {
  sampleMs: 30_000,
  memoryLimitMB: 1536,
  cpuLimitPercent: 90,
  hotSamplesBeforeReport: 10,
  reportCooldownMs: 15 * 60_000,
  profileMs: 5_000,
}

interface RendererState {
  hot: number
  reportedAt: number
  prompting: boolean
  /** cumulative CPU seconds at the previous sample (absent on the first) */
  cpuSeconds?: number
  sampledAt?: number
}

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms))

/** short CDP CPU profile of a live renderer; false when another debugger holds it */
async function captureCpuProfile(wc: WebContents, outPath: string, ms: number): Promise<boolean> {
  if (wc.isDestroyed() || wc.debugger.isAttached()) return false
  wc.debugger.attach('1.3')
  try {
    await wc.debugger.sendCommand('Profiler.enable')
    await wc.debugger.sendCommand('Profiler.setSamplingInterval', { interval: 1000 })
    await wc.debugger.sendCommand('Profiler.start')
    await sleep(ms)
    const { profile } = (await wc.debugger.sendCommand('Profiler.stop')) as { profile: unknown }
    await writeFile(outPath, JSON.stringify(profile))
    return true
  } finally {
    try {
      if (!wc.isDestroyed()) wc.debugger.detach()
    } catch {
      /* already gone */
    }
  }
}

export function startRendererWatchdog(options: RendererWatchdogOptions): () => void {
  const opts = { ...DEFAULTS, ...options }
  const states = new Map<number, RendererState>()

  const rendererFor = (pid: number): WebContents | undefined =>
    webContents.getAllWebContents().find((wc) => !wc.isDestroyed() && wc.getOSProcessId() === pid)

  const report = async (
    wc: WebContents,
    metric: ProcessMetric,
    all: ProcessMetric[],
    memoryMB: number,
    cpu: number,
  ): Promise<void> => {
    const info = opts.describe(wc)
    const stamp = Date.now()
    const base = join(app.getPath('userData'), `renderer-watchdog-${stamp}`)
    try {
      await writeFile(
        `${base}.json`,
        JSON.stringify(
          {
            at: new Date(stamp).toISOString(),
            webContentsId: wc.id,
            pid: metric.pid,
            document: info,
            url: wc.getURL(),
            memoryMB: Math.round(memoryMB),
            cpuPercent: Math.round(cpu),
            metric,
            appMetrics: all,
          },
          null,
          2,
        ),
      )
      if (opts.profileMs > 0) await captureCpuProfile(wc, `${base}.cpuprofile`, opts.profileMs)
    } catch {
      /* diagnostics are best effort and must never make the situation worse */
    }
    // chrome renderers (Home) have no document to close: logging is all that applies
    if (!info || wc.isDestroyed()) return
    const state = states.get(wc.id)
    if (!state || state.prompting) return
    state.prompting = true
    try {
      const parent = opts.parentWindow(wc)
      const dialogOptions = {
        type: 'warning' as const,
        message: opts.t('watchdogTitle'),
        detail: opts.t('watchdogBody', {
          title: info.title,
          memory: Math.round(memoryMB),
          cpu: Math.round(cpu),
        }),
        buttons: [opts.t('watchdogWait'), opts.t('watchdogClose')],
        defaultId: 0,
        cancelId: 0,
      }
      const { response } =
        parent && !parent.isDestroyed()
          ? await dialog.showMessageBox(parent, dialogOptions)
          : await dialog.showMessageBox(dialogOptions)
      if (response === 1 && !wc.isDestroyed()) await opts.closeDocument(wc)
    } catch {
      /* a failed prompt only means the renderer keeps running */
    } finally {
      state.prompting = false
    }
  }

  const sample = (): void => {
    let metrics: ProcessMetric[]
    try {
      metrics = app.getAppMetrics()
    } catch {
      return
    }
    const seen = new Set<number>()
    for (const metric of metrics) {
      if (metric.type !== 'Tab') continue
      const wc = rendererFor(metric.pid)
      if (!wc) continue
      seen.add(wc.id)
      // workingSetSize is reported in kilobytes
      const memoryMB = metric.memory.workingSetSize / 1024
      const state = states.get(wc.id) ?? { hot: 0, reportedAt: 0, prompting: false }
      states.set(wc.id, state)
      // percentCPUUsage is normalized differently per platform (a full core
      // read 10% on a 10-core Mac): derive one-core percent from the
      // cumulative CPU seconds between two samples instead
      const now = Date.now()
      const cpuSeconds = metric.cpu.cumulativeCPUUsage
      let cpu = 0
      if (cpuSeconds !== undefined) {
        if (
          state.cpuSeconds !== undefined &&
          state.sampledAt !== undefined &&
          now > state.sampledAt
        )
          cpu = ((cpuSeconds - state.cpuSeconds) / ((now - state.sampledAt) / 1000)) * 100
        state.cpuSeconds = cpuSeconds
        state.sampledAt = now
      }
      const hot = memoryMB > opts.memoryLimitMB || cpu > opts.cpuLimitPercent
      state.hot = hot ? state.hot + 1 : 0
      if (
        state.hot >= opts.hotSamplesBeforeReport &&
        !state.prompting &&
        Date.now() - state.reportedAt >= opts.reportCooldownMs
      ) {
        state.reportedAt = Date.now()
        state.hot = 0
        void report(wc, metric, metrics, memoryMB, cpu)
      }
    }
    for (const id of states.keys()) if (!seen.has(id)) states.delete(id)
  }

  const timer = setInterval(sample, opts.sampleMs)
  // a pending sample must never keep the app alive at quit
  timer.unref?.()
  return () => clearInterval(timer)
}
