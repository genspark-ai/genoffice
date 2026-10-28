import { test, expect, _electron as electron } from '@playwright/test'
import type { ElectronApplication, Page } from '@playwright/test'
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { arch, cpus, freemem, platform, release, tmpdir, totalmem } from 'node:os'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'
import { SHELL_DIR, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import type { LaunchedApp } from './helpers'

/**
 * Runtime performance baseline (measurement, not optimization).
 *
 * Inert unless GENOFFICE_PERF=1 — CI runs every e2e spec on ubuntu, and these
 * numbers only mean something on a quiet developer machine. Run:
 *   GENOFFICE_PERF=1 npx playwright test --config e2e/playwright.config.ts perf-startup
 *
 * Five consecutive cold launches (fresh mkdtemp userDataDir, onboarding seen,
 * same env/args as launchShell) measure, in wall-clock ms:
 *   t_launch       electron.launch() call -> resolve (main-process boot)
 *   t_window       electron.launch() resolve -> firstWindow() (first BrowserWindow)
 *   t_content_load launch call -> renderer navigation start (content loading)
 *   t_ready        launch call -> document readyState complete (non-about:blank)
 *   t_load         launch call -> renderer 'load' event (PerformanceNavigationTiming)
 *   t_fcp          launch call -> first-contentful-paint (PerformancePaintTiming)
 *   t_interactive  launch call -> Home "AI Docs" new-document quick card visible + enabled
 *
 * Memory: Electron's app.getAppMetrics() after 10s idle on Home, then again
 * after opening a real .md document (window.aiOffice.openPath) and 5s idle.
 *
 * launchShell() cannot be called directly here: it returns only after the page
 * is fully ready, hiding every sub-launch timestamp this baseline needs, and it
 * always records video (Playwright's screencast distorts startup timings). The
 * launch below mirrors helpers.ts launchShell exactly — same userDataDir
 * seeding, env and args — with per-phase timing and video off. Closing still
 * goes through closeAndSaveVideo, which tolerates the missing recording.
 */

const RUNS = 5
const SHELL_MAIN = join(SHELL_DIR, 'out/main/index.js')
const OUT_DIR = resolve(__dirname, '../.task/audit')
const OUT_FILE = join(OUT_DIR, 'perf-baseline.json')

interface RunTimings {
  /** electron.launch() call -> resolve (main-process boot) */
  t_launch: number
  /** electron.launch() resolve -> firstWindow() (includes Playwright attach) */
  t_window: number
  /** launch call -> renderer navigation start (epoch-aligned via renderer Date) */
  t_content_load: number
  /** launch call -> document readyState complete on the final document */
  t_ready: number
  /** launch call -> renderer 'load' event, or -1 when not observable */
  t_load: number
  /** launch call -> first contentful paint, or -1 when not observable */
  t_fcp: number
  /** launch call -> primary action visible + enabled */
  t_interactive: number
  /** navigation start -> load event end, renderer clock */
  load_event_end_since_nav_start: number
  /** navigation start -> first contentful paint, renderer clock */
  fcp_since_nav_start: number
}

interface RendererTimings {
  url: string
  /** wall-clock epoch (ms) of the renderer time origin, derived at read time */
  originEpoch: number
  /** performance.timeOrigin as reported by the renderer (cross-check) */
  timeOrigin: number
  /** navigation start -> load event end (ms), or -1 */
  loadEventEnd: number
  /** navigation start -> first contentful paint (ms), or -1 */
  fcpStart: number
}

interface ProcessEntry {
  pid: number
  type: string
  workingSetKB: number
}

interface MemorySnapshot {
  at: string
  processCount: number
  totalWorkingSetKB: number
  byType: Record<string, { count: number; workingSetKB: number }>
  processes: ProcessEntry[]
}

interface AppVersions {
  electron: string
  node: string
  chrome: string
}

interface RunResult {
  run: number
  userDataDir: string
  timings: RunTimings
  rendererUrl: string
  rendererTimeOrigin: number
  versions: AppVersions
  memoryBefore: MemorySnapshot
  memoryAfter?: MemorySnapshot
}

const roundMs = (v: number) => Math.round(v * 10) / 10

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b)
  const mid = Math.floor(sorted.length / 2)
  const hi = sorted[mid] ?? 0
  const lo = sorted[mid - 1] ?? hi
  return sorted.length % 2 === 1 ? hi : (lo + hi) / 2
}

/** Mirrors helpers.ts waitForDocumentReady (polling; lifecycle events can be missed on attach). */
async function waitDocumentReady(page: Page, timeoutMs = 30_000): Promise<void> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    const ready = await page
      .evaluate(() =>
        document.readyState !== 'loading' && window.location.href !== 'about:blank'
          ? document.readyState
          : null,
      )
      .catch(() => null)
    if (ready) return
    if (Date.now() > deadline) throw new Error(`Shell window never loaded (url: ${page.url()})`)
    await new Promise((r) => setTimeout(r, 100))
  }
}

async function readRendererTimings(page: Page): Promise<RendererTimings> {
  const rt = await page.evaluate(() => {
    const nav = performance.getEntriesByType('navigation')[0] as
      PerformanceNavigationTiming | undefined
    const fcp = performance
      .getEntriesByType('paint')
      .find((p) => p.name === 'first-contentful-paint')
    return {
      url: window.location.href,
      wallNow: Date.now(),
      nowSinceOrigin: performance.now(),
      timeOrigin: performance.timeOrigin,
      loadEventEnd: nav ? nav.loadEventEnd : -1,
      fcpStart: fcp ? fcp.startTime : -1,
    }
  })
  return {
    url: rt.url,
    originEpoch: rt.wallNow - rt.nowSinceOrigin,
    timeOrigin: rt.timeOrigin,
    loadEventEnd: rt.loadEventEnd,
    fcpStart: rt.fcpStart,
  }
}

async function collectMemory(app: ElectronApplication): Promise<MemorySnapshot> {
  const metrics = await app.evaluate(({ app: electronApp }) => electronApp.getAppMetrics())
  const byType: Record<string, { count: number; workingSetKB: number }> = {}
  const processes: ProcessEntry[] = metrics.map((m) => {
    const workingSetKB = m.memory.workingSetSize
    const bucket = (byType[m.type] ??= { count: 0, workingSetKB: 0 })
    bucket.count += 1
    bucket.workingSetKB += workingSetKB
    return { pid: m.pid, type: m.type, workingSetKB }
  })
  return {
    at: new Date().toISOString(),
    processCount: processes.length,
    totalWorkingSetKB: processes.reduce((sum, p) => sum + p.workingSetKB, 0),
    byType,
    processes,
  }
}

async function collectVersions(app: ElectronApplication): Promise<AppVersions> {
  return app.evaluate(() => ({
    electron: process.versions.electron,
    node: process.versions.node,
    chrome: process.versions.chrome,
  }))
}

/** launchShell's exact launch recipe (helpers.ts), instrumented per phase, video off. */
async function launchTimed(): Promise<{
  launched: LaunchedApp
  timings: RunTimings
  renderer: RendererTimings
}> {
  if (!existsSync(SHELL_MAIN)) {
    throw new Error(`Missing build output at ${SHELL_MAIN} — run \`npm run build:all\` first`)
  }
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-perf-'))
  await writeFile(join(userDataDir, 'app-settings.json'), JSON.stringify({ onboardingSeen: true }))
  const require = createRequire(join(SHELL_DIR, 'package.json'))
  const executablePath = require('electron') as unknown as string
  const { ELECTRON_RUN_AS_NODE: _electronRunAsNode, ...hostEnv } = process.env
  const args: string[] = []
  if (process.platform === 'linux') args.push('--no-sandbox', '--disable-gpu')
  args.push(SHELL_DIR)

  const t0 = performance.now()
  const t0Epoch = Date.now()
  const app = await electron.launch({
    executablePath,
    args,
    env: {
      ...hostEnv,
      GENOFFICE_USER_DATA: userDataDir,
      GENOFFICE_NO_SPARE_VIEW: '1',
      GENOFFICE_LANG: 'en',
      ...(process.platform === 'linux' ? { ELECTRON_DISABLE_SANDBOX: '1' } : {}),
    },
  })
  const tLaunch = performance.now() // electron.launch() resolved
  const page = await app.firstWindow()
  const tWindow = performance.now() // first BrowserWindow

  // Started before the app document loads so the moment the primary action
  // ("AI Docs" quick card, Home.tsx renderQuickCards) becomes visible+enabled
  // is caught in the act. The shell window shows immediately (no ready-to-show
  // gating in createShellWindow), so rAF polling runs unthrottled.
  const interactiveAt = page
    .waitForFunction(
      () => {
        const el = document.querySelector('.quick-card')
        return (
          el instanceof HTMLButtonElement &&
          !el.disabled &&
          el.getClientRects().length > 0 &&
          getComputedStyle(el).visibility !== 'hidden'
        )
      },
      undefined,
      { timeout: 60_000 },
    )
    .then(() => performance.now())
  const tInteractive = await interactiveAt
  const tReady = await waitDocumentReady(page).then(() => performance.now())
  // FCP can land after the document is interactive (layout precedes paint);
  // poll for the paint entry instead of reading a single racing snapshot
  await page
    .waitForFunction(
      () => performance.getEntriesByType('paint').some((p) => p.name === 'first-contentful-paint'),
      undefined,
      { timeout: 15_000 },
    )
    .catch(() => undefined)
  const renderer = await readRendererTimings(page)

  const fromLaunch = (epochMs: number) => roundMs(epochMs - t0Epoch)
  const timings: RunTimings = {
    t_launch: roundMs(tLaunch - t0),
    t_window: roundMs(tWindow - tLaunch),
    t_content_load: fromLaunch(renderer.originEpoch),
    t_ready: roundMs(tReady - t0),
    t_load:
      renderer.loadEventEnd >= 0 ? fromLaunch(renderer.originEpoch + renderer.loadEventEnd) : -1,
    t_fcp: renderer.fcpStart >= 0 ? fromLaunch(renderer.originEpoch + renderer.fcpStart) : -1,
    t_interactive: roundMs(tInteractive - t0),
    load_event_end_since_nav_start: roundMs(renderer.loadEventEnd),
    fcp_since_nav_start: roundMs(renderer.fcpStart),
  }
  return { launched: { app, page, userDataDir }, timings, renderer }
}

test('startup and memory baseline across five launches', async () => {
  test.skip(process.env.GENOFFICE_PERF !== '1', 'set GENOFFICE_PERF=1 to run this measurement')
  test.setTimeout(600_000)
  expect(existsSync(SHELL_MAIN), `missing build output: ${SHELL_MAIN}`).toBe(true)

  const runs: RunResult[] = []
  for (let i = 1; i <= RUNS; i += 1) {
    const { launched, timings, renderer } = await launchTimed()
    const { app, page, userDataDir } = launched
    try {
      // settle on Home with no interaction, then sample process memory
      await page.waitForTimeout(10_000)
      const memoryBefore = await collectMemory(app)
      const versions = await collectVersions(app)

      let memoryAfter: MemorySnapshot | undefined
      if (i === RUNS) {
        // a real document: a small markdown file opened like the user would
        const mdPath = join(userDataDir, 'perf-sample.md')
        await writeFile(
          mdPath,
          '# Perf baseline\n\nA small markdown document for the memory-after sample.\n\n- one\n- two\n- three\n',
        )
        await page.evaluate((p) => window.aiOffice.openPath(p), mdPath)
        await waitForPageWithUrl(app, '://markdown/')
        await page.waitForTimeout(5_000)
        memoryAfter = await collectMemory(app)
      }

      runs.push({
        run: i,
        userDataDir,
        timings,
        rendererUrl: renderer.url,
        rendererTimeOrigin: renderer.timeOrigin,
        versions,
        memoryBefore,
        ...(memoryAfter ? { memoryAfter } : {}),
      })
      console.log(
        `perf run ${i}/${RUNS}: t_window=${timings.t_window} t_load=${timings.t_load} ` +
          `t_fcp=${timings.t_fcp} t_interactive=${timings.t_interactive} ` +
          `memory=${memoryBefore.totalWorkingSetKB}KB`,
      )
    } finally {
      await closeAndSaveVideo(launched, `perf-run-${i}`)
    }
    if (i < RUNS) await new Promise((r) => setTimeout(r, 1_000))
  }

  const timingKeys: (keyof RunTimings)[] = [
    't_launch',
    't_window',
    't_content_load',
    't_ready',
    't_load',
    't_fcp',
    't_interactive',
  ]
  const medians: Record<string, number> = {}
  for (const key of timingKeys) medians[key] = median(runs.map((r) => r.timings[key]))

  const beforeTotals = runs.map((r) => r.memoryBefore.totalWorkingSetKB)
  const after = runs[RUNS - 1]?.memoryAfter ?? null
  const byTypeMedian: Record<string, number> = {}
  for (const type of Object.keys(runs[0]?.memoryBefore.byType ?? {})) {
    byTypeMedian[type] = median(runs.map((r) => r.memoryBefore.byType[type]?.workingSetKB ?? 0))
  }

  const lastVersions = runs[runs.length - 1]?.versions
  const machine = {
    platform: platform(),
    release: release(),
    arch: arch(),
    cpuModel: cpus()[0]?.model ?? 'unknown',
    logicalCores: cpus().length,
    totalMemBytes: totalmem(),
    freeMemBytes: freemem(),
    runnerNode: process.version,
    electron: lastVersions?.electron ?? 'unknown',
    electronNode: lastVersions?.node ?? 'unknown',
    chrome: lastVersions?.chrome ?? 'unknown',
  }

  const report = {
    generatedAt: new Date().toISOString(),
    spec: 'e2e/perf-startup.spec.ts',
    machine,
    runs,
    medians,
    memory: {
      before: {
        perRunTotalWorkingSetKB: beforeTotals,
        medianTotalWorkingSetKB: median(beforeTotals),
        medianByTypeWorkingSetKB: byTypeMedian,
      },
      after,
    },
    notes: [
      'Timings are wall-clock ms on this machine; run 1 includes OS/disk cache warm-up.',
      'Renderer-derived values (t_load, t_fcp, t_content_load) are aligned to the test clock via Date.now() read in the renderer at sampling time.',
      'No Playwright video recording: the screencast distorts startup timings. The launch env/args otherwise mirror helpers.ts launchShell.',
      'Fresh userDataDir per launch with onboardingSeen=true, GENOFFICE_NO_SPARE_VIEW=1, English UI.',
      't_interactive polls the first .quick-card button ("AI Docs", Home.tsx renderQuickCards) with rAF polling.',
    ],
  }
  await mkdir(OUT_DIR, { recursive: true })
  await writeFile(OUT_FILE, JSON.stringify(report, null, 2))

  console.log('\n=== GenOffice startup baseline (5 launches, ms) ===')
  console.log(
    `machine: ${machine.platform} ${machine.release} | ${machine.logicalCores} cores | ` +
      `electron ${machine.electron} (node ${machine.electronNode})`,
  )
  for (const key of timingKeys) {
    const values = runs.map((r) => r.timings[key])
    console.log(
      `  ${key.padEnd(14)} median ${String(medians[key]).padStart(9)}  min ${String(
        Math.min(...values),
      ).padStart(9)}  max ${String(Math.max(...values)).padStart(9)}`,
    )
  }
  console.log(
    `memory (workingSetSize, KB): home idle 10s median ${median(beforeTotals)}` +
      ` | markdown doc open 5s: ${after ? after.totalWorkingSetKB : 'n/a'}`,
  )
  console.log(`baseline written to ${OUT_FILE}`)

  expect(runs, 'expected five measured runs').toHaveLength(RUNS)
  expect(medians.t_load, 'load event timing not observable').toBeGreaterThan(0)
  expect(medians.t_fcp, 'first contentful paint not observable').toBeGreaterThan(0)
  expect(medians.t_interactive, 'primary action never became interactive').toBeGreaterThan(0)
  expect(existsSync(OUT_FILE), 'baseline JSON not written').toBe(true)
})
