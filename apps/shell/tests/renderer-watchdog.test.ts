import { beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * startRendererWatchdog (src/main/renderer-watchdog.ts): a big-but-stable
 * document used to redraw the warning dialog on every cooldown forever
 * (field report: a ~1.9 GB CSV re-prompted every quarter hour). After the
 * first prompt a memory-triggered report must repeat only when the renderer
 * grew ~25% past the memory at the previous prompt, a CPU-triggered report
 * keeps repeating (an actively burning core), and the new third button
 * silences the renderer for the session.
 */

const showMessageBox = vi.fn<(...args: unknown[]) => Promise<{ response: number }>>()
let metrics: unknown[] = []
let allWebContents: unknown[] = []

vi.mock('electron', () => ({
  app: {
    getAppMetrics: () => metrics,
    getPath: () => '/tmp',
  },
  dialog: { showMessageBox: (...args: unknown[]) => showMessageBox(...args) },
  webContents: { getAllWebContents: () => allWebContents },
}))
vi.mock('node:fs/promises', () => ({ writeFile: vi.fn(async () => {}) }))

let startRendererWatchdog: typeof import('../src/main/renderer-watchdog').startRendererWatchdog

function fakeWc(id: number, pid: number) {
  return {
    id,
    isDestroyed: () => false,
    getOSProcessId: () => pid,
    getURL: () => 'genoffice-app://sheets/index.html',
    debugger: { isAttached: () => true },
  }
}

function tabMetric(pid: number, memMB: number, cpuSeconds: number) {
  return {
    type: 'Tab',
    pid,
    memory: { workingSetSize: memMB * 1024 },
    cpu: { cumulativeCPUUsage: cpuSeconds },
  }
}

function startWatchdog(closeDocument = vi.fn()) {
  return startRendererWatchdog({
    describe: () => ({ kind: 'sheets', title: 'big.csv' }),
    parentWindow: () => null,
    closeDocument,
    t: (key) => key,
    sampleMs: 1_000,
    memoryLimitMB: 1_000,
    cpuLimitPercent: 90,
    hotSamplesBeforeReport: 2,
    reportCooldownMs: 5_000,
    profileMs: 0,
  })
}

beforeEach(async () => {
  vi.useFakeTimers()
  showMessageBox.mockReset()
  showMessageBox.mockResolvedValue({ response: 0 })
  allWebContents = [fakeWc(1, 10)]
  metrics = [tabMetric(10, 2_000, 0)]
  ;({ startRendererWatchdog } = await import('../src/main/renderer-watchdog'))
})

describe('renderer watchdog re-prompt policy', () => {
  it('prompts once a renderer stays over the memory limit, with the ignore button', async () => {
    const stop = startWatchdog()
    await vi.advanceTimersByTimeAsync(2_100)
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    const options = showMessageBox.mock.calls[0]![0] as { buttons: string[] }
    expect(options.buttons).toEqual(['watchdogWait', 'watchdogClose', 'watchdogIgnore'])
    stop()
  })

  it('does not re-prompt on flat memory, re-prompts after ~25% growth', async () => {
    const stop = startWatchdog()
    await vi.advanceTimersByTimeAsync(20_000)
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    metrics = [tabMetric(10, 2_600, 0)]
    await vi.advanceTimersByTimeAsync(3_000)
    expect(showMessageBox).toHaveBeenCalledTimes(2)
    stop()
  })

  it('the third button silences the renderer for the session', async () => {
    showMessageBox.mockResolvedValue({ response: 2 })
    const stop = startWatchdog()
    await vi.advanceTimersByTimeAsync(2_100)
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    metrics = [tabMetric(10, 8_000, 0)]
    await vi.advanceTimersByTimeAsync(30_000)
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    stop()
  })

  it('a CPU-hot renderer keeps re-prompting on every cooldown', async () => {
    // cumulative CPU grows one second per one-second sample: a pinned core
    let cpuSeconds = 0
    metrics = [tabMetric(10, 100, cpuSeconds)]
    const stop = startWatchdog()
    for (let tick = 0; tick < 12; tick += 1) {
      cpuSeconds += 1
      metrics = [tabMetric(10, 100, cpuSeconds)]
      await vi.advanceTimersByTimeAsync(1_000)
    }
    expect(showMessageBox.mock.calls.length).toBeGreaterThanOrEqual(2)
    stop()
  })

  it('close button routes through closeDocument', async () => {
    showMessageBox.mockResolvedValue({ response: 1 })
    const closeDocument = vi.fn()
    const stop = startWatchdog(closeDocument)
    await vi.advanceTimersByTimeAsync(2_100)
    expect(closeDocument).toHaveBeenCalledTimes(1)
    stop()
  })
})
