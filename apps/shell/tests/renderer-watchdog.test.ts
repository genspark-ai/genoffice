import { mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

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
// every report also prunes the userData directory the reports live in — the
// mocked getPath points at a per-test temp dir, so a simulated report never
// touches the real /tmp (or deletes a real report pair from it)
let userDataDir = ''

vi.mock('electron', () => ({
  app: {
    getAppMetrics: () => metrics,
    getPath: () => userDataDir,
  },
  dialog: { showMessageBox: (...args: unknown[]) => showMessageBox(...args) },
  webContents: { getAllWebContents: () => allWebContents },
}))
// report() writes its .json through node:fs/promises and prunes through
// node:fs readdirSync/unlinkSync: both are stood in/wrapped so a shared log
// can pin their ORDER — the prune must observe the write, not race it
const ioLog: string[] = []
vi.mock('node:fs/promises', () => ({
  writeFile: vi.fn(async (path: unknown) => {
    ioLog.push(`write:${String(path)}`)
  }),
}))
vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    readdirSync: (path: unknown, ...rest: unknown[]) => {
      if (String(path).includes('watchdog-userdata')) ioLog.push(`readdir:${String(path)}`)
      return (actual.readdirSync as (p: unknown, ...r: unknown[]) => unknown)(path, ...rest)
    },
  }
})

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
  userDataDir = mkdtempSync(join(tmpdir(), 'watchdog-userdata-'))
  ioLog.length = 0
  ;({ startRendererWatchdog } = await import('../src/main/renderer-watchdog'))
})

afterEach(() => {
  rmSync(userDataDir, { recursive: true, force: true })
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

  it('every report prunes the directory it reports into, after writing', async () => {
    // the write is mocked (real file I/O cannot resolve inside a fake-timer
    // flush), so the disk holds only the fifteen seeded pairs — but the
    // shared ioLog pins the ORDER the code cannot fake: the report's write
    // must be logged before the prune's readdir, and a wiring that pruned
    // before the write (or from sample()) would log them backwards
    for (let stamp = 2_000; stamp < 2_015; stamp++) {
      writeFileSync(join(userDataDir, `renderer-watchdog-${stamp}.json`), '{}')
      if (stamp % 2 === 0) {
        writeFileSync(join(userDataDir, `renderer-watchdog-${stamp}.cpuprofile`), '[]')
      }
    }
    const stop = startWatchdog()
    await vi.advanceTimersByTimeAsync(2_100)
    stop()
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    // write first, prune second — exactly one of each for one report
    const reportWrite = ioLog.findIndex((entry) => entry.startsWith('write:'))
    const pruneReaddir = ioLog.findIndex((entry) => entry.startsWith('readdir:'))
    expect(reportWrite).toBeGreaterThanOrEqual(0)
    expect(pruneReaddir).toBeGreaterThanOrEqual(0)
    expect(reportWrite).toBeLessThan(pruneReaddir)
    expect(ioLog.filter((entry) => entry.startsWith('readdir:'))).toHaveLength(1)
    // and the prune kept the newest ten pairs, profiles riding with stamps:
    // 2005 is the boundary survivor, 2004 and below are gone
    const everything = readdirSync(userDataDir)
    expect(everything.filter((n) => n.endsWith('.json'))).toHaveLength(10)
    expect(everything).toContain('renderer-watchdog-2014.json')
    expect(everything).toContain('renderer-watchdog-2010.cpuprofile')
    expect(everything).toContain('renderer-watchdog-2005.json')
    expect(everything).not.toContain('renderer-watchdog-2004.json')
    expect(everything).not.toContain('renderer-watchdog-2000.cpuprofile')
  })
})
