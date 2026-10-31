import {
  mkdirSync,
  mkdtempSync,
  readdirSync,
  rmSync,
  statSync,
  utimesSync,
  writeFileSync,
} from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * listRuns (src/store.ts) used to read and parse every run file on every
 * call, and nothing ever deleted one (issue #27 items 4+8). The read-window
 * half is invisible to any functional output — an unread file and a skipped
 * file look the same — so this file wraps readFileSync with a counter, and
 * readdirSync with another (a prune pass is one readdir + N stats, so a
 * readdir is the observable "a pass ran"); the retention half is plainly
 * observable on disk. Retention compares against the real clock, so every
 * mtime here is relative to now.
 */
const runFileReads = vi.fn()
const runDirReads = vi.fn()
vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    readFileSync: (path: unknown, ...rest: unknown[]) => {
      if (/[/\\]runs[/\\][^/\\]+\.jsonl$/.test(String(path))) runFileReads()
      return (actual.readFileSync as (p: unknown, ...r: unknown[]) => unknown)(path, ...rest)
    },
    readdirSync: (path: unknown, ...rest: unknown[]) => {
      if (/[/\\]runs$/.test(String(path).replace(/[/\\]+$/, ''))) runDirReads()
      return (actual.readdirSync as (p: unknown, ...r: unknown[]) => unknown)(path, ...rest)
    },
    unlinkSync: (path: unknown) => {
      if (String(path).endsWith('run-locked.jsonl')) {
        throw Object.assign(new Error('EACCES: unlink denied (simulated)'), { code: 'EACCES' })
      }
      return (actual.unlinkSync as (p: unknown) => void)(path)
    },
  }
})

type ProjectStore = typeof import('../src/store.js').ProjectStore
let store: InstanceType<ProjectStore>
let tmpDir: string
let runsDir: string
// the on-disk assertions read the directory through the real readdir, so
// they never pollute the readdir counter that pins the prune cadence
let realReaddir: typeof readdirSync

/** a folded run: start with an instruction, finish done, both at mtime's ts */
function writeRunFile(index: number, mtime: Date): void {
  const path = join(runsDir, `run-${String(index).padStart(4, '0')}.jsonl`)
  const ts = mtime.toISOString()
  writeFileSync(
    path,
    `{"kind":"start","ts":"${ts}","instruction":"run ${index}"}\n` +
      `{"kind":"finish","ts":"${ts}","outcome":"done"}\n`,
  )
  utimesSync(path, mtime, mtime)
}

/**
 * A run whose start timestamp and file mtime diverge — a long run overlapping
 * the newest batch, a dismissed row whose finish just landed in a months-old
 * file, or (reversed) the boundary case below. The read window ranks files
 * by mtime while the result ranks runs by startedAt; these are the files
 * that make the two orderings disagree.
 */
function writeDivergentRun(id: string, startedAt: Date, mtime: Date): void {
  const path = join(runsDir, `${id}.jsonl`)
  const ts = startedAt.toISOString()
  writeFileSync(
    path,
    `{"kind":"start","ts":"${ts}","instruction":"${id}"}\n` +
      `{"kind":"finish","ts":"${ts}","outcome":"done"}\n`,
  )
  utimesSync(path, mtime, mtime)
}

/** run files on disk right now, sorted by name */
function runFilesOnDisk(): string[] {
  return realReaddir(runsDir).sort()
}

beforeEach(async () => {
  const { ProjectStore: Store } = await import('../src/store.js')
  realReaddir = ((await vi.importActual('node:fs')) as typeof import('node:fs')).readdirSync
  runFileReads.mockReset()
  runDirReads.mockReset()
  tmpDir = mkdtempSync(join(tmpdir(), 'project-store-retention-'))
  store = new Store(tmpDir)
  store.ensureDefaultProject()
  runsDir = join(tmpDir, 'projects', 'default', 'runs')
  mkdirSync(runsDir, { recursive: true })
})

afterEach(() => {
  rmSync(tmpDir, { recursive: true, force: true })
})

describe('listRuns read window (issue 27 item 4)', () => {
  it('reads newest-first and stops once the listed window is full', () => {
    // 300 accumulated runs: the old listRuns read all 300 on every Ask-panel
    // mount. The window reads the newest batch-wise and the boundary stop
    // lands here after 13 batches — 208 of 300 files, within one batch
    // granularity of the window filling. What is pinned is that the bound
    // sits far below the directory's size.
    const now = Date.now()
    for (let i = 0; i < 300; i++) {
      writeRunFile(i, new Date(now - (300 - i) * 60_000))
    }
    const runs = store.listRuns('default')
    expect(runs).toHaveLength(200)
    // newest 200 by startedAt: runs 100..299, newest first
    expect(runs[0]).toMatchObject({ instruction: 'run 299' })
    expect(runs[199]).toMatchObject({ instruction: 'run 100' })
    expect(runFileReads.mock.calls.length).toBeLessThanOrEqual(210)
    expect(runFileReads.mock.calls.length).toBeLessThan(300)
  })

  it('keeps the run a fresh-mtime stale-start cohort would displace', () => {
    // the window ranks files by mtime, the result ranks runs by startedAt.
    // Nine runs with three-day-old starts but fresh mtimes (long runs, or
    // dismissed rows whose finish events just landed in old files) sit at
    // the head of the mtime order. run-x ranks 150th of 209 by startedAt —
    // comfortably inside the newest 200 — but its mtime is the oldest in
    // the directory: 209th of 209, so a window that stops once 200 summaries
    // are folded (208 files read) never reaches it. The old all-files
    // listing kept run-x and dropped every Y; a count-stop window drops
    // run-x and lists the newest Y in its place.
    const now = Date.now()
    const day = 24 * 60 * 60 * 1000
    for (let i = 0; i < 199; i++) {
      const at = new Date(now - (200 - i) * 60_000)
      writeDivergentRun(`run-s${String(i).padStart(3, '0')}`, at, at)
    }
    for (let i = 0; i < 9; i++) {
      writeDivergentRun(`run-y${i}`, new Date(now - 3 * day), new Date(now - (9 - i) * 1000))
    }
    writeDivergentRun(
      'run-x',
      new Date(now - (150 * 60_000 + 30_000)),
      new Date(now - 220 * 60_000),
    )
    const instructions = store.listRuns('default').map((r) => r.instruction)
    // 199 S runs + run-x is exactly the newest 200 by startedAt
    expect(instructions).toHaveLength(200)
    expect(instructions).toContain('run-x')
    // the stale-start cohort sorts below every S run: none made the list
    expect(instructions).not.toContain('run-y4')
    // and the result is still ordered by startedAt, newest first
    expect(instructions[0]).toBe('run-s198')
    expect(instructions[199]).toBe('run-s000')
  })

  it('an unparsable startedAt fails the window open: the read runs to the end', () => {
    // a run whose start ts cannot be parsed makes every later boundary NaN;
    // the stop rule must refuse to stop on it (fail open) rather than cut
    // the read short — the files ranked behind the garbage run still hold
    // runs that belong in the list
    const now = Date.now()
    const day = 24 * 60 * 60 * 1000
    for (let i = 0; i < 199; i++) {
      const at = new Date(now - (200 - i) * 60_000)
      writeDivergentRun(`run-s${String(i).padStart(3, '0')}`, at, at)
    }
    for (let i = 0; i < 9; i++) {
      writeDivergentRun(`run-y${i}`, new Date(now - 3 * day), new Date(now - (9 - i) * 1000))
    }
    writeDivergentRun(
      'run-x',
      new Date(now - (150 * 60_000 + 30_000)),
      new Date(now - 220 * 60_000),
    )
    // the newest mtime in the directory, so the first batch reads it and
    // every boundary computed after it is NaN
    const garbage = join(runsDir, 'run-garbage.jsonl')
    writeFileSync(
      garbage,
      '{"kind":"start","ts":"not-a-date","instruction":"run-garbage"}\n' +
        '{"kind":"finish","ts":"not-a-date","outcome":"done"}\n',
    )
    utimesSync(garbage, new Date(now), new Date(now))

    const instructions = store.listRuns('default').map((r) => r.instruction)
    // the read ran to the end, so the mtime-oldest boundary run still made
    // the list — a boundary that treated NaN as "stop now" would miss it
    expect(instructions).toContain('run-x')
    expect(instructions).toContain('run-garbage')
  })
})

describe('run retention (issue 27 item 8)', () => {
  it('deletes runs older than the retention window, keeping the fresh ones', () => {
    const now = Date.now()
    const stale = new Date(now - 40 * 24 * 60 * 60 * 1000)
    for (let i = 0; i < 10; i++) {
      writeRunFile(i, new Date(now - (i + 1) * 60_000))
    }
    for (let i = 10; i < 50; i++) {
      writeRunFile(i, stale)
    }
    const runs = store.listRuns('default')
    // 40 stale files pruned from disk, 10 fresh kept and listed
    expect(runFilesOnDisk()).toHaveLength(10)
    expect(runs).toHaveLength(10)
  })

  it('caps the directory at the newest 500 once the grace window has passed', () => {
    // 600 runs, the newest 500 from the last hour and the oldest 100 from
    // yesterday: age retention keeps them all, but count retention prunes
    // yesterday's tail — files inside the grace window are spared (see the
    // in-flight test), files past it are not
    const now = Date.now()
    const pastGrace = 25 * 60 * 60 * 1000
    for (let i = 0; i < 100; i++) {
      writeRunFile(i, new Date(now - pastGrace - (100 - i) * 1000))
    }
    for (let i = 100; i < 600; i++) {
      writeRunFile(i, new Date(now - (600 - i) * 1000))
    }
    const runs = store.listRuns('default')
    expect(runFilesOnDisk()).toHaveLength(500)
    // yesterday's runs are gone; the newest 200 of the survivors are listed
    expect(runFilesOnDisk()[0]).toBe('run-0100.jsonl')
    expect(runs).toHaveLength(200)
    expect(runs[0]).toMatchObject({ instruction: 'run 599' })
  })

  it('count retention does not delete a run that is still in flight', () => {
    // an in-flight run writes its start event and then goes quiet for hours;
    // 500 runs started behind it must not delete its file — the terminal
    // event would land in a start-less file that folds to nothing
    const now = Date.now()
    const quiet = new Date(now - 2 * 60 * 60 * 1000)
    const path = join(runsDir, 'run-inflight.jsonl')
    writeFileSync(
      path,
      `{"kind":"start","ts":"${quiet.toISOString()}","instruction":"in flight"}\n`,
    )
    utimesSync(path, quiet, quiet)
    for (let i = 0; i < 500; i++) {
      writeRunFile(i, new Date(quiet.getTime() + (i + 1) * 1000))
    }
    store.listRuns('default')
    // the grace window kept the quiet file past the count cap
    expect(runFilesOnDisk()).toHaveLength(501)
    expect(runFilesOnDisk()).toContain('run-inflight.jsonl')
  })

  it('a file that cannot be deleted is kept, and still listed', () => {
    const now = Date.now()
    for (let i = 0; i < 10; i++) writeRunFile(i, new Date(now - (i + 1) * 60_000))
    const stale = new Date(now - 40 * 24 * 60 * 60 * 1000)
    const locked = join(runsDir, 'run-locked.jsonl')
    writeFileSync(
      locked,
      `{"kind":"start","ts":"${stale.toISOString()}","instruction":"stale but locked"}\n`,
    )
    utimesSync(locked, stale, stale)
    const runs = store.listRuns('default')
    // the deletion failed: it warns, stays on disk, and stays listable
    expect(runFilesOnDisk()).toContain('run-locked.jsonl')
    expect(runs.map((r) => r.instruction)).toContain('stale but locked')
  })

  it('appends prune too, at most one readdir pass per interval', () => {
    // a tray-resident session that never opens the Ask panel still grows the
    // runs directory; the first append after the interval prunes right after
    // its write — yesterday's expired tail goes now, and the newest file is
    // of course kept
    const now = Date.now()
    const pastGrace = 25 * 60 * 60 * 1000
    for (let i = 0; i < 100; i++) {
      writeRunFile(i, new Date(now - pastGrace - (100 - i) * 1000))
    }
    for (let i = 100; i < 600; i++) {
      writeRunFile(i, new Date(now - (600 - i) * 1000))
    }
    runDirReads.mockReset()
    store.appendRunEvent('default', 'run-zzz', { kind: 'start', instruction: 'fresh' })
    // the expired 100 are gone; the recent burst's oldest is inside the grace
    // window, so it survives the count cap — 501 files, cap exact in effect
    expect(runFilesOnDisk()).toHaveLength(501)
    expect(runFilesOnDisk()).toContain('run-zzz.jsonl')
    // the readdir count is the prune pass made visible: the first append paid
    // exactly one
    expect(runDirReads.mock.calls.length).toBe(1)
    // the second event of the same run is inside the interval: no pass at all
    store.appendRunEvent('default', 'run-zzz', { kind: 'finish', outcome: 'done' })
    expect(runDirReads.mock.calls.length).toBe(1)
    expect(statSync(join(runsDir, 'run-zzz.jsonl')).size).toBeGreaterThan(0)
    // and the guard — not some accident — is what kept the pass free: past
    // the interval the pass is paid again (RUNS_PRUNE_INTERVAL_MS is 1h; the
    // extra minute swallows the wall clock the fixture setup itself spent)
    const pastInterval = vi.spyOn(Date, 'now').mockReturnValue(now + 3_600_000 + 60_000)
    try {
      store.appendRunEvent('default', 'run-bbb', { kind: 'start', instruction: 'second' })
    } finally {
      pastInterval.mockRestore()
    }
    expect(runDirReads.mock.calls.length).toBe(2)
    expect(runFilesOnDisk()).toContain('run-bbb.jsonl')
  })
})
