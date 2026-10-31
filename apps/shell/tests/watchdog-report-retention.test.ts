import { mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * Retention for watchdog reports (issue #27 item 8): each report is a
 * renderer-watchdog-<stamp>.json plus an optional .cpuprofile pair, and
 * nothing ever removed them. pruneWatchdogReports keeps the newest stamps and
 * deletes every older pair — tested against a real directory; only the
 * electron module is stood in, because the module imports it at the top.
 */

vi.mock('electron', () => ({ app: {}, dialog: {}, webContents: {} }))

import { WATCHDOG_REPORTS_KEPT, pruneWatchdogReports } from '../src/main/renderer-watchdog'

let dir: string

/** one report pair: the .json always, the .cpuprofile only for profiled ones */
function writeReport(stamp: number, withProfile: boolean): void {
  writeFileSync(join(dir, `renderer-watchdog-${stamp}.json`), '{}')
  if (withProfile) writeFileSync(join(dir, `renderer-watchdog-${stamp}.cpuprofile`), '[]')
}

function stampedReports(): string[] {
  return readdirSync(dir)
    .filter((n) => /^renderer-watchdog-\d+\.json$/.test(n))
    .sort()
}

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'watchdog-reports-'))
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('watchdog report retention (issue 27 item 8)', () => {
  it('keeps the newest pairs and deletes every older one', () => {
    for (let stamp = 1_000; stamp < 1_015; stamp++) writeReport(stamp, stamp % 2 === 0)
    writeFileSync(join(dir, 'unrelated.txt'), 'stay')
    writeFileSync(join(dir, 'renderer-watchdog-not-a-stamp.json'), 'stay')

    pruneWatchdogReports(dir)

    // stamps 1005..1014 are the newest ten of the fifteen
    expect(stampedReports()[0]).toBe('renderer-watchdog-1005.json')
    expect(stampedReports()).toHaveLength(10)
    expect(readdirSync(dir)).not.toContain('renderer-watchdog-1004.json')
    // profiles ride with their stamps: a pruned one is gone, a kept one stays
    expect(readdirSync(dir)).not.toContain('renderer-watchdog-1002.cpuprofile')
    expect(readdirSync(dir)).toContain('renderer-watchdog-1010.cpuprofile')
    // foreign names are not ours to touch
    expect(readdirSync(dir)).toContain('unrelated.txt')
    expect(readdirSync(dir)).toContain('renderer-watchdog-not-a-stamp.json')
  })

  it('a protected stamp survives even when it ranks oldest', () => {
    // a backwards clock step can leave the report that just triggered the
    // prune numerically oldest: report() passes its own stamp through, and
    // the prune never deletes it — the doomed list skips it
    writeReport(999, true)
    for (let stamp = 1_000; stamp < 1_015; stamp++) writeReport(stamp, stamp % 2 === 0)
    pruneWatchdogReports(dir, 10, 999)
    const names = readdirSync(dir)
    // 999 is doomed by rank but protected; 1000..1004 are doomed and gone
    expect(names).toContain('renderer-watchdog-999.json')
    expect(names).toContain('renderer-watchdog-999.cpuprofile')
    expect(names).not.toContain('renderer-watchdog-1004.json')
    expect(names).not.toContain('renderer-watchdog-1000.json')
    // ten by rank plus the protected pair
    expect(names.filter((n) => n.endsWith('.json'))).toHaveLength(11)
  })

  it('a directory under the cap, or missing, is left alone without error', () => {
    writeReport(7, true)
    pruneWatchdogReports(dir)
    expect(readdirSync(dir)).toHaveLength(2)
    // a missing directory is a no-op, not an error
    expect(() => pruneWatchdogReports(join(dir, 'no-such-dir'))).not.toThrow()
  })

  it('the default keep count is the documented ten', () => {
    expect(WATCHDOG_REPORTS_KEPT).toBe(10)
  })
})
