import { linkSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { refreshCsvDigest, type CsvDigestSession } from '../src/main/csv-digest'

type CsvSession = CsvDigestSession

/**
 * A CSV export that lands on the session's own source file must refresh that
 * session's guard digest.
 *
 * The session was opened with one spelling and the export writes another, and
 * macOS and Windows are case-insensitive by default — so the two names one
 * file. Matching on the string misses, the digest is left stale, and the save
 * guard that runs next (hashing content, so case-immune already) refuses the
 * next ⌘S as an external change to a file the user never touched.
 */
describe('refreshCsvDigest', () => {
  function scratch(): string {
    return mkdtempSync(join(tmpdir(), 'csv-digest-'))
  }

  const session = (over: CsvSession = {}): CsvSession => ({ ...over })

  it('refreshes when only the spelling differs', () => {
    const dir = scratch()
    // written once, as `Report.csv`, because the export dialog's default does
    const lower = join(dir, 'Report.csv')
    writeFileSync(lower, 'a,b\n1,2\n')
    // `Report.CSV` — the other spelling of the same file. A hard link where
    // the disk needs one: CI runs on a case-sensitive filesystem, where the
    // two spellings would be two files, while a link is one inode under two
    // names on every OS. On a case-insensitive disk (macOS, Windows — where
    // the bug lives) the lookup of `Report.CSV` already lands on the written
    // `Report.csv`, so linkSync raises EEXIST and there is nothing to link.
    const upper = join(dir, 'Report.CSV')
    try {
      linkSync(lower, upper)
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code !== 'EEXIST') throw e
    }

    const sessions = new Map<string, CsvSession>([
      ['s1', session({ csvSourcePath: upper, csvSourceSha: 'stale' })],
    ])
    refreshCsvDigest(sessions, lower, 'fresh')

    expect(sessions.get('s1')?.csvSourceSha).toBe('fresh')
  })

  it('leaves a session that was reading a different file alone', () => {
    const dir = scratch()
    const target = join(dir, 'Report.csv')
    const other = join(dir, 'Other.csv')
    writeFileSync(target, 'a\n')
    writeFileSync(other, 'b\n')

    const sessions = new Map<string, CsvSession>([
      ['s1', session({ csvSourcePath: other, csvSourceSha: 'keep' })],
    ])
    refreshCsvDigest(sessions, target, 'fresh')

    expect(sessions.get('s1')?.csvSourceSha).toBe('keep')
  })

  it('is a no-op when the export target does not exist', () => {
    const dir = scratch()
    const missing = join(dir, 'nope.csv')
    const sessions = new Map([['s1', session({ csvSourcePath: missing, csvSourceSha: 'keep' })]])
    refreshCsvDigest(sessions, missing, 'fresh')
    expect(sessions.get('s1')?.csvSourceSha).toBe('keep')
  })

  it('skips sessions with no CSV source at all', () => {
    const dir = scratch()
    const target = join(dir, 'Book.csv')
    writeFileSync(target, 'a\n')
    const sessions = new Map<string, CsvSession>([['s1', session({ csvSourceSha: 'keep' })]])
    refreshCsvDigest(sessions, target, 'fresh')
    expect(sessions.get('s1')?.csvSourceSha).toBe('keep')
  })
})
