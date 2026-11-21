import { describe, expect, it } from 'vitest'
import {
  consumeLastSessionPlan,
  launchPathsWithLastSession,
  readLastSessionDocuments,
} from '../src/main/last-session'

const all = () => true

describe('readLastSessionDocuments', () => {
  it('keeps absolute string paths once, in order', () => {
    expect(
      readLastSessionDocuments({
        lastSessionDocuments: ['/a.docx', 'relative.docx', 7, null, '/b.xlsx', '/a.docx'],
      }),
    ).toEqual(['/a.docx', '/b.xlsx'])
  })

  it('treats a missing or malformed value as no documents', () => {
    expect(readLastSessionDocuments({})).toEqual([])
    expect(readLastSessionDocuments({ lastSessionDocuments: '/a.docx' })).toEqual([])
  })

  it('caps a hand-edited list', () => {
    const many = Array.from({ length: 500 }, (_, i) => `/doc-${i}.md`)
    expect(readLastSessionDocuments({ lastSessionDocuments: many })).toHaveLength(50)
  })
})

describe('launchPathsWithLastSession', () => {
  const saved = { lastSessionDocuments: ['/a.docx', '/gone.pptx', '/b.xlsx'] }

  it('opens only the launch paths once switched off', () => {
    const settings = { ...saved, restoreLastSession: false }
    expect(launchPathsWithLastSession(settings, ['/c.pdf'], all)).toEqual(['/c.pdf'])
  })

  it('restores by default: existing documents first, then the launch paths, without duplicates', () => {
    const settings = saved
    const exists = (p: string) => p !== '/gone.pptx'
    expect(launchPathsWithLastSession(settings, ['/b.xlsx', '/c.pdf'], exists)).toEqual([
      '/a.docx',
      '/b.xlsx',
      '/c.pdf',
    ])
  })
})

describe('consumeLastSessionPlan — the restore survives the launch (#2891)', () => {
  const saved = { lastSessionDocuments: ['/a.docx', '/gone.pptx', '/b.xlsx'] }
  const exists = (p: string) => p !== '/gone.pptx'

  it('a fresh launch restores and arms the restore, then disarms once opened', () => {
    const plan = consumeLastSessionPlan(saved, [], exists)
    expect(plan.paths).toEqual(['/a.docx', '/b.xlsx'])
    expect(plan.writes).toEqual([
      { lastSessionRestoreInFlight: true, lastSessionRestoreFailures: 0 },
    ])
    expect(plan.completedWrite).toEqual({
      lastSessionRestoreInFlight: false,
      lastSessionRestoreFailures: 0,
    })
    expect(plan.dropped).toBe(false)
    // nothing to restore: no writes at all, not even the arm
    const empty = consumeLastSessionPlan({ lastSessionDocuments: [] }, [], exists)
    expect(empty.writes).toEqual([])
    expect(empty.completedWrite).toBeNull()
  })

  it('the first crashed attempt retries once and counts the failure', () => {
    const crashed = { ...saved, lastSessionRestoreInFlight: true, lastSessionRestoreFailures: 0 }
    const plan = consumeLastSessionPlan(crashed, [], exists)
    // the session gets its second chance: same paths, failure recorded up front
    expect(plan.paths).toEqual(['/a.docx', '/b.xlsx'])
    expect(plan.writes).toEqual([{ lastSessionRestoreFailures: 1 }])
    expect(plan.dropped).toBe(false)
  })

  it('the second crashed attempt drops the list — the crash-loop guard', () => {
    const crashedTwice = {
      ...saved,
      lastSessionRestoreInFlight: true,
      lastSessionRestoreFailures: 1,
    }
    const plan = consumeLastSessionPlan(crashedTwice, [], exists)
    expect(plan.paths).toEqual([])
    expect(plan.dropped).toBe(true)
    expect(plan.writes).toEqual([
      {
        lastSessionDocuments: undefined,
        lastSessionRestoreInFlight: undefined,
        lastSessionRestoreFailures: 0,
      },
    ])
    expect(plan.completedWrite).toBeNull()
  })

  it('a crash after the documents opened costs nothing — the flag was already disarmed', () => {
    // the launch completed (flag cleared), then the app crashed an hour later:
    // the next launch restores the last clean close's set again
    const later = { ...saved, lastSessionRestoreInFlight: false, lastSessionRestoreFailures: 0 }
    const plan = consumeLastSessionPlan(later, [], exists)
    expect(plan.paths).toEqual(['/a.docx', '/b.xlsx'])
    expect(plan.dropped).toBe(false)
  })

  it('a retry with nothing left to restore clears the failed state without dropping', () => {
    const crashed = {
      lastSessionDocuments: ['/gone.pptx'],
      lastSessionRestoreInFlight: true,
      lastSessionRestoreFailures: 0,
    }
    const plan = consumeLastSessionPlan(crashed, [], exists)
    expect(plan.paths).toEqual([])
    expect(plan.writes).toEqual([
      { lastSessionRestoreInFlight: undefined, lastSessionRestoreFailures: 0 },
    ])
    expect(plan.dropped).toBe(false)
  })

  it('switched off, the restore stays off — even mid-failed-attempt', () => {
    const off = {
      ...saved,
      restoreLastSession: false,
      lastSessionRestoreInFlight: true,
      lastSessionRestoreFailures: 0,
    }
    const plan = consumeLastSessionPlan(off, ['/new.md'], exists)
    expect(plan.paths).toEqual(['/new.md'])
    expect(plan.dropped).toBe(false)
    expect(plan.writes).toEqual([
      {
        lastSessionDocuments: undefined,
        lastSessionRestoreInFlight: undefined,
        lastSessionRestoreFailures: 0,
      },
    ])
  })
})
