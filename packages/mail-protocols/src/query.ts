/**
 * Local MailQuery post-filter for the P3 sync engine.
 *
 * Server-side coverage differs per driver (IMAP runs all ten fields through
 * its SEARCH builder; Gmail/Graph run the core text/from/since/onlyUnread
 * subset). matchesQuery applies the full ten-field MailQuery to an already
 * fetched MailEnvelope, so the extended fields (to/subject/seen/flagged/
 * hasAttach/before) are honored on every driver one way or the other.
 * Pure and deterministic; no I/O, no secrets.
 */
import type { MailEnvelope, MailQuery } from './types.js'

function includes(haystack: string, needle: string): boolean {
  return haystack.toLowerCase().includes(needle.toLowerCase())
}

function anyIncludes(values: string[], needle: string): boolean {
  return values.some((value) => includes(value, needle))
}

/** Day-granularity floor: IMAP SINCE/BEFORE compare calendar dates. */
function dayFloor(value: string): number | undefined {
  const time = new Date(value).getTime()
  if (Number.isNaN(time)) return undefined
  const day = new Date(time)
  day.setHours(0, 0, 0, 0)
  return day.getTime()
}

export function matchesQuery(env: MailEnvelope, query: MailQuery): boolean {
  if (query.text !== undefined && query.text.trim() !== '') {
    const hay = [...env.from, ...env.to, env.subject, env.snippet].join(' ')
    if (!includes(hay, query.text.trim())) return false
  }
  if (query.from !== undefined && query.from.trim() !== '') {
    if (!anyIncludes(env.from, query.from.trim())) return false
  }
  if (query.to !== undefined && query.to.trim() !== '') {
    if (!anyIncludes(env.to, query.to.trim())) return false
  }
  if (query.subject !== undefined && query.subject.trim() !== '') {
    if (!includes(env.subject, query.subject.trim())) return false
  }
  if (query.seen !== undefined && env.flags.seen !== query.seen) return false
  if (query.flagged !== undefined && env.flags.flagged !== query.flagged) return false
  if (query.hasAttach !== undefined && env.hasAttach !== query.hasAttach) return false
  if (query.since !== undefined && query.since.trim() !== '') {
    const floor = dayFloor(query.since)
    if (floor !== undefined) {
      const date = dayFloor(env.date)
      if (date === undefined || date < floor) return false
    }
  }
  if (query.before !== undefined && query.before.trim() !== '') {
    const floor = dayFloor(query.before)
    if (floor !== undefined) {
      const date = dayFloor(env.date)
      if (date === undefined || date >= floor) return false
    }
  }
  if (query.onlyUnread === true && env.flags.seen) return false
  return true
}
