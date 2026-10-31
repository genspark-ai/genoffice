import { describe, expect, it } from 'vitest'

import {
  normalizeAskDeepLink,
  normalizeScopePaths,
  planPendingAskQuery,
  type AskQueryOrigin,
  type PendingAskQuery,
} from '../src/renderer/src/ask-intent'

/**
 * The ask-intent handoff: the palette's Ask AI fallback and the
 * `home:deep-link` bridge both route a query into Home's ask view, and the
 * panel consumes each handoff exactly once (seq-guarded). The origin decides
 * the submit policy: the palette's user-typed text submits itself, a deep
 * link's web-supplied text only ever parks in the composer.
 *
 * A handoff may carry scopePaths (the desktop's multi-select): the round
 * scope rides the plan — a question plus a selection submits scoped, a
 * selection with no text parks the scope alone (the panel opens, the chip
 * attaches, nothing sends itself).
 */

const pending = (text: string, seq = 1, origin: AskQueryOrigin = 'palette'): PendingAskQuery => ({
  text,
  seq,
  origin,
})

const scoped = (
  text: string,
  scopePaths: readonly string[],
  seq = 1,
  origin: AskQueryOrigin = 'palette',
): PendingAskQuery => ({ text, seq, origin, scopePaths })

describe('planPendingAskQuery', () => {
  it('does nothing without a pending query or for an already-consumed seq', () => {
    expect(planPendingAskQuery(null, 0, false)).toEqual({ consume: false })
    expect(planPendingAskQuery(pending('hello', 3), 3, false)).toEqual({ consume: false })
  })

  it('submits a fresh palette query when the panel is idle', () => {
    expect(planPendingAskQuery(pending('hello', 1), 0, false)).toEqual({
      consume: true,
      seq: 1,
      action: 'submit',
      text: 'hello',
    })
  })

  it('parks a palette query in the composer while an answer streams, instead of dropping it', () => {
    expect(planPendingAskQuery(pending('hello', 1), 0, true)).toEqual({
      consume: true,
      seq: 1,
      action: 'park',
      text: 'hello',
    })
  })

  it('parks a deep-link query even while the panel is idle: web-supplied text never sends itself', () => {
    expect(planPendingAskQuery(pending('hello', 1, 'deep-link'), 0, false)).toEqual({
      consume: true,
      seq: 1,
      action: 'park',
      text: 'hello',
    })
    expect(planPendingAskQuery(pending('hello', 2, 'deep-link'), 1, true)).toEqual({
      consume: true,
      seq: 2,
      action: 'park',
      text: 'hello',
    })
  })

  it('trims the text and drops a blank one (consumed, so it never repeats)', () => {
    expect(planPendingAskQuery(pending('  hello  ', 2), 1, false)).toEqual({
      consume: true,
      seq: 2,
      action: 'submit',
      text: 'hello',
    })
    expect(planPendingAskQuery(pending('   ', 2), 1, false)).toEqual({
      consume: true,
      seq: 2,
      action: 'drop',
      text: '',
    })
    expect(planPendingAskQuery(pending('   ', 3, 'deep-link'), 2, false)).toEqual({
      consume: true,
      seq: 3,
      action: 'drop',
      text: '',
    })
  })

  it('the same text twice is two submissions, because the seq differs', () => {
    const first = planPendingAskQuery(pending('again', 1), 0, false)
    const second = planPendingAskQuery(pending('again', 2), 1, false)
    expect(first.consume && first.action).toBe('submit')
    expect(second.consume && second.action).toBe('submit')
  })
})

describe('planPendingAskQuery with a round scope', () => {
  const paths = ['/desk/a.docx', '/desk/b.xlsx']

  it('a question plus a selection submits scoped while the panel is idle', () => {
    expect(planPendingAskQuery(scoped('summarize these', paths), 0, false)).toEqual({
      consume: true,
      seq: 1,
      action: 'submit',
      text: 'summarize these',
      scopePaths: paths,
    })
  })

  it('a busy panel parks the question and keeps the scope for the send that follows', () => {
    expect(planPendingAskQuery(scoped('summarize these', paths), 0, true)).toEqual({
      consume: true,
      seq: 1,
      action: 'park',
      text: 'summarize these',
      scopePaths: paths,
    })
  })

  it('blank text with a selection parks instead of dropping: the chip alone is the handoff', () => {
    expect(planPendingAskQuery(scoped('', paths), 0, false)).toEqual({
      consume: true,
      seq: 1,
      action: 'park',
      text: '',
      scopePaths: paths,
    })
    // web-supplied text never sends itself, with or without a scope
    expect(planPendingAskQuery(scoped('', paths, 2, 'deep-link'), 1, false)).toEqual({
      consume: true,
      seq: 2,
      action: 'park',
      text: '',
      scopePaths: paths,
    })
  })

  it('a handoff whose scope sanitizes to nothing behaves like a plain query', () => {
    expect(planPendingAskQuery(scoped('', ['', 42, null]), 0, false)).toEqual({
      consume: true,
      seq: 1,
      action: 'drop',
      text: '',
    })
    expect(planPendingAskQuery(scoped('hello', [42]), 0, false)).toEqual({
      consume: true,
      seq: 1,
      action: 'submit',
      text: 'hello',
    })
  })
})

describe('normalizeScopePaths', () => {
  it('keeps the non-empty strings, deduplicated in order', () => {
    expect(normalizeScopePaths(['/a.docx', '/a.docx', '/b.xlsx'])).toEqual(['/a.docx', '/b.xlsx'])
  })

  it('drops non-strings and empty strings rather than throwing', () => {
    expect(normalizeScopePaths([undefined, '/a.docx', '', 42, null])).toEqual(['/a.docx'])
  })

  it('a non-array answers empty, whatever junk it carries', () => {
    expect(normalizeScopePaths(undefined)).toEqual([])
    expect(normalizeScopePaths('/a.docx')).toEqual([])
    expect(normalizeScopePaths({ 0: '/a.docx' })).toEqual([])
  })

  it('caps at 50 paths, mirroring the main-side onlyPaths cap', () => {
    const many = Array.from({ length: 80 }, (_, i) => `/desk/${i}.docx`)
    expect(normalizeScopePaths(many)).toEqual(many.slice(0, 50))
  })
})

describe('normalizeAskDeepLink', () => {
  it('takes the ask payload, defaulting a missing or non-string query to empty', () => {
    expect(normalizeAskDeepLink({ view: 'ask', query: 'hello' })).toEqual({ query: 'hello' })
    expect(normalizeAskDeepLink({ view: 'ask' })).toEqual({ query: '' })
    expect(normalizeAskDeepLink({ view: 'ask', query: 42 })).toEqual({ query: '' })
  })

  it('ignores other views and malformed payloads', () => {
    expect(normalizeAskDeepLink({ view: 'recent' })).toBeNull()
    expect(normalizeAskDeepLink(null)).toBeNull()
    expect(normalizeAskDeepLink('ask')).toBeNull()
    expect(normalizeAskDeepLink(undefined)).toBeNull()
  })

  it('truncates an over-long query at 2000 characters', () => {
    expect(normalizeAskDeepLink({ view: 'ask', query: 'x'.repeat(2500) })).toEqual({
      query: 'x'.repeat(2000),
    })
    expect(normalizeAskDeepLink({ view: 'ask', query: 'y'.repeat(2000) }).query).toHaveLength(2000)
  })
})
