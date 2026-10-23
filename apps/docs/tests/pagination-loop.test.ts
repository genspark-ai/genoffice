import { describe, expect, it } from 'vitest'
import { FollowUpChain, MAX_FOLLOW_UP_PASSES } from '../src/renderer/pagination-loop'

describe('FollowUpChain', () => {
  it('runs distinct reasons and a reason that re-fires a few times', () => {
    const c = new FollowUpChain()
    expect(c.request('width', null)).toBe('run')
    expect(c.request('rowFills', null)).toBe('run')
    expect(c.request('width', null)).toBe('run')
    expect(c.request('width', null)).toBe('run')
    expect(c.stops).toBe(0)
  })

  it('refuses a reason past its repeat budget and marks the log', () => {
    const c = new FollowUpChain()
    for (let i = 0; i < 3; i++) {
      expect(c.request('rowFills', null)).toBe('run')
      c.beginPass()
    }
    expect(c.request('rowFills', null)).toBe('repeat')
    expect(c.stops).toBe(1)
    expect(c.log.at(-1)).toBe('rowFills!repeat')
  })

  it('keys a shrink request on its paragraph', () => {
    const c = new FollowUpChain()
    for (let i = 0; i < 3; i++) {
      expect(c.request('justify-shrink', 10)).toBe('run')
      c.beginPass()
    }
    expect(c.request('justify-shrink', 11)).toBe('run')
    expect(c.request('justify-shrink', 10)).toBe('repeat')
  })

  it('re-asking before the pass runs is one request', () => {
    const c = new FollowUpChain()
    for (let i = 0; i < 10; i++) expect(c.request('justify-shrink', 0)).toBe('run')
    c.beginPass()
    expect(c.request('justify-shrink', 0)).toBe('run')
    expect(c.stops).toBe(0)
  })

  it('two decisions undoing each other exhaust their budgets', () => {
    const c = new FollowUpChain()
    const verdicts: string[] = []
    for (let i = 0; i < 8; i++) {
      verdicts.push(c.request(i % 2 ? 'b' : 'a', null))
      c.beginPass()
    }
    expect(verdicts).toEqual(['run', 'run', 'run', 'run', 'run', 'run', 'cap', 'cap'])
  })

  it('a refused cause stays refused even when asked again', () => {
    const c = new FollowUpChain()
    for (let i = 0; i < 3; i++) {
      c.request('rowFills', null)
      c.beginPass()
    }
    expect(c.request('rowFills', null)).toBe('repeat')
    expect(c.request('rowFills', null)).toBe('repeat')
    expect(c.stops).toBe(1)
    expect(c.log.slice(-2)).toEqual(['rowFills!repeat', 'rowFills!repeat'])
  })

  it('a refused cause does not unblock through the pending set', () => {
    const c = new FollowUpChain()
    c.passes = MAX_FOLLOW_UP_PASSES
    expect(c.request('x', 0)).toBe('cap')
    expect(c.request('x', 0)).toBe('cap')
  })

  it('caps the pass count as a backstop', () => {
    const c = new FollowUpChain()
    c.passes = MAX_FOLLOW_UP_PASSES
    expect(c.request('fresh', null)).toBe('cap')
  })

  it('reset starts a new chain', () => {
    const c = new FollowUpChain()
    for (let i = 0; i < 4; i++) {
      c.request('x', null)
      c.beginPass()
    }
    c.reset()
    expect(c.request('x', null)).toBe('run')
    expect(c.log).toEqual(['x'])
    expect(c.stops).toBe(0)
  })
})
