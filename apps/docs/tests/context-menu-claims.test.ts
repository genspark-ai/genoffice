/**
 * The renderer claims each body right-click synchronously from its DOM
 * handler, before Blink requests a context menu for it, so the main process
 * sees claims and context-menu events in the same order and pairs them FIFO.
 * A claimed click is the React menu's (no native popup), an unclaimed one is
 * native; every claim is consumed by its event or expires.
 */
import { describe, expect, it } from 'vitest'
import { ClickClaims } from '../src/shared/context-menu-claims'

describe('ClickClaims', () => {
  it('claim then event: the event takes that claim, a second event is native', () => {
    const claims = new ClickClaims()
    claims.claim(1, 0)
    expect(claims.take(1)).toBe(1)
    expect(claims.take(2)).toBeNull()
  })

  it('event without a claim (header/footer surface, input): native', () => {
    const claims = new ClickClaims()
    expect(claims.take(0)).toBeNull()
  })

  it('overlapping clicks: events pair with their claims in order', () => {
    const claims = new ClickClaims()
    claims.claim(1, 0)
    claims.claim(2, 50)
    expect(claims.take(60)).toBe(1)
    expect(claims.take(70)).toBe(2)
    expect(claims.take(80)).toBeNull()
  })

  it('a body click followed by an input click: only the first is claimed', () => {
    const claims = new ClickClaims()
    claims.claim(1, 0)
    expect(claims.take(10)).toBe(1)
    expect(claims.take(20)).toBeNull()
  })

  it('a claim whose event never comes expires instead of hiding a later native menu', () => {
    const claims = new ClickClaims()
    claims.claim(1, 0)
    expect(claims.take(2001)).toBeNull()
    claims.claim(2, 5000)
    claims.claim(3, 5000)
    expect(claims.take(7500)).toBeNull()
  })
})
