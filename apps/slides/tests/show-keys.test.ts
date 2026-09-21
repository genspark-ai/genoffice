import { describe, expect, it } from 'vitest'
import {
  gotoPosition,
  INITIAL_SHOW_KEYS,
  pushVisited,
  reduceAudienceNav,
  reduceShowKey,
  toggleScreen,
  type ShowKeyState,
} from '../src/renderer/show-keys'

/** Feed keys in sequence, returning the final state + the action of the last key */
function run(keys: string[], start: ShowKeyState = INITIAL_SHOW_KEYS) {
  let state = start
  let last: ReturnType<typeof reduceShowKey> = null
  for (const k of keys) {
    last = reduceShowKey(state, k)
    if (last) state = last.state
  }
  return { state, action: last?.action ?? null }
}

describe('slide show keys', () => {
  it('maps PowerPoint navigation keys', () => {
    expect(run([' ']).action).toEqual({ type: 'next' })
    expect(run(['ArrowRight']).action).toEqual({ type: 'next' })
    expect(run(['PageDown']).action).toEqual({ type: 'next' })
    expect(run(['Enter']).action).toEqual({ type: 'next' })
    expect(run(['ArrowLeft']).action).toEqual({ type: 'prev' })
    expect(run(['PageUp']).action).toEqual({ type: 'prev' })
    expect(run(['Home']).action).toEqual({ type: 'first' })
    expect(run(['End']).action).toEqual({ type: 'last' })
    expect(run(['Escape']).action).toEqual({ type: 'exit' })
    expect(reduceShowKey(INITIAL_SHOW_KEYS, 'x')).toBeNull()
    expect(reduceShowKey(INITIAL_SHOW_KEYS, 'F5')).toBeNull()
  })

  it('B and . toggle black, W and , toggle white', () => {
    expect(run(['b']).state.screen).toBe('black')
    expect(run(['B']).state.screen).toBe('black')
    expect(run(['.']).state.screen).toBe('black')
    expect(run(['w']).state.screen).toBe('white')
    expect(run([',']).state.screen).toBe('white')
    expect(run(['b', 'b']).state.screen).toBe('none')
    expect(run(['w', 'w']).state.screen).toBe('none')
    expect(run(['b']).action).toEqual({ type: 'none' })
  })

  it('a blacked-out screen swallows the next key instead of navigating', () => {
    const r = run(['b', ' '])
    expect(r.state.screen).toBe('none')
    expect(r.action).toEqual({ type: 'none' })
    expect(run(['w', 'ArrowRight']).action).toEqual({ type: 'none' })
    expect(run(['b', '3']).state.digits).toBe('')
  })

  it('Escape still ends the show from a blackout', () => {
    expect(run(['b', 'Escape']).action).toEqual({ type: 'exit' })
  })

  it('digits buffer silently and Enter jumps to that slide number', () => {
    const typed = run(['1', '2'])
    expect(typed.action).toEqual({ type: 'none' })
    expect(typed.state.digits).toBe('12')
    const r = run(['Enter'], typed.state)
    expect(r.action).toEqual({ type: 'goto', slideNumber: 12 })
    expect(r.state.digits).toBe('')
  })

  it('Escape clears the digit buffer without exiting', () => {
    const r = run(['4', 'Escape'])
    expect(r.action).toEqual({ type: 'none' })
    expect(r.state.digits).toBe('')
    expect(run(['Escape'], r.state).action).toEqual({ type: 'exit' })
  })

  it('navigation and blackout keys drop pending digits', () => {
    expect(run(['7', 'ArrowRight']).state.digits).toBe('')
    expect(run(['7', 'ArrowRight']).action).toEqual({ type: 'next' })
    expect(run(['7', 'b']).state.digits).toBe('')
  })

  it('toggleScreen from the menu mirrors the key toggles', () => {
    const black = toggleScreen(INITIAL_SHOW_KEYS, 'black')
    expect(black.screen).toBe('black')
    expect(toggleScreen(black, 'black').screen).toBe('none')
    expect(toggleScreen(black, 'white').screen).toBe('white')
  })
})

describe('gotoPosition', () => {
  // deck of 5: slides 2 and 5 hidden
  const order = [0, 2, 3]

  it('uses deck numbering, not play-order numbering', () => {
    expect(gotoPosition(1, order)).toBe(0)
    expect(gotoPosition(3, order)).toBe(1)
    expect(gotoPosition(4, order)).toBe(2)
  })

  it("a hidden slide's number lands on the next slide the show plays", () => {
    expect(gotoPosition(2, order)).toBe(1)
  })

  it('past the end lands on the last shown slide even when the deck ends hidden', () => {
    expect(gotoPosition(5, order)).toBe(2)
    expect(gotoPosition(99, order)).toBe(2)
  })

  it('rejects zero, NaN and an empty show', () => {
    expect(gotoPosition(0, order)).toBeNull()
    expect(gotoPosition(Number.NaN, order)).toBeNull()
    expect(gotoPosition(1, [])).toBeNull()
  })

  it('follows a custom show order by the nearest following deck index', () => {
    expect(gotoPosition(2, [4, 1, 3])).toBe(1)
    expect(gotoPosition(3, [4, 1, 3])).toBe(2)
    expect(gotoPosition(5, [4, 1, 3])).toBe(0)
  })
})

describe('reduceAudienceNav', () => {
  it('relays navigation while the slide is visible', () => {
    expect(reduceAudienceNav(INITIAL_SHOW_KEYS, 'next').action).toEqual({ type: 'next' })
    expect(reduceAudienceNav(INITIAL_SHOW_KEYS, 'prev').action).toEqual({ type: 'prev' })
  })

  it('restores a blacked-out screen instead of advancing', () => {
    const black = toggleScreen(INITIAL_SHOW_KEYS, 'black')
    const r = reduceAudienceNav(black, 'next')
    expect(r.action).toEqual({ type: 'none' })
    expect(r.state.screen).toBe('none')
    expect(reduceAudienceNav(toggleScreen(INITIAL_SHOW_KEYS, 'white'), 'prev').action).toEqual({
      type: 'none',
    })
  })

  it('always lets the audience end the show', () => {
    const black = toggleScreen(INITIAL_SHOW_KEYS, 'black')
    expect(reduceAudienceNav(black, 'exit').action).toEqual({ type: 'exit' })
  })
})

describe('pushVisited', () => {
  it('records a stack without consecutive duplicates', () => {
    expect(pushVisited([], 0)).toEqual([0])
    expect(pushVisited([0], 0)).toEqual([0])
    expect(pushVisited([0], 2)).toEqual([0, 2])
    expect(pushVisited([0, 2], 0)).toEqual([0, 2, 0])
  })
})
