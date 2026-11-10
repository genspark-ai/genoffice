import { describe, expect, it } from 'vitest'
import {
  MAX_SHOWS,
  MIN_AGE_MS,
  MIN_DOC_OPENS,
  RESHOW_AFTER_MS,
  asStarPromptState,
  isUpgradeLaunch,
  shouldShowStarPrompt,
  shouldShowUpgradeStarPrompt,
  withDocOpen,
  withFirstRun,
  withResolved,
  withShown,
} from '../src/main/star-prompt'

const NOW = 1_700_000_000_000

/** a state that satisfies every display condition */
const eligible = () => ({
  firstRunAt: NOW - MIN_AGE_MS,
  docOpens: MIN_DOC_OPENS,
  shownCount: 0,
})

describe('asStarPromptState', () => {
  it('returns empty state for missing or corrupt values', () => {
    expect(asStarPromptState(undefined)).toEqual({})
    expect(asStarPromptState('garbage')).toEqual({})
    expect(asStarPromptState([1, 2])).toEqual({})
    expect(asStarPromptState(null)).toEqual({})
  })

  it('drops non-numeric and negative fields', () => {
    const state = asStarPromptState({
      firstRunAt: 'soon',
      docOpens: -3,
      shownCount: 1,
      lastShownAt: Number.NaN,
      resolved: 'yes',
    })
    expect(state).toEqual({
      firstRunAt: undefined,
      docOpens: undefined,
      shownCount: 1,
      lastShownAt: undefined,
      resolved: false,
    })
  })
})

describe('shouldShowStarPrompt', () => {
  it('shows once every condition is met', () => {
    expect(shouldShowStarPrompt(eligible(), NOW)).toBe(true)
  })

  it('never shows before the install-age threshold', () => {
    expect(shouldShowStarPrompt({ ...eligible(), firstRunAt: NOW - MIN_AGE_MS + 1 }, NOW)).toBe(
      false,
    )
  })

  it('never shows without the first-run stamp', () => {
    expect(shouldShowStarPrompt({ ...eligible(), firstRunAt: undefined }, NOW)).toBe(false)
  })

  it('never shows below the doc-open threshold', () => {
    expect(shouldShowStarPrompt({ ...eligible(), docOpens: MIN_DOC_OPENS - 1 }, NOW)).toBe(false)
  })

  it('never shows once resolved', () => {
    expect(shouldShowStarPrompt({ ...eligible(), resolved: true }, NOW)).toBe(false)
  })

  it('respects the snooze window after a dismissal', () => {
    const dismissed = withShown(eligible(), NOW - RESHOW_AFTER_MS + 1)
    expect(shouldShowStarPrompt(dismissed, NOW)).toBe(false)
    const longAgo = withShown(eligible(), NOW - RESHOW_AFTER_MS)
    expect(shouldShowStarPrompt(longAgo, NOW)).toBe(true)
  })

  it('caps at the lifetime maximum', () => {
    let state = eligible()
    for (let i = 0; i < MAX_SHOWS; i++) {
      state = withShown(state, NOW - (MAX_SHOWS - i) * RESHOW_AFTER_MS)
    }
    expect(state.shownCount).toBe(MAX_SHOWS)
    expect(shouldShowStarPrompt(state, NOW)).toBe(false)
  })
})

describe('isUpgradeLaunch', () => {
  it('detects a version change', () => {
    expect(isUpgradeLaunch('1.2.0', '1.3.0', true)).toBe(true)
    expect(isUpgradeLaunch('1.3.0', '1.3.0', true)).toBe(false)
  })

  it('treats a pre-tracking install that finished onboarding as an upgrade', () => {
    expect(isUpgradeLaunch(null, '1.3.0', true)).toBe(true)
  })

  it('does not treat a fresh install as an upgrade', () => {
    expect(isUpgradeLaunch(null, '1.3.0', false)).toBe(false)
  })
})

describe('shouldShowUpgradeStarPrompt', () => {
  it('shows for a never-prompted, unresolved user regardless of value gates', () => {
    expect(shouldShowUpgradeStarPrompt({})).toBe(true)
    expect(shouldShowUpgradeStarPrompt({ firstRunAt: NOW, docOpens: 0 })).toBe(true)
  })

  it('never shows once prompted or resolved', () => {
    expect(shouldShowUpgradeStarPrompt(withShown({}, NOW))).toBe(false)
    expect(shouldShowUpgradeStarPrompt(withResolved({}))).toBe(false)
  })
})

describe('transitions', () => {
  it('withFirstRun stamps only once', () => {
    const stamped = withFirstRun({}, NOW)
    expect(stamped.firstRunAt).toBe(NOW)
    expect(withFirstRun(stamped, NOW + 1)).toBe(stamped)
  })

  it('withDocOpen keeps counting past the threshold (personalized copy)', () => {
    let state: ReturnType<typeof asStarPromptState> = {}
    for (let i = 1; i <= MIN_DOC_OPENS + 3; i++) {
      state = withDocOpen(state)
      expect(state.docOpens).toBe(i)
    }
  })

  it('withDocOpen stops once no prompt can ever show again', () => {
    const resolved = { resolved: true }
    expect(withDocOpen(resolved)).toBe(resolved)
    const shownOut = { shownCount: MAX_SHOWS }
    expect(withDocOpen(shownOut)).toBe(shownOut)
  })

  it('withResolved silences the prompt for good', () => {
    const state = withResolved(eligible())
    expect(shouldShowStarPrompt(state, NOW + 365 * 24 * 60 * 60 * 1000)).toBe(false)
  })
})
