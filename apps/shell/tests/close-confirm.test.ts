import { describe, expect, it } from 'vitest'
import {
  CONFIRM_CLOSE_ALL_KEY,
  MIN_DOCUMENTS_TO_CONFIRM_CLOSE,
  closeConfirmEnabled,
  shouldConfirmCloseAll,
} from '../src/main/close-confirm'

describe('closeConfirmEnabled', () => {
  it('asks by default when the setting was never written', () => {
    expect(closeConfirmEnabled({})).toBe(true)
  })

  it('honours an explicit opt-out and any other value means asking', () => {
    expect(closeConfirmEnabled({ [CONFIRM_CLOSE_ALL_KEY]: false })).toBe(false)
    expect(closeConfirmEnabled({ [CONFIRM_CLOSE_ALL_KEY]: true })).toBe(true)
    expect(closeConfirmEnabled({ [CONFIRM_CLOSE_ALL_KEY]: 'no' })).toBe(true)
  })
})

describe('shouldConfirmCloseAll', () => {
  it('confirms once more than one document tab is open', () => {
    expect(MIN_DOCUMENTS_TO_CONFIRM_CLOSE).toBe(2)
    expect(shouldConfirmCloseAll(2, {})).toBe(true)
    expect(shouldConfirmCloseAll(3, {})).toBe(true)
  })

  it('leaves zero or one document on the fast path', () => {
    expect(shouldConfirmCloseAll(0, {})).toBe(false)
    expect(shouldConfirmCloseAll(1, {})).toBe(false)
  })

  it('never confirms after the opt-out', () => {
    expect(shouldConfirmCloseAll(5, { [CONFIRM_CLOSE_ALL_KEY]: false })).toBe(false)
  })
})
