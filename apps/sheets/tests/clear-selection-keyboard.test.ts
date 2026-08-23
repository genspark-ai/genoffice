import { describe, expect, it } from 'vitest'

import {
  isClearSelectionHotkey,
  shouldInterceptClearSelection,
  SKIP_HOST_SELECTOR,
  type ClearSelectionKeyEvent,
} from '../src/renderer/clear-selection-keyboard'

function keyEvent(
  key: string,
  hosts: readonly string[],
  extra: Partial<ClearSelectionKeyEvent> = {},
): ClearSelectionKeyEvent {
  return {
    key,
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    shiftKey: false,
    defaultPrevented: false,
    isComposing: false,
    target: {
      closest(selector: string) {
        return hosts.some((host) => selector === host || selector.includes(host)) ? {} : null
      },
    },
    ...extra,
  }
}

describe('isClearSelectionHotkey', () => {
  it('matches unmodified Delete and Backspace', () => {
    expect(
      isClearSelectionHotkey({
        key: 'Delete',
        metaKey: false,
        ctrlKey: false,
        altKey: false,
        shiftKey: false,
      }),
    ).toBe(true)
    expect(
      isClearSelectionHotkey({
        key: 'Backspace',
        metaKey: false,
        ctrlKey: false,
        altKey: false,
        shiftKey: false,
      }),
    ).toBe(true)
  })

  it('ignores modifiers and other keys', () => {
    expect(
      isClearSelectionHotkey({
        key: 'Backspace',
        metaKey: false,
        ctrlKey: false,
        altKey: false,
        shiftKey: true,
      }),
    ).toBe(false)
    expect(
      isClearSelectionHotkey({
        key: 'Delete',
        metaKey: false,
        ctrlKey: true,
        altKey: false,
        shiftKey: false,
      }),
    ).toBe(false)
    expect(
      isClearSelectionHotkey({
        key: 'Enter',
        metaKey: false,
        ctrlKey: false,
        altKey: false,
        shiftKey: false,
      }),
    ).toBe(false)
  })
})

describe('shouldInterceptClearSelection', () => {
  it('intercepts grid-hosted contenteditable (Univer hidden editor)', () => {
    const hosts = ['[contenteditable="true"]', '#univer-container']
    expect(shouldInterceptClearSelection(keyEvent('Backspace', hosts), false)).toBe(true)
    expect(shouldInterceptClearSelection(keyEvent('Delete', hosts), false)).toBe(true)
  })

  it('does not intercept while a cell is being edited', () => {
    expect(shouldInterceptClearSelection(keyEvent('Backspace', []), true)).toBe(false)
  })

  it('does not intercept native inputs, including Univer panel fields', () => {
    expect(
      shouldInterceptClearSelection(keyEvent('Backspace', ['input, textarea, select']), false),
    ).toBe(false)
    expect(
      shouldInterceptClearSelection(
        keyEvent('Backspace', ['[data-u-comp="input"]', '#univer-container']),
        false,
      ),
    ).toBe(false)
    expect(
      shouldInterceptClearSelection(
        keyEvent('Backspace', ['[data-u-comp="panel"]', '#univer-container']),
        false,
      ),
    ).toBe(false)
  })

  it('does not intercept app chrome contenteditable outside the grid', () => {
    expect(
      shouldInterceptClearSelection(keyEvent('Backspace', ['[contenteditable="true"]']), false),
    ).toBe(false)
  })

  it('does not intercept the formula bar or a focused visual', () => {
    expect(
      shouldInterceptClearSelection(
        keyEvent('Backspace', ['[data-u-comp="formula-bar"]', '#univer-container']),
        false,
      ),
    ).toBe(false)
    expect(shouldInterceptClearSelection(keyEvent('Delete', ['.shape-editable']), false)).toBe(
      false,
    )
  })

  it('does not intercept dialog fields', () => {
    expect(shouldInterceptClearSelection(keyEvent('Backspace', ['[role="dialog"]']), false)).toBe(
      false,
    )
  })

  it('lists Univer chrome hosts in the skip selector', () => {
    expect(SKIP_HOST_SELECTOR).toContain('[data-u-comp="input"]')
    expect(SKIP_HOST_SELECTOR).toContain('[data-u-comp="panel"]')
    expect(SKIP_HOST_SELECTOR).toContain('[data-u-comp="formula-bar"]')
  })
})
