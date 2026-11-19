/** @vitest-environment jsdom */
import { describe, expect, it } from 'vitest'
import { installPopoverDismiss } from '../src/popover-dismiss'

describe('installPopoverDismiss', () => {
  it('keeps the open class while another popover is open when one teardown runs twice', () => {
    const first = installPopoverDismiss(() => {})
    const second = installPopoverDismiss(() => {})
    expect(document.documentElement.classList.contains('genoffice-popover-open')).toBe(true)
    first()
    first()
    expect(document.documentElement.classList.contains('genoffice-popover-open')).toBe(true)
    second()
    expect(document.documentElement.classList.contains('genoffice-popover-open')).toBe(false)
  })
})
