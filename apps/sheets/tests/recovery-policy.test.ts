import { describe, expect, it } from 'vitest'

import {
  allowsAutomaticWorkbookRecovery,
  MAX_AUTOMATIC_RECOVERY_WORKSHEET_XML_BYTES,
} from '../src/main/recovery-policy'

describe('automatic workbook recovery policy', () => {
  it('allows normal workbooks and older sidecar metadata', () => {
    expect(allowsAutomaticWorkbookRecovery([{ sourceXmlBytes: 8 * 1024 * 1024 }])).toBe(true)
    expect(allowsAutomaticWorkbookRecovery([{}])).toBe(true)
  })

  it('disables background rewriting when any worksheet XML entry is oversized', () => {
    expect(
      allowsAutomaticWorkbookRecovery([
        { sourceXmlBytes: 4 * 1024 * 1024 },
        { sourceXmlBytes: MAX_AUTOMATIC_RECOVERY_WORKSHEET_XML_BYTES + 1 },
      ]),
    ).toBe(false)
  })

  it('accepts an entry at the memory-safe limit', () => {
    expect(
      allowsAutomaticWorkbookRecovery([
        { sourceXmlBytes: MAX_AUTOMATIC_RECOVERY_WORKSHEET_XML_BYTES },
      ]),
    ).toBe(true)
  })
})
