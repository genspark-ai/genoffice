import { describe, expect, it } from 'vitest'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { textboxBandBottom } from '../src/renderer/editor/extensions'

describe('wrapTopAndBottom band vs live box height', () => {
  const box: TextboxDisplay = { paras: [], bandTopPx: 20, bandBottomPx: 70, heightPx: 50 }

  it('matches the parse-time bottom at the parse-time height', () => {
    expect(textboxBandBottom(box)).toBe(70)
  })

  it('follows a changed height (autogrow / committed heightPx)', () => {
    expect(textboxBandBottom(box, 120)).toBe(140)
    expect(textboxBandBottom({ ...box, heightPx: 120 })).toBe(140)
  })

  it('keeps the parse-time bottom for boxes without a known height', () => {
    const autoFit: TextboxDisplay = { paras: [], bandBottomPx: 90 }
    expect(textboxBandBottom(autoFit)).toBe(90)
    expect(textboxBandBottom(autoFit, undefined)).toBe(90)
  })
})
