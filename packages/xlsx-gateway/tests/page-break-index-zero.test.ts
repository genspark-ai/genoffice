import { describe, expect, it } from 'vitest'

import { applyPageSetupState } from '../src/gateway/xlsx-page-setup'

/// brk@id is a 0-based row/column index, so 0 is a real break (a page break
/// above the first row / left of the first column) and must be written rather
/// than dropped.
const BARE = '<worksheet><sheetData/></worksheet>'
const WITH_BREAKS =
  '<worksheet><sheetData/>' +
  '<rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks>' +
  '<colBreaks count="1" manualBreakCount="1"><brk id="2" max="1048575" man="1"/></colBreaks>' +
  '</worksheet>'

describe('applyPageSetupState manual break at index 0', () => {
  it('writes a row break at index 0', () => {
    const xml = applyPageSetupState(BARE, { sheetName: 'S', rowBreaks: [0] })
    expect(xml).toBe(
      '<worksheet><sheetData/>' +
        '<rowBreaks count="1" manualBreakCount="1"><brk id="0" max="16383" man="1"/></rowBreaks>' +
        '</worksheet>',
    )
  })

  it('writes a column break at index 0', () => {
    const xml = applyPageSetupState(BARE, { sheetName: 'S', colBreaks: [0] })
    expect(xml).toBe(
      '<worksheet><sheetData/>' +
        '<colBreaks count="1" manualBreakCount="1"><brk id="0" max="1048575" man="1"/></colBreaks>' +
        '</worksheet>',
    )
  })

  it('sorts index 0 first and keeps it in the count', () => {
    const xml = applyPageSetupState(BARE, { sheetName: 'S', rowBreaks: [20, 0, 20] })
    expect(xml).toBe(
      '<worksheet><sheetData/>' +
        '<rowBreaks count="2" manualBreakCount="2">' +
        '<brk id="0" max="16383" man="1"/><brk id="20" max="16383" man="1"/>' +
        '</rowBreaks>' +
        '</worksheet>',
    )
  })

  it('replaces an existing rowBreaks element instead of deleting it', () => {
    const xml = applyPageSetupState(WITH_BREAKS, { sheetName: 'S', rowBreaks: [0] })
    expect(xml).toContain('<brk id="0" max="16383" man="1"/>')
    expect(xml).not.toContain('id="5"')
    expect(xml).toContain('<colBreaks count="1" manualBreakCount="1">')
  })

  it('replaces an existing colBreaks element instead of deleting it', () => {
    const xml = applyPageSetupState(WITH_BREAKS, { sheetName: 'S', colBreaks: [0] })
    expect(xml).toContain('<brk id="0" max="1048575" man="1"/>')
    expect(xml).not.toContain('id="2"')
    expect(xml).toContain('<rowBreaks count="1" manualBreakCount="1">')
  })

  it('still clears both elements for an empty set', () => {
    const xml = applyPageSetupState(WITH_BREAKS, { sheetName: 'S', rowBreaks: [], colBreaks: [] })
    expect(xml).toBe(BARE)
  })

  it('leaves existing breaks alone when the key is absent', () => {
    const xml = applyPageSetupState(WITH_BREAKS, { sheetName: 'S', zoomScale: 100 })
    expect(xml).toBe(WITH_BREAKS)
  })
})
