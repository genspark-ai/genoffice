/**
 * The main process parses every save through workbookSaveRequestSchema, so its
 * "at least one edit" refine must exempt explicit Save As — a clean workbook
 * saved to a new path is a valid request with nothing to apply.
 */
import { describe, expect, it } from 'vitest'
import { workbookSaveRequestSchema } from '../src/shared/desktop-api'

function emptyRequest(mode: 'save' | 'save-as') {
  return {
    sessionId: '5d4f6f7a-1c2b-4e3d-9a8f-0b1c2d3e4f5a',
    mode,
    edits: [],
    structuralOps: [],
    chartEdits: [],
    visualEdits: [],
    visualAdditions: [],
    tableAdditions: [],
    pivotAdditions: [],
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    formulaValues: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    sheetProtections: [],
    sparklineAdditions: [],
    definedNamesState: null,
  }
}

describe('workbookSaveRequestSchema', () => {
  it('accepts an empty save-as request', () => {
    expect(() => workbookSaveRequestSchema.parse(emptyRequest('save-as'))).not.toThrow()
  })

  it('still rejects an empty ordinary save', () => {
    expect(() => workbookSaveRequestSchema.parse(emptyRequest('save'))).toThrow(/at least one edit/)
  })
})
