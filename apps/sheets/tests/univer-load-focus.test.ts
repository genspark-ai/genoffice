// @vitest-environment jsdom
import { IUniverInstanceService, LogLevel, Univer } from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/facade'
import { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it } from 'vitest'

import { loadSnapshotIntoUniver, loadWorkbookSkeleton } from '../src/renderer/univer-sync'
import type { UniverRuntime } from '../src/renderer/univer-state'
import type { WorkbookFile } from '../src/shared/desktop-api'

/**
 * A workbook loaded into an already-mounted view (the prewarmed spare) must
 * stay the FOCUSED unit. `disposeUnit` of the replaced workbook resets the
 * focused unit to null (`_tryResetFocusOnRemoval`) and `createWorkbook` does
 * not focus what it created — a load into the spare then sits with NO focused
 * unit, while the stale FOCUSING_* context bits keep the shortcut-gated keys
 * (Enter, arrows) alive and hide the damage: the cell editor's character-key
 * routing and ILayoutService.focus() (early return on a null focused unit)
 * are dead, so typing into the adopted view never opens the editor
 * (e2e open-focus-typing, genoffice#1150's CI-only signature).
 */

const FILE: WorkbookFile = {
  sessionId: 'session',
  path: '/tmp/book.xlsx',
  name: 'Book1.xlsx',
  sha256: 'abc123',
  fileBytes: new Uint8Array(),
  visuals: [],
  sheets: [
    {
      id: 'sheet-1',
      name: 'Sheet1',
      rowCount: 1,
      columnCount: 1,
      hidden: false,
      showGridLines: true,
      tabColor: null,
      defaultRowHeight: null,
      defaultColumnWidth: null,
      freeze: null,
      columnWidths: [],
      cells: {},
      tables: [],
      protections: [],
      merges: [],
      notes: [],
      pivotTables: [],
      showFormulas: false,
    },
  ],
} as unknown as WorkbookFile

function bootRealUniver(): UniverRuntime {
  const univer = new Univer({ logLevel: LogLevel.SILENT })
  univer.registerPlugin(UniverSheetsPlugin)
  const univerAPI = FUniver.newAPI(univer)
  return { univer, univerAPI } as unknown as UniverRuntime
}

describe('a workbook load leaves its unit focused', () => {
  it('the boot placeholder is the focused unit, and opening a file re-focuses the file unit', () => {
    const runtime = bootRealUniver()
    // the prewarmed spare boots with the placeholder demo workbook
    loadSnapshotIntoUniver(
      runtime,
      { revision: 0, sheets: [{ id: 'sheet-1', name: 'Sheet1', cells: {} }] },
      'new-workbook',
      'Untitled',
    )
    const instances = runtime.univer.__getInjector().get(IUniverInstanceService)
    expect(instances.getFocusedUnit()?.getUnitId()).toBe('new-workbook')

    // opening a workbook disposes the placeholder and creates the file unit
    loadWorkbookSkeleton(runtime, FILE)
    expect(instances.getFocusedUnit()?.getUnitId()).toBe('file-abc123')
  })

  it('a rebuild (same unitId) re-focuses the rebuilt workbook', () => {
    const runtime = bootRealUniver()
    loadSnapshotIntoUniver(
      runtime,
      { revision: 0, sheets: [{ id: 'sheet-1', name: 'Sheet1', cells: {} }] },
      'new-workbook',
      'Untitled',
    )
    loadSnapshotIntoUniver(
      runtime,
      { revision: 0, sheets: [{ id: 'sheet-1', name: 'Sheet2', cells: {} }] },
      'new-workbook',
      'Untitled',
    )
    const instances = runtime.univer.__getInjector().get(IUniverInstanceService)
    expect(instances.getFocusedUnit()?.getUnitId()).toBe('new-workbook')
  })
})
