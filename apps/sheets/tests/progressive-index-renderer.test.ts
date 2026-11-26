import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { setTimeout as delay } from 'node:timers/promises'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { afterEach, expect, it, vi } from 'vitest'

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import {
  workbookFileSchema,
  workbookFormulaCellsResultSchema,
  workbookRangeResultSchema,
  workbookRowOutlineResultSchema,
  type WorkbookFormulaCellsRequest,
  type WorkbookRangeRequest,
} from '../src/shared/desktop-api'
import { createEditJournal } from '../src/renderer/edit-journal'
import { progressiveFullLoadMode } from '../src/renderer/load-budget'
import { loadRowOutlines } from '../src/renderer/outline-actions'
import {
  activateFormulaClosure,
  ensureSheetFileState,
  formulaClosureStatus,
  recalcOverBudgetAtOpen,
} from '../src/renderer/univer-sync'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

interface Tab {
  name: string
  hidden?: boolean
  grouped?: boolean
  formula?: string
  dimension?: string
  longInlineStrings?: boolean
}

const active = { name: 'Dashboard', formula: 'A1*2', dimension: 'A1:R1000' }
const manyTabs = Array.from({ length: 6 }, (_, i) => ({ name: 'Tab' + (i + 1) }))
const viewport = { startRow: 0, endRow: 20, startColumn: 0, endColumn: 1 }
const openedResources: (() => Promise<void>)[] = []

afterEach(async () => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
  for (const cleanup of openedResources.splice(0).reverse()) await cleanup()
})

async function workbookBytes(tabs: Tab[]): Promise<Buffer> {
  const zip = new JSZip()
  const names = tabs.map(({ name, hidden }, index) => {
    const id = index + 1
    return (
      '<sheet name="' +
      name +
      '" sheetId="' +
      id +
      '" r:id="rId' +
      id +
      '"' +
      (hidden ? ' state="hidden"' : '') +
      '/>'
    )
  })
  zip.file(
    'xl/workbook.xml',
    '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets>' +
      names.join('') +
      '</sheets></workbook>',
  )
  const rels = tabs.map(
    (_, i) =>
      '<Relationship Id="rId' + (i + 1) + '" Target="worksheets/sheet' + (i + 1) + '.xml"/>',
  )
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships>' + rels.join('') + '</Relationships>')
  tabs.forEach((tab, i) => {
    const formula = tab.formula ? '<f>' + tab.formula + '</f><v>4</v>' : '<v>10</v>'
    const group = tab.grouped
      ? '<row r="600" outlineLevel="1" hidden="1"><c r="A600"><v>8</v></c></row>' +
        '<row r="601" collapsed="1"><c r="A601"><v>9</v></c></row>'
      : ''
    const text = tab.longInlineStrings ? 'x'.repeat(32_767) : ''
    const longRows = text
      ? Array.from({ length: 1_024 }, (_, offset) => {
          const row = offset + 2
          return `<row r="${row}"><c r="A${row}" t="inlineStr"><is><t>${text}</t></is></c></row>`
        }).join('')
      : ''
    zip.file(
      'xl/worksheets/sheet' + (i + 1) + '.xml',
      '<worksheet><dimension ref="' +
        (tab.dimension ?? (group ? 'A1:B601' : 'A1:B21')) +
        '"/>' +
        (group ? '<sheetFormatPr outlineLevelRow="1"/>' : '') +
        '<sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1">' +
        formula +
        '</c></row>' +
        group +
        longRows +
        (i === 0 && !text ? ' '.repeat(32 * 1024 * 1024) : '') +
        '</sheetData></worksheet>',
    )
  })
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function harness(tabs: Tab[]) {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-progressive-xlsx-'))
  const path = join(dir, 'workbook.xlsx')
  const executable = fileURLToPath(
    new URL(
      '../native/xlsx-engine/target/release/xlsx-sidecar' +
        (process.platform === 'win32' ? '.exe' : ''),
      import.meta.url,
    ),
  )
  const client = new XlsxSidecarClient(executable)
  openedResources.push(async () => {
    client.stop()
    await rm(dir, { recursive: true, force: true })
  })
  await writeFile(path, await workbookBytes(tabs))
  const raw = await client.open(path, 'en', undefined, true)
  const file = workbookFileSchema.parse({
    ...(raw as object),
    sha256: '0'.repeat(64),
    readOnly: false,
  })
  openedResources.push(() => client.close(file.sessionId))
  expect(file.sheets[0]?.progressive).toBe(true)

  const readWorkbookRange = vi.fn(async (input: WorkbookRangeRequest) =>
    workbookRangeResultSchema.parse(await client.readRange(input)),
  )
  const readWorkbookFormulas = vi.fn(async (input: WorkbookFormulaCellsRequest) =>
    workbookFormulaCellsResultSchema.parse(await client.readFormulaCells(input)),
  )
  const readWorkbookRowOutline = vi.fn(async (input: WorkbookFormulaCellsRequest) =>
    workbookRowOutlineResultSchema.parse(await client.readRowOutline(input)),
  )
  vi.stubGlobal('window', {
    desktopApi: {
      readWorkbookFormulas,
      readWorkbookRange,
      readWorkbookRowOutline,
    },
  })
  const state = {
    file,
    formulaMode: false,
    editJournal: createEditJournal(),
    closure: { status: 'idle', pinned: new Map() },
    formulaText: new Map(),
    sharedFormulaGroups: new Map(),
    cachedFormulaValues: new Map(),
    sheetPageBreaks: new Map(),
    sheetProtections: new Map(),
    sheetFilePageSetups: new Map(),
    sheetProtectedRanges: new Map(),
    rowColStyleKeys: new Map(),
    hiddenFileRows: new Map(),
    hiddenRowsCoveredThrough: new Map(),
    outline: new Map(),
    recalc: {
      timer: null,
      generation: 0,
      failures: 0,
      engineOverBudget: recalcOverBudgetAtOpen(
        file.fileBytes,
        file.sheets.reduce((sum, sheet) => sum + sheet.rowCount * sheet.columnCount, 0),
      ),
      formulaCells: new Map(),
      overlay: new Map(),
      incompleteRuns: new Set(),
      follow: new Map(),
      running: false,
      lastRunAt: 0,
    },
  } as unknown as LazyWorkbookState
  openedResources.push(async () => {
    if (state.recalc.timer) clearTimeout(state.recalc.timer)
  })
  const ref = { current: state }
  const setMessage = vi.fn()
  let activeSheetId = 'sheet-1'
  const workbook = {
    getActiveSheet: () => ({ getSheetId: () => activeSheetId }),
    getSheetBySheetId: (sheetId: string) => ({
      getSheetId: () => sheetId,
      getRange: () => ({ setValues: () => undefined }),
    }),
  }
  const runtime = {
    univerAPI: { getActiveWorkbook: () => workbook },
  } as unknown as UniverRuntime
  const visit = async (sheetId: string) => {
    activeSheetId = sheetId
    const deadline = performance.now() + 5000
    do {
      const read = await readWorkbookRange({
        sessionId: file.sessionId,
        sheetId,
        range: viewport,
        foreground: true,
        mode: 'viewport',
      })
      if (read.indexingComplete) return read
      await delay(15)
    } while (performance.now() < deadline)
    throw new Error('Foreground indexing did not finish: ' + sheetId)
  }
  const peekFormulas = async (sheetId: string) =>
    workbookFormulaCellsResultSchema.parse(
      await client.readFormulaCells({
        sessionId: file.sessionId,
        sheetId,
        background: true,
      }),
    )
  const peekOutline = async (sheetId: string) =>
    workbookRowOutlineResultSchema.parse(
      await client.readRowOutline({
        sessionId: file.sessionId,
        sheetId,
        background: true,
      }),
    )
  await visit('sheet-1')
  return {
    state,
    ref,
    visit,
    peekFormulas,
    peekOutline,
    readWorkbookFormulas,
    readWorkbookRowOutline,
    status: (sheetId: string) => formulaClosureStatus(state, sheetId),
    formulaReads: () => readWorkbookFormulas.mock.calls.map(([req]) => req.sheetId),
    activate: () => activateFormulaClosure(runtime, ref, setMessage),
  }
}

it('activates local formulas without probing or starting an unrelated hidden sheet', async () => {
  const h = await harness([active, { name: 'Hidden', hidden: true }])
  await h.activate()

  expect(h.status('sheet-1')).toBe('active')
  expect(h.state.closure.pinned.get('sheet-1')?.get('0:1')).toEqual({ f: '=A1*2', v: 4 })
  expect(h.formulaReads()).toEqual(['sheet-1'])
  expect((await h.peekFormulas('sheet-2')).indexingComplete).toBe(false)
  expect(h.state.recalc.timer).toBeNull()
})

it('uses EOF counts for a progressive sheet with long inline strings and an oversized declared extent', async () => {
  const h = await harness([{ ...active, dimension: 'A1:R120000', longInlineStrings: true }])
  const countsChanged = vi.fn()
  h.state.onProgressiveCountsChanged = countsChanged
  h.state.recalc.formulaIndexRetired = true

  expect(h.state.recalc.engineOverBudget).toBe(true)
  expect(progressiveFullLoadMode(h.state.file.sheets)).toBeUndefined()

  expect(await ensureSheetFileState(h.state, 'sheet-1', h.ref)).toBe(true)
  expect(h.state.file.sheets[0]).toMatchObject({
    storedCellCount: 1_026,
    valueRowCount: 1_025,
    valueColumnCount: 2,
  })
  expect(progressiveFullLoadMode(h.state.file.sheets)).toBe('formulas')
  expect(h.state.recalc.engineOverBudget).toBe(false)
  expect(h.state.recalc.formulaIndexRetired).toBe(true)
  expect(countsChanged).toHaveBeenCalledExactlyOnceWith(true)
})

it('indexes transitive formula precedents and leaves another hidden sheet parked', async () => {
  const h = await harness([
    { ...active, formula: "'Hidden link'!B1" },
    { name: 'Hidden link', hidden: true, formula: 'Source!B1' },
    { name: 'Source', hidden: true, formula: 'A1*2' },
    { name: 'Unused', hidden: true },
  ])
  await h.activate()

  expect([h.status('sheet-1'), h.status('sheet-2'), h.status('sheet-3')]).toEqual([
    'active',
    'active',
    'active',
  ])
  expect(new Set(h.formulaReads())).toEqual(new Set(['sheet-1', 'sheet-2', 'sheet-3']))
  expect(h.state.closure.pinned.get('sheet-3')?.get('0:1')).toEqual({ f: '=A1*2', v: 4 })
  expect(h.state.closure.pinned.get('sheet-3')?.get('0:0')).toEqual({ v: 2 })
  expect((await h.peekFormulas('sheet-4')).indexingComplete).toBe(false)
})

it('activates a previously unstarted visible sheet on first visit and preserves the first closure', async () => {
  const h = await harness([
    { ...active, dimension: 'A1:R120000' },
    ...manyTabs,
    { name: 'Late', formula: 'A1*2' },
  ])
  expect(h.state.recalc.engineOverBudget).toBe(true)
  expect((await h.peekFormulas('sheet-8')).indexingComplete).toBe(false)
  await h.activate()
  expect(h.formulaReads()).toEqual(['sheet-1'])
  expect(h.status('sheet-1')).toBe('active')
  expect(h.status('sheet-8')).toBe('idle')
  const firstPins = h.state.closure.pinned.get('sheet-1')

  await h.visit('sheet-8')
  await h.activate()
  expect(h.status('sheet-8')).toBe('active')
  expect(h.state.closure.pinned.get('sheet-1')).toBe(firstPins)
  expect(h.state.closure.pinned.get('sheet-8')?.get('0:1')).toEqual({ f: '=A1*2', v: 4 })
  await h.visit('sheet-1')
  await h.activate()
  expect(h.formulaReads()).toEqual(['sheet-1', 'sheet-8'])
})

it('keeps an unresolved formula cached without disabling closure on other sheets', async () => {
  const h = await harness([
    { ...active, formula: 'Missing!A1' },
    { name: 'Other', formula: 'A1*2' },
  ])
  await h.activate()
  expect(h.status('sheet-1')).toBe('unavailable')
  expect(h.state.closure.pinned.size).toBe(0)
  expect(h.state.recalc.timer).not.toBeNull()
  clearTimeout(h.state.recalc.timer!)

  await h.visit('sheet-2')
  await h.activate()
  expect(h.status('sheet-2')).toBe('active')
  expect(h.formulaReads()).toEqual(['sheet-1', 'sheet-2'])
})

it('does not demand more dependencies after the active tab changes mid-read', async () => {
  const h = await harness([
    { ...active, formula: 'Source!A1' },
    { name: 'Source', hidden: true },
    { name: 'Other', formula: 'A1*2' },
  ])
  const read = h.readWorkbookFormulas.getMockImplementation()!
  let release!: () => void
  const blocked = new Promise<void>((resolve) => {
    release = resolve
  })
  h.readWorkbookFormulas.mockImplementationOnce(async (request) => {
    await blocked
    return read(request)
  })

  const first = h.activate()
  expect(h.status('sheet-1')).toBe('pending')
  await h.visit('sheet-3')
  await h.activate()
  release()
  await first
  expect(h.status('sheet-1')).toBe('idle')
  expect(h.status('sheet-3')).toBe('active')
  expect(h.formulaReads()).toEqual(['sheet-1', 'sheet-3'])

  await h.visit('sheet-1')
  await h.activate()
  expect(h.status('sheet-1')).toBe('active')
  expect(h.formulaReads()).toContain('sheet-2')
})

it('loads and deduplicates distant row outlines on first visit after the old deadline', async () => {
  const h = await harness([
    active,
    { name: 'HiddenGroup', hidden: true, grouped: true },
    ...manyTabs,
    { name: 'LateGroup', grouped: true },
  ])
  const late = 'sheet-9'
  await loadRowOutlines(h.state, h.ref)
  expect(h.readWorkbookRowOutline).not.toHaveBeenCalled()
  expect(await h.peekOutline(late)).toEqual({ indexingComplete: false, rows: [] })
  vi.spyOn(Date, 'now').mockReturnValue(Date.now() + 180_001)
  const loaded = await h.visit(late)
  expect(loaded.rows.some((row) => row.row === 599)).toBe(false)

  await Promise.all([loadRowOutlines(h.state, h.ref, late), loadRowOutlines(h.state, h.ref, late)])
  await loadRowOutlines(h.state, h.ref, late)
  expect(h.readWorkbookRowOutline).toHaveBeenCalledTimes(1)
  expect(h.readWorkbookRowOutline.mock.calls[0]?.[0].sheetId).toBe(late)
  expect(h.state.outline.get(late)?.rows.get(599)).toMatchObject({ level: 1, fileHidden: true })
  expect(h.state.outline.get(late)?.rows.get(600)).toMatchObject({ collapsed: true })
  expect((await h.peekOutline('sheet-2')).indexingComplete).toBe(false)
})
