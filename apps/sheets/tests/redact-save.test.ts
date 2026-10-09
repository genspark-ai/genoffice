/**
 * The save side: the renderer's marks must actually reach the package part, and
 * a rename done in the same save must not orphan them.
 *
 * The gateway half of this is already covered in xlsx-redaction-roundtrip.test.ts
 * (the part is written, declared, and re-keyed). What is *not* covered anywhere
 * is the plumbing between the two: that the renderer hands the marks to the
 * save, that the request schema accepts them, and that a sheet renamed in the
 * same save carries its marks to the new name. A break in that seam is silent —
 * the save succeeds and the file claims nothing is withheld.
 */
import JSZip from 'jszip'
import { describe, expect, it, vi, afterAll, beforeAll, beforeEach } from 'vitest'

import { handleSave, type SaveContext } from '../src/renderer/save-actions'
import { workbookSaveRequestSchema } from '../src/shared/desktop-api'
import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'
import type { SheetRedactionState } from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import {
  REDACTION_PART_PATH,
  parseRedactionPart,
} from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import { saveWorkbookViaSidecar } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import { buildEditFixture, buildStructureFixture } from './fixture-builder'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const saveWorkbookEdits = vi.fn()
const writeWorkbookRecovery = vi.fn()

beforeEach(() => {
  saveWorkbookEdits.mockReset().mockResolvedValue({ canceled: true })
  writeWorkbookRecovery.mockReset().mockResolvedValue({ ok: true })
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: { saveWorkbookEdits, writeWorkbookRecovery },
  }
})

const MARKS: SheetRedactionState[] = [
  {
    sheetName: 'Data',
    marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }],
  },
]

function ctx(redactionStates: readonly SheetRedactionState[], dirty = true): SaveContext {
  const journal = createEditJournal()
  if (dirty) recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'edited' } } })
  return {
    univerRef: { current: null },
    lazyWorkbookRef: {
      current: {
        editJournal: journal,
        recalc: {
          timer: null,
          generation: 0,
          failures: 0,
          formulaCells: new Map(),
          overlay: new Map(),
        },
        file: { sessionId: '11111111-1111-4111-8111-111111111111' },
      },
    } as never,
    setMessage: () => {},
    openLazyWorkbook: () => {},
    stashViewRestore: () => {},
    redactionStates: () => redactionStates,
  }
}

describe('the save carries the withheld cells', () => {
  it('puts the marks in the request it sends', async () => {
    await handleSave(ctx(MARKS), 'save')
    expect(saveWorkbookEdits).toHaveBeenCalledTimes(1)
    const request = saveWorkbookEdits.mock.calls[0]![0] as { redactionStates: unknown }
    expect(request.redactionStates).toEqual([
      {
        sheetName: 'Data',
        marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }],
      },
    ])
  })

  it('carries them in the crash-recovery copy too', async () => {
    // A recovery copy that dropped the marks would resurrect the file with
    // nothing withheld after a crash.
    await handleSave(ctx(MARKS), 'recovery')
    const payload = writeWorkbookRecovery.mock.calls[0]![0] as { redactionStates: unknown }
    expect(payload.redactionStates).toEqual([
      {
        sheetName: 'Data',
        marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }],
      },
    ])
  })

  it('saves a workbook whose only pending change is a newly withheld cell', async () => {
    // Nothing in the journal is dirty here. Refusing this save would leave the
    // file claiming nothing is hidden while the session says otherwise.
    await handleSave(ctx(MARKS, false), 'save')
    expect(saveWorkbookEdits).toHaveBeenCalledTimes(1)
  })

  it('sends nothing when the workbook withholds nothing', async () => {
    await handleSave(ctx([]), 'save')
    const request = saveWorkbookEdits.mock.calls[0]![0] as { redactionStates: unknown }
    expect(request.redactionStates).toEqual([])
  })

  it('is accepted by the schema the main process parses with', () => {
    // The renderer and main share this schema; a field the preload or the
    // renderer does not send would fail the whole save.
    const parsed = workbookSaveRequestSchema.parse({
      sessionId: '11111111-1111-4111-8111-111111111111',
      mode: 'save',
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
      redactionStates: [
        {
          sheetName: 'Data',
          marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'x' }],
        },
      ],
    })
    expect(parsed.redactionStates).toHaveLength(1)
  })
})

describe('the marks reach the package through a real save', () => {
  let directory: string
  let client: XlsxSidecarClient

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'redaction-save-'))
    client = new XlsxSidecarClient(sidecarBinaryPath())
  })

  afterAll(async () => {
    client.stop()
    await rm(directory, { recursive: true, force: true })
  })

  async function savedWith(
    redactionStates: readonly SheetRedactionState[],
    plan?: SheetEditPlan,
  ): Promise<string> {
    const sourcePath = join(directory, 'source.xlsx')
    const targetPath = join(directory, `saved-${Math.random().toString(36).slice(2)}.xlsx`)
    await writeFile(sourcePath, await buildEditFixture())
    await saveWorkbookViaSidecar({
      client,
      sourcePath,
      targetPath,
      edits: [],
      sheetPlan: plan,
      redactionStates,
    })
    return targetPath
  }

  it('writes the part the renderer sent, and reads it back', async () => {
    const saved = await savedWith(MARKS)
    const zip = await JSZip.loadAsync(await readFile(saved))
    const raw = await zip.file(REDACTION_PART_PATH)?.async('string')
    expect(raw, 'the part is missing from the saved package').toBeDefined()
    expect(parseRedactionPart(raw!)).toEqual(MARKS)
  })

  it('re-keys the marks for a sheet renamed in the same save', async () => {
    // The part is keyed by name, so a rename that did not travel with the
    // marks would leave the part pointing at a sheet that no longer exists —
    // and the withheld values would go back to the model with nothing on
    // screen saying so.
    const plan: SheetEditPlan = {
      renames: [{ sheetName: 'Data', newName: 'Q3' }],
      removals: [],
      additions: [],
      order: ['Q3'],
    }
    const saved = await savedWith(MARKS, plan)
    const zip = await JSZip.loadAsync(await readFile(saved))
    const raw = await zip.file(REDACTION_PART_PATH)?.async('string')
    expect(parseRedactionPart(raw!)).toEqual([{ ...MARKS[0], sheetName: 'Q3' }])
  })

  it('drops the marks of a sheet removed in the same save', async () => {
    // The structure fixture has no defined names pointing at its sheets, so a
    // removal is a legal save. "Other" keeps its own mark, so the part is
    // rewritten — and must not carry the removed sheet's marks, whose cells no
    // longer exist to withhold.
    const plan: SheetEditPlan = {
      renames: [],
      removals: ['Data'],
      additions: [],
      order: ['Other'],
    }
    const states: SheetRedactionState[] = [
      ...MARKS,
      {
        sheetName: 'Other',
        marks: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0, label: 'kept' }],
      },
    ]
    const sourcePath = join(directory, 'removal-source.xlsx')
    const targetPath = join(directory, 'removal-saved.xlsx')
    await writeFile(sourcePath, await buildStructureFixture())
    await saveWorkbookViaSidecar({
      client,
      sourcePath,
      targetPath,
      edits: [],
      sheetPlan: plan,
      redactionStates: states,
    })
    const zip = await JSZip.loadAsync(await readFile(targetPath))
    const raw = await zip.file(REDACTION_PART_PATH)?.async('string')
    expect(raw, 'the surviving sheet still withholds a cell').toBeDefined()
    expect(parseRedactionPart(raw!)).toEqual([states[1]])
  })
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}
