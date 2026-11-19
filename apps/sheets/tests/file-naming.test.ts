/**
 * The sheets editor's file naming: the one path a model name takes.
 *
 * ## What is covered here
 *
 * `modelSaveAsName` is the app-side half of the shared naming client, and the
 * ribbon's "Name with AI…" button is its only caller. These tests go through
 * it rather than through `suggestFileName` directly: a test of the shared
 * client would still pass if this app stopped using the projection — which is
 * the failure that leaks a value, so the app's own call site is what needs
 * covering.
 *
 * The context is built with the app's own `WorkbookReadContext` shape and the
 * real `InMemoryWorkbookAdapter`, because the module reads the workbook through
 * the same ref the AI panel does. A hand-rolled stand-in would let the test
 * pass while the app kept handing the model raw cell values.
 *
 * ## What is deliberately not here
 *
 * There is no test for a name arriving on its own: the automatic path is gone,
 * and the tests below say so — the naming module has no automatic entry point,
 * the save path never asks the model, and the ribbon button is the only way in.
 */

import { readFileSync } from 'node:fs'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import type { WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { SheetRedactionState } from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import { buildRedactionIndex, NO_REDACTIONS } from '../src/renderer/ai/redact'
import type { WorkbookReadContext } from '../src/renderer/ai/workbook-readers'
import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'
import { handleSave as handleSaveImpl, type SaveContext } from '../src/renderer/save-actions'
import * as fileNaming from '../src/renderer/file-naming'
import { modelSaveAsName } from '../src/renderer/file-naming'

const LABEL = '\u5ba2\u6237\u7535\u8bdd'
const SECRET = '13800138000'

const SHEETS = [
  { id: 'sh1', name: 'Customers' },
  { id: 'sh2', name: 'Orders' },
]

/** B2 holds the phone number the reader hid; A1 and C1 are ordinary cells. */
const SNAPSHOT: WorkbookSnapshot = {
  revision: 1,
  sheets: [
    {
      id: 'sh1',
      name: 'Customers',
      cells: {
        A1: { value: '\u5b63\u5ea6\u5ba2\u6237\u56de\u8bbf' },
        B2: { value: SECRET },
        C1: { value: 42 },
      },
    },
    { id: 'sh2', name: 'Orders', cells: { A1: { value: 'widget' } } },
  ],
}

const MARKS: SheetRedactionState[] = [
  {
    sheetName: 'Customers',
    marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: LABEL }],
  },
]

/** A read context over an opened file: cells live in the grid, the adapter is the blank boot workbook. */
function fileContext(): WorkbookReadContext {
  const grid: Record<string, string> = { A1: 'Region', B1: 'Revenue', A2: 'North', B2: '1200' }
  // C1 is a rich-text title: no plain v, only the document stream
  const worksheet = {
    getSheetId: () => 'sheet-1',
    getSheetName: () => 'Sales',
    getRange: (address: string) => ({
      getFormula: () => '',
      getValue: () => grid[address] ?? null,
      getCellDatas: () => [
        [
          address === 'C1'
            ? { p: { body: { dataStream: 'FY26 Plan\r\n' } } }
            : grid[address] === undefined
              ? undefined
              : { v: grid[address] },
        ],
      ],
    }),
  }
  const workbook = { getSheets: () => [worksheet] }
  return {
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
    lazyWorkbookRef: {
      current: {
        editJournal: createEditJournal(),
        file: { sheets: [{ id: 'sheet-1', rowCount: 2, columnCount: 3 }] },
      },
    },
    adapterRef: { current: new InMemoryWorkbookAdapter({ revision: 1, sheets: [] }) },
    redactionIndexRef: { current: NO_REDACTIONS },
  } as never
}

/** A read context over the snapshot, withholding nothing unless asked to. */
function context(withheld: boolean): WorkbookReadContext {
  return {
    univerRef: { current: null },
    lazyWorkbookRef: { current: null },
    adapterRef: { current: new InMemoryWorkbookAdapter(SNAPSHOT) },
    redactionIndexRef: {
      current: withheld ? buildRedactionIndex(MARKS, SHEETS) : NO_REDACTIONS,
    },
  }
}

/** The text the naming bridge received, in order. */
let sent: string[] = []

/**
 * Install the naming bridge the preload exposes. Sheets' suite runs in a node
 * environment with no DOM, so the window object is created here rather than
 * borrowed from jsdom.
 */
function installBridge(bridge: {
  suggestFileName: (input: { content: string }) => Promise<unknown>
}): void {
  vi.stubGlobal('window', { aiNamingApi: bridge })
}

/** The bridge of a model that answers with a name. */
function answeringBridge(name = 'Q3 customer callbacks'): {
  suggestFileName: (input: { content: string }) => Promise<unknown>
} {
  return {
    suggestFileName: async (input) => {
      sent.push(input.content)
      return name ? { ok: true, name } : { ok: false, reason: 'empty-name' }
    },
  }
}

/** The save flow's own stub, with a naming bridge that must never be reached. */
const saveWorkbookEdits = vi.fn(async () => ({ canceled: true }))
const suggestFileName = vi.fn(async () => ({ ok: true, name: 'a name nobody asked for' }))

/** A save context with one edit pending, so the save reaches the IPC. */
function saveContext(): SaveContext {
  const journal = createEditJournal()
  recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'edited' } } })
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
  }
}

beforeEach(() => {
  sent = []
  saveWorkbookEdits.mockClear()
  suggestFileName.mockClear()
  installBridge(answeringBridge())
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

describe('the name the reader asked for', () => {
  it('is the one the model proposed', async () => {
    expect(await modelSaveAsName(context(false))).toBe('Q3 customer callbacks')
    expect(sent).toHaveLength(1)
  })

  it('never hands the model a value the reader withheld', async () => {
    await modelSaveAsName(context(true))
    // The one assertion that matters. The file name is written to disk, synced
    // and emailed, so a value that reached it would outlive the redaction meant
    // to contain it.
    expect(sent).toHaveLength(1)
    expect(sent[0]).not.toContain(SECRET)
    expect(JSON.stringify(sent[0])).not.toContain('13800138')
  })

  it('sends the marker the reader chose, so the model can still read the sheet', async () => {
    await modelSaveAsName(context(true))
    expect(sent[0]).toContain(`{{${LABEL}}}`)
    expect(sent[0]).toContain('\u5b63\u5ea6\u5ba2\u6237\u56de\u8bbf')
  })

  it('leaves the cells around a mark readable', async () => {
    // Over-redaction is its own bug: a workbook where everything reads as
    // withheld teaches the model nothing.
    await modelSaveAsName(context(true))
    expect(sent[0]).toContain('A1: \u5b63\u5ea6\u5ba2\u6237\u56de\u8bbf')
    expect(sent[0]).toContain('C1: 42')
  })

  it('reads an opened file off the grid, not the blank adapter snapshot', async () => {
    await modelSaveAsName(fileContext())
    expect(sent[0]).toContain('Sales')
    expect(sent[0]).toContain('A1: Region')
    expect(sent[0]).toContain('B2: 1200')
    expect(sent[0]).toContain('C1: FY26 Plan')
  })

  it('sends the whole workbook for one with nothing withheld', async () => {
    await modelSaveAsName(context(false))
    expect(sent[0]).toContain('B2: 13800138000')
    expect(sent[0]).toContain('widget')
  })

  it('reads the index the app keeps live, not a snapshot of it', async () => {
    // The point of reusing `redactionsOf` is that there is one index. A second
    // index built for naming would be correct until the reader added or
    // cleared a mark, and the drift is silent: the file would simply start
    // carrying a value the reader had hidden.
    const ctx = context(false)
    await modelSaveAsName(ctx)
    expect(sent[0]).toContain('B2: 13800138000')

    // exactly what the app does when a mark is confirmed
    const ref = ctx.redactionIndexRef as { current: ReturnType<typeof buildRedactionIndex> | null }
    ref.current = buildRedactionIndex(MARKS, SHEETS)
    await modelSaveAsName(ctx)
    expect(sent[1]).not.toContain(SECRET)
    expect(sent[1]).toContain(`{{${LABEL}}}`)
  })

  it('is repeatable: a second request asks again', async () => {
    // The reader may change their mind and click again — there is no latch
    // here, because there is no automatic attempt to ration.
    await modelSaveAsName(context(false))
    await modelSaveAsName(context(false))
    expect(sent).toHaveLength(2)
  })
})

describe('a naming that fails offers no name', () => {
  it('offers nothing when the model declines', async () => {
    installBridge({
      suggestFileName: async () => ({ ok: false, reason: 'no-provider' }),
    })
    expect(await modelSaveAsName(context(false))).toBe('')
  })

  it('offers nothing when the model returns an empty name', async () => {
    installBridge(answeringBridge(''))
    expect(await modelSaveAsName(context(false))).toBe('')
  })

  it('offers nothing when the call throws', async () => {
    installBridge({
      suggestFileName: async () => {
        throw new Error('sidecar unreachable')
      },
    })
    expect(await modelSaveAsName(context(false))).toBe('')
  })

  it('offers nothing when there is no naming bridge at all', async () => {
    vi.stubGlobal('window', {})
    expect(await modelSaveAsName(context(false))).toBe('')
  })

  it('offers nothing for a workbook with nothing in it', async () => {
    // An empty workbook has no name to derive, so there is nothing worth sending.
    const empty: WorkbookReadContext = {
      univerRef: { current: null },
      lazyWorkbookRef: { current: null },
      adapterRef: { current: new InMemoryWorkbookAdapter({ revision: 1, sheets: [] }) },
      redactionIndexRef: { current: NO_REDACTIONS },
    }
    expect(await modelSaveAsName(empty)).toBe('')
    expect(sent).toHaveLength(0)
  })
})

describe('nothing is named automatically', () => {
  it('leaves the naming module without an automatic entry point', () => {
    // The module's whole surface: one function, and it is the explicit action.
    // A latch and its re-arm hook would both be exports here.
    expect(Object.keys(fileNaming)).toEqual(['modelSaveAsName'])
  })

  it('never asks the model while saving', async () => {
    vi.stubGlobal('window', {
      desktopApi: { saveWorkbookEdits, writeWorkbookRecovery: vi.fn() },
      aiNamingApi: { suggestFileName },
    })
    await handleSaveImpl(saveContext(), 'save')
    expect(saveWorkbookEdits).toHaveBeenCalledTimes(1)
    expect(suggestFileName).not.toHaveBeenCalled()
  })

  it("keeps the model out of the app's save path", () => {
    // The save path lives in the App component, which no unit test can run, so
    // it is checked in the source: a naming call between the save and the
    // write is exactly what the reader asked us to take out.
    const appSrc = readFileSync(new URL('../src/renderer/App.tsx', import.meta.url), 'utf8')
    expect(appSrc).not.toMatch(/firstSaveName|noteDocumentSwapped/)
    const savePath = between(appSrc, 'async function handleSave(', 'async function nameWithAi(')
    expect(savePath).not.toMatch(/modelSaveAsName|suggestFileName|aiNamingApi/)
  })

  it('leaves the ribbon button as the only way in', () => {
    const appSrc = readFileSync(new URL('../src/renderer/App.tsx', import.meta.url), 'utf8')
    const shellSrc = readFileSync(
      new URL('../src/renderer/ExcelShell.tsx', import.meta.url),
      'utf8',
    )
    expect(shellSrc).toMatch(/onClick=\{onNameWithAi\}/)
    expect(appSrc).toContain('onNameWithAi={() => void nameWithAi()}')
    const handler = between(appSrc, 'async function nameWithAi(', 'closeSaveRef.current = async')
    expect(handler).toContain('modelSaveAsName(')
  })

  it('exposes one naming channel, and no preference for one', () => {
    // The preference and its two channels went with the automatic path: a name
    // is proposed because the reader asked for one.
    const preloadSrc = readFileSync(new URL('../src/preload/index.ts', import.meta.url), 'utf8')
    expect(preloadSrc).toContain("ipcRenderer.invoke('ai:suggest-file-name', input)")
    expect(preloadSrc).not.toMatch(
      /getFileNamingEnabled|setFileNamingEnabled|ai:get-file-naming|ai:set-file-naming/,
    )
  })
})

/** The part of `src` between two markers, both of which must be there. */
function between(src: string, from: string, to: string): string {
  const start = src.indexOf(from)
  const end = src.indexOf(to)
  expect(start, `marker not found: ${from}`).toBeGreaterThan(-1)
  expect(end, `marker not found: ${to}`).toBeGreaterThan(start)
  return src.slice(start, end)
}
