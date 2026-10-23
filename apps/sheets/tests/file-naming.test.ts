import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import type { WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { SheetRedactionState } from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import { buildRedactionIndex, NO_REDACTIONS } from '../src/renderer/ai/redact'
import type { WorkbookReadContext } from '../src/renderer/ai/workbook-readers'
import { firstSaveName, modelSaveAsName, noteDocumentSwapped } from '../src/renderer/file-naming'

/**
 * These go through the naming functions the save path and the ribbon actually
 * call, not through `nameForSave` directly. A test of the shared client would
 * still pass if this app stopped using the projection — which is the failure
 * that leaks a value, so the app's own call site is what needs covering.
 *
 * The context is built with the app's own `WorkbookReadContext` shape and the
 * real `InMemoryWorkbookAdapter`, because the module reads the workbook through
 * the same ref the AI panel does. A hand-rolled stand-in would let the test
 * pass while the app kept handing the model raw cell values.
 */

const LABEL = '客户电话'
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
        A1: { value: '季度客户回访' },
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

/** Calls the naming bridge received, in order. */
let sent: Array<{ content: string; trigger: string; filePath: string | null }> = []

/** What the model answers with; '' is a model that returned nothing usable. */
let answer = 'Q3 customer callbacks'
/** What the preference reads as. First-save naming consults it; manual does not. */
let namingEnabled = true

/**
 * Install the naming bridge the preload exposes. Sheets' suite runs in a node
 * environment with no DOM, so the window object is created here rather than
 * borrowed from jsdom.
 */
function installBridge(api: {
  suggestFileName: (input: {
    content: string
    trigger: string
    filePath?: string | null
  }) => Promise<unknown>
  getFileNamingEnabled: () => Promise<boolean>
}): void {
  vi.stubGlobal('window', { aiOffice: api })
}

beforeEach(() => {
  sent = []
  answer = 'Q3 customer callbacks'
  namingEnabled = true
  // The latch is module state, deliberately: one attempt per document. Every
  // test starts from a freshly swapped document so the cases stay independent.
  noteDocumentSwapped()
  installBridge({
    suggestFileName: async (input) => {
      sent.push({
        content: input.content,
        trigger: input.trigger,
        filePath: input.filePath ?? null,
      })
      return answer ? { ok: true, name: answer } : { ok: false, reason: 'empty-name' }
    },
    getFileNamingEnabled: async () => namingEnabled,
  })
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

describe('the first-save name is derived from the projected text', () => {
  it('never hands the model a value the reader withheld', async () => {
    await firstSaveName(context(true))
    // The one assertion that matters. The file name is written to disk, synced
    // and emailed, so a value that reached it would outlive the redaction meant
    // to contain it.
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(JSON.stringify(sent[0]!.content)).not.toContain('13800138')
  })

  it('sends the marker the reader chose, so the model can still read the sheet', async () => {
    await firstSaveName(context(true))
    expect(sent[0]!.content).toContain(`{{${LABEL}}}`)
    expect(sent[0]!.content).toContain('季度客户回访')
  })

  it('leaves the cells around a mark readable', async () => {
    // Over-redaction is its own bug: a workbook where everything reads as
    // withheld teaches the model nothing.
    await firstSaveName(context(true))
    expect(sent[0]!.content).toContain('A1: 季度客户回访')
    expect(sent[0]!.content).toContain('C1: 42')
  })

  it('sends the whole workbook for one with nothing withheld', async () => {
    await firstSaveName(context(false))
    expect(sent[0]!.content).toContain('B2: 13800138000')
    expect(sent[0]!.content).toContain('widget')
  })

  it('asks as a first-save attempt, which is the one the preference governs', async () => {
    await firstSaveName(context(false))
    expect(sent[0]!.trigger).toBe('first-save')
  })

  it('reads the index the app keeps live, not a snapshot of it', async () => {
    // The point of reusing `redactionsOf` is that there is one index. A second
    // index built for naming would be correct until the reader added or
    // cleared a mark, and the drift is silent: the file would simply start
    // carrying a value the reader had hidden.
    const ctx = context(false)
    await firstSaveName(ctx)
    expect(sent[0]!.content).toContain('B2: 13800138000')

    // exactly what the app does when a mark is confirmed
    const ref = ctx.redactionIndexRef as { current: ReturnType<typeof buildRedactionIndex> | null }
    ref.current = buildRedactionIndex(MARKS, SHEETS)
    noteDocumentSwapped()
    await firstSaveName(ctx)
    expect(sent[1]!.content).not.toContain(SECRET)
    expect(sent[1]!.content).toContain(`{{${LABEL}}}`)
  })
})

describe('the first-save name falls back to the workbook own sheets', () => {
  it('proposes nothing when naming is declined', async () => {
    // Sheets' own name for an untitled workbook is the shell's untitled default,
    // retargeted through autoRenameWorkbook, so the fallback here is an empty
    // proposal: the save keeps the name it would have had anyway.
    namingEnabled = false
    expect(await firstSaveName(context(false))).toBe('')
    expect(sent).toHaveLength(0)
  })

  it('proposes nothing when the model errors', async () => {
    installBridge({
      suggestFileName: async () => {
        throw new Error('sidecar unreachable')
      },
      getFileNamingEnabled: async () => true,
    })
    expect(await firstSaveName(context(false))).toBe('')
  })

  it('proposes nothing when the model times out with nothing to show', async () => {
    // The main process aborts a naming turn at 30s and answers ok:false, so a
    // timeout reaches the app as a decline rather than as a throw.
    installBridge({
      suggestFileName: async () => ({ ok: false, reason: 'timeout' }),
      getFileNamingEnabled: async () => true,
    })
    expect(await firstSaveName(context(false))).toBe('')
  })

  it('proposes nothing when the model returns an empty name', async () => {
    answer = ''
    expect(await firstSaveName(context(false))).toBe('')
  })

  it('proposes nothing when there is no naming bridge at all', async () => {
    vi.stubGlobal('window', {})
    expect(await firstSaveName(context(false))).toBe('')
  })

  it('proposes nothing for a workbook with no cells to work from', async () => {
    // An empty workbook has no name to derive, so there is nothing worth sending.
    await firstSaveName({
      univerRef: { current: null },
      lazyWorkbookRef: { current: null },
      adapterRef: { current: new InMemoryWorkbookAdapter({ revision: 1, sheets: [] }) },
      redactionIndexRef: { current: NO_REDACTIONS },
    })
    expect(sent).toHaveLength(0)
  })

  it('proposes the model name when one does arrive', async () => {
    expect(await firstSaveName(context(false))).toBe('Q3 customer callbacks')
  })
})

describe('one naming turn per document', () => {
  it('does not ask again on a second save', async () => {
    expect(await firstSaveName(context(false))).toBe('Q3 customer callbacks')
    expect(await firstSaveName(context(false))).toBe('')
    expect(sent).toHaveLength(1)
  })

  it('does not start a second call while the first is still in flight', async () => {
    // Two saves overlapping is what a second ⌘S during a slow turn looks like.
    const [first, second] = await Promise.all([
      firstSaveName(context(false)),
      firstSaveName(context(false)),
    ])
    expect(sent).toHaveLength(1)
    expect(first).toBe('Q3 customer callbacks')
    expect(second).toBe('')
  })

  it('asks again once a different workbook is loaded', async () => {
    expect(await firstSaveName(context(false))).toBe('Q3 customer callbacks')
    noteDocumentSwapped()
    expect(await firstSaveName(context(false))).toBe('Q3 customer callbacks')
    expect(sent).toHaveLength(2)
  })
})

describe('the manual name goes to the Save As dialog, projected', () => {
  it('never hands the model a value the reader withheld', async () => {
    await modelSaveAsName(context(true), '/tmp/book.xlsx')
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(sent[0]!.content).toContain(`{{${LABEL}}}`)
  })

  it('tells the main process which file is being renamed, so it can refuse a real one', async () => {
    await modelSaveAsName(context(false), '/tmp/book.xlsx')
    expect(sent[0]!.filePath).toBe('/tmp/book.xlsx')
    expect(sent[0]!.trigger).toBe('manual')
  })

  it('goes through even with first-save naming switched off', async () => {
    // The reader asked for it by clicking, so the preference does not apply.
    namingEnabled = false
    expect(await modelSaveAsName(context(false), '/tmp/book.xlsx')).toBe('Q3 customer callbacks')
    expect(sent).toHaveLength(1)
  })

  it('returns nothing when the model declines, leaving the dialog on the current name', async () => {
    answer = ''
    expect(await modelSaveAsName(context(false), '/tmp/book.xlsx')).toBe('')
  })

  it('is repeatable: a second request asks again', async () => {
    // The latch is a first-save promise, not a rate limit — a reader who
    // changed their mind and clicked again gets a fresh answer.
    await modelSaveAsName(context(false), '/tmp/book.xlsx')
    await modelSaveAsName(context(false), '/tmp/book.xlsx')
    expect(sent).toHaveLength(2)
  })
})
