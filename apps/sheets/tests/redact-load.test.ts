/**
 * The load side: reading `xl/gxRedactions.json` at open — or, when that part is
 * gone, the copy of the same marks in the workbook's defined names — and what
 * the renderer does with each answer.
 *
 * The load has to tell three cases apart, and only the middle one is subtle:
 *
 * - **absent** — a workbook that withholds nothing. The common case.
 * - **ok** — the marks, ready to index.
 * - **unreadable** — the record exists and this version cannot make sense of it.
 *
 * The last one is the whole point of this file. Reading it as "no marks" is not
 * a degraded feature, it is a disclosure: `[]` answers "no cell is withheld"
 * to every read the model makes, so the model receives exactly the values the
 * reader hid, and nothing anywhere says so. The defined-name fallback exists so
 * that a file which has been through another spreadsheet program lands on `ok`
 * rather than `unreadable`, without ever weakening that rule.
 */
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { readWorkbookRedactionPart } from '../src/main/workbook-redactions'
import type { ArchiveClient } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import {
  REDACTION_PART_PATH,
  applyRedactionNameMirror,
  type SheetRedactionState,
} from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import { NO_REDACTIONS } from '../src/renderer/ai/redact'
import { redactionSessionFor, WITHHELD_UNKNOWN } from '../src/renderer/ai/redact-load'
import { readCells, type WorkbookReadContext } from '../src/renderer/ai/workbook-readers'
import { workbookRedactionsResultSchema } from '../src/shared/desktop-api'
import type { TFunc } from '../src/renderer/i18n/locale'

const SHEETS = [
  { id: 'sh1', name: 'Customers' },
  { id: 'sh2', name: 'Orders' },
]
const SHEET_ORDER = ['Customers', 'Orders']

const MARK = { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }
const STATES: SheetRedactionState[] = [{ sheetName: 'Customers', marks: [MARK] }]

const PART = JSON.stringify({ version: 1, sheets: [{ name: 'Customers', marks: [MARK] }] })

const WORKBOOK_XML =
  '<workbook xmlns:r="urn:r"><sheets>' +
  '<sheet name="Customers" sheetId="1" r:id="rId1"/>' +
  '<sheet name="Orders" sheetId="2" r:id="rId2"/>' +
  '</sheets></workbook>'

/** A workbook that has been saved with the marks mirrored into its defined names. */
const MIRRORED = applyRedactionNameMirror(WORKBOOK_XML, STATES, SHEET_ORDER)

/** The same thing after a reader rebuilt the package and dropped the part. */
const CLEAN_WORKBOOK_XML = MIRRORED

/**
 * An archive client that answers from a fixed list of entry names.
 *
 * `readArchiveEntryText` reads the extracted file off disk, so the texts are
 * written to a real temp dir rather than returned inline.
 */
function client(entries: readonly string[], texts: Record<string, string> = {}): ArchiveClient {
  const dir = mkdtempSync(join(tmpdir(), 'redaction-load-'))
  return {
    archiveManifest: () => Promise.resolve({ entries: entries.map((name) => ({ name })) }),
    readEntries: ({ entries: wanted }: { entries: readonly string[] }) => {
      const name = wanted[0]
      const text = name === undefined ? undefined : texts[name]
      if (name === undefined || text === undefined) {
        return Promise.reject(new Error(`Workbook is missing ${name ?? ''}.`))
      }
      const extracted = join(dir, name.replace(/[\\/]/g, '_'))
      writeFileSync(extracted, text)
      return Promise.resolve({ entries: [{ name, path: extracted }] })
    },
    scanEntries: () => Promise.resolve({ matches: [] }),
    saveArchive: () => Promise.resolve({ beforeEntries: [], afterEntries: [] }),
  } as unknown as ArchiveClient
}

/** A snapshot of an ordinary workbook: workbook.xml, and no redaction part. */
function cleanSnapshot(workbookXml = WORKBOOK_XML): ArchiveClient {
  return client(['[Content_Types].xml', 'xl/workbook.xml'], { 'xl/workbook.xml': workbookXml })
}

/// The renderer's translate function, stubbed to the key so tests can read it.
const t = ((key: string) => key) as TFunc

/**
 * The hop the answer really makes: main hands its `RedactionPartRead` to the
 * IPC schema, which is the shape the renderer is typed against. Running every
 * recovery through it keeps the schema honest — a `source` the schema does not
 * declare would throw here rather than in a user's session.
 */
const overIpc = workbookRedactionsResultSchema.parse

/// A demo-mode read context carrying one index, holding the real value in B2.
function context(index: ReturnType<typeof redactionSessionFor>['index']): WorkbookReadContext {
  const worksheet = {
    getSheetId: () => 'sh1',
    getSheetName: () => 'Customers',
    getRange: () => ({ getValue: () => null }),
  }
  return {
    univerRef: {
      current: {
        univerAPI: {
          getActiveWorkbook: () => ({
            getActiveSheet: () => worksheet,
            getSheetBySheetId: () => worksheet,
            getActiveRange: () => null,
            getSheets: () => [worksheet],
          }),
        },
      },
    } as never,
    lazyWorkbookRef: { current: null },
    adapterRef: {
      current: {
        getSnapshot: () => ({
          revision: 1,
          sheets: [{ id: 'sh1', name: 'Customers', cells: { B2: { value: '13800138000' } } }],
        }),
      },
    } as never,
    redactionIndexRef: { current: index },
  }
}

describe('reading the redaction part out of a session snapshot', () => {
  it('reports a workbook with no part and no mirror as absent', async () => {
    // The distinction is the whole contract: absent means "this file withholds
    // nothing", which is true, and is not the same claim as "we could not tell".
    const result = await readWorkbookRedactionPart(cleanSnapshot(), '/snap.xlsx')
    expect(result).toEqual({ status: 'absent' })
  })

  it('parses the marks when the part is there', async () => {
    const result = await readWorkbookRedactionPart(
      client([REDACTION_PART_PATH, 'xl/workbook.xml'], {
        [REDACTION_PART_PATH]: PART,
        'xl/workbook.xml': MIRRORED,
      }),
      '/snap.xlsx',
    )
    expect(result).toEqual({ status: 'ok', states: STATES, source: 'part' })
  })

  it('refuses a corrupt part when there is nothing to fall back on', async () => {
    // Version 2 is a part this build does not understand. Returning `ok` with
    // zero states here is the leak: the model would read the marked cell.
    const future = JSON.stringify({ version: 2, sheets: [{ name: 'Customers', marks: [MARK] }] })
    const result = await readWorkbookRedactionPart(
      client([REDACTION_PART_PATH, 'xl/workbook.xml'], {
        [REDACTION_PART_PATH]: future,
        'xl/workbook.xml': WORKBOOK_XML,
      }),
      '/snap.xlsx',
    )
    expect(result.status).toBe('unreadable')
    if (result.status !== 'unreadable') throw new Error('expected an unreadable result')
    expect(result.error).toContain('version')
    // The marks are in the part; nothing may report them as absent.
    expect(JSON.stringify(result)).not.toContain('absent')
  })

  it('refuses a part that is not JSON at all, with no mirror either', async () => {
    const result = await readWorkbookRedactionPart(
      client([REDACTION_PART_PATH, 'xl/workbook.xml'], {
        [REDACTION_PART_PATH]: '{ not json',
        'xl/workbook.xml': WORKBOOK_XML,
      }),
      '/snap.xlsx',
    )
    expect(result.status).toBe('unreadable')
  })

  it('refuses rather than guessing when the part cannot be read back', async () => {
    // The manifest says it exists, so "missing" cannot be the answer — and a
    // read that failed is not evidence that nothing is withheld.
    const result = await readWorkbookRedactionPart(client([REDACTION_PART_PATH]), '/snap.xlsx')
    expect(result.status).toBe('unreadable')
  })
})

describe('when the part is gone — the LibreOffice case', () => {
  it('rebuilds the marks from the defined names', async () => {
    // `soffice --convert-to xlsx` rebuilds the package and drops the part, its
    // relationship, and its content-type override. workbook.xml is the one file
    // it keeps, so it is the one place the marks can still be found.
    const result = await readWorkbookRedactionPart(cleanSnapshot(CLEAN_WORKBOOK_XML), '/snap.xlsx')
    expect(result).toEqual({ status: 'ok', states: STATES, source: 'mirror' })
  })

  it('keeps the marked cell out of the model after the recovery', async () => {
    // The end-to-end consequence: a file that went through another program is
    // still protected. The value behind the mark never reaches a read — and the
    // placeholder carries the label, so the model knows what it lost.
    const result = await readWorkbookRedactionPart(cleanSnapshot(CLEAN_WORKBOOK_XML), '/snap.xlsx')
    const session = redactionSessionFor(overIpc(result), SHEETS, t)
    const read = readCells(context(session.index), ['B2'])
    expect(read.B2).toEqual({ value: '{{client phone}}' })
    expect(JSON.stringify(read)).not.toContain('13800138000')
  })

  it('still withholds everything when the mirrored names cannot be read', async () => {
    // A mirror is a record of what was withheld too. Damaging it must not turn
    // into "nothing is withheld" — that is the same leak, one file further along.
    const damaged = MIRRORED.replace('{""v"":1', '{""v"":2')
    const result = await readWorkbookRedactionPart(cleanSnapshot(damaged), '/snap.xlsx')
    expect(result.status).toBe('unreadable')
  })

  it('prefers the part whenever it can be read', async () => {
    // Both copies hold the same marks, but the part is the record: if the
    // mirror has fallen behind, the part is the newer of the two.
    const stale = applyRedactionNameMirror(
      WORKBOOK_XML,
      [{ sheetName: 'Orders', marks: [{ ...MARK, label: 'stale' }] }],
      SHEET_ORDER,
    )
    const result = await readWorkbookRedactionPart(
      client([REDACTION_PART_PATH, 'xl/workbook.xml'], {
        [REDACTION_PART_PATH]: PART,
        'xl/workbook.xml': stale,
      }),
      '/snap.xlsx',
    )
    expect(result).toEqual({ status: 'ok', states: STATES, source: 'part' })
  })

  it('falls back to the mirror when the part is present but unreadable', async () => {
    const result = await readWorkbookRedactionPart(
      client([REDACTION_PART_PATH, 'xl/workbook.xml'], {
        [REDACTION_PART_PATH]: '{ not json',
        'xl/workbook.xml': CLEAN_WORKBOOK_XML,
      }),
      '/snap.xlsx',
    )
    expect(result).toEqual({ status: 'ok', states: STATES, source: 'mirror' })
  })

  it('resolves the mirror against the sheet order in the file, not the name it had', async () => {
    // The mirror records a position, so the mark follows whichever sheet is at
    // that position. A rename moves the mark with the table instead of
    // orphaning it — the trade the mirror makes, stated as a test.
    const renamed = applyRedactionNameMirror(
      WORKBOOK_XML.replace('name="Customers"', 'name="Clients"'),
      STATES,
      SHEET_ORDER,
    )
    const result = await readWorkbookRedactionPart(cleanSnapshot(renamed), '/snap.xlsx')
    expect(result).toEqual({
      status: 'ok',
      source: 'mirror',
      states: [{ sheetName: 'Clients', marks: [MARK] }],
    })
  })
})

describe('what the renderer installs for each answer', () => {
  it('keeps an absent part on the shared empty index', () => {
    const session = redactionSessionFor({ status: 'absent' }, SHEETS, t)
    expect(session.index).toBe(NO_REDACTIONS)
    expect(session.error).toBeNull()
    expect(session.notice).toBeNull()
  })

  it('indexes the marks by sheet id when the part reads back', () => {
    const session = redactionSessionFor(
      overIpc({ status: 'ok', states: STATES, source: 'part' }),
      SHEETS,
      t,
    )
    expect(session.error).toBeNull()
    expect(session.index.labelAt('sh1', 1, 1)).toBe('client phone')
    expect(session.index.labelAt('sh2', 1, 1)).toBeNull()
  })

  it('indexes the same marks when they came from the mirror, and says so', () => {
    // The reader is told the file has been through another program, because
    // saving here is what puts the part back — an action they cannot guess at.
    const session = redactionSessionFor(
      overIpc({ status: 'ok', states: STATES, source: 'mirror' }),
      SHEETS,
      t,
    )
    expect(session.error).toBeNull()
    expect(session.notice).toBe('redactRecoveredFromNames')
    expect(session.index.labelAt('sh1', 1, 1)).toBe('client phone')
  })

  it('withholds every cell when the part could not be read, and says why', () => {
    const session = redactionSessionFor({ status: 'unreadable', error: 'version 2' }, SHEETS, t)
    expect(session.error).toBe('version 2')
    // Not empty: an empty index is what would let the reads through.
    expect(session.index.isEmpty).toBe(false)
    expect(session.index).toBe(WITHHELD_UNKNOWN)
  })

  it('shows the model a placeholder instead of the withheld value after a bad read', () => {
    // The end-to-end consequence, stated as a test because it is the whole
    // reason the state exists: with the part unreadable, read_cells hands over
    // a placeholder, never 13800138000.
    const session = redactionSessionFor({ status: 'unreadable', error: 'version 2' }, SHEETS, t)
    const result = readCells(context(session.index), ['B2'])
    expect(result.B2).toEqual({ value: '{{private}}' })
    expect(JSON.stringify(result)).not.toContain('13800138000')
  })

  it('leaves a workbook that withholds nothing exactly as it was', () => {
    const session = redactionSessionFor({ status: 'absent' }, SHEETS, t)
    const result = readCells(context(session.index), ['B2'])
    expect(result.B2).toEqual({ value: '13800138000' })
  })
})
