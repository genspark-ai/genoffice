/**
 * The load side: reading `xl/gxRedactions.json` at open, and what the renderer
 * does with each answer.
 *
 * The load has to tell three cases apart, and only the middle one is subtle:
 *
 * - **absent** — a workbook that withholds nothing. The common case, and it
 *   must cost one manifest lookup and nothing else.
 * - **ok** — the marks, ready to index.
 * - **unreadable** — the part exists and this version cannot make sense of it.
 *
 * The last one is the whole point of this file. Reading it as "no marks" is not
 * a degraded feature, it is a disclosure: `[]` answers "no cell is withheld"
 * to every read the model makes, so the model receives exactly the values the
 * reader hid, and nothing anywhere says so.
 */
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { readWorkbookRedactionPart } from '../src/main/workbook-redactions'
import type { ArchiveClient } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import { REDACTION_PART_PATH } from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import { NO_REDACTIONS } from '../src/renderer/ai/redact'
import { redactionSessionFor, WITHHELD_UNKNOWN } from '../src/renderer/ai/redact-load'
import { readCells, type WorkbookReadContext } from '../src/renderer/ai/workbook-readers'

const SHEETS = [
  { id: 'sh1', name: 'Customers' },
  { id: 'sh2', name: 'Orders' },
]

const MARK = { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: '客户电话' }

const PART = JSON.stringify({ version: 1, sheets: [{ name: 'Customers', marks: [MARK] }] })

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
  it('reports a workbook with no part as absent, not as an empty mark list', async () => {
    // The distinction is the whole contract: absent means "this file withholds
    // nothing", which is true, and is not the same claim as "we could not tell".
    const result = await readWorkbookRedactionPart(client(['[Content_Types].xml']), '/snap.xlsx')
    expect(result).toEqual({ status: 'absent' })
  })

  it('parses the marks when the part is there', async () => {
    const result = await readWorkbookRedactionPart(
      client([REDACTION_PART_PATH], { [REDACTION_PART_PATH]: PART }),
      '/snap.xlsx',
    )
    expect(result).toEqual({
      status: 'ok',
      states: [{ sheetName: 'Customers', marks: [MARK] }],
    })
  })

  it('refuses a corrupt part instead of reporting no marks', async () => {
    // Version 2 is a part this build does not understand. Returning `ok` with
    // zero states here is the leak: the model would read the marked cell.
    const future = JSON.stringify({ version: 2, sheets: [{ name: 'Customers', marks: [MARK] }] })
    const result = await readWorkbookRedactionPart(
      client([REDACTION_PART_PATH], { [REDACTION_PART_PATH]: future }),
      '/snap.xlsx',
    )
    expect(result.status).toBe('unreadable')
    if (result.status !== 'unreadable') throw new Error('expected an unreadable result')
    expect(result.error).toContain('version')
    // The marks are in the part; nothing may report them as absent.
    expect(JSON.stringify(result)).not.toContain('absent')
  })

  it('refuses a part that is not JSON at all', async () => {
    const result = await readWorkbookRedactionPart(
      client([REDACTION_PART_PATH], { [REDACTION_PART_PATH]: '{ not json' }),
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

describe('what the renderer installs for each answer', () => {
  it('keeps an absent part on the shared empty index', () => {
    const session = redactionSessionFor({ status: 'absent' }, SHEETS)
    expect(session.index).toBe(NO_REDACTIONS)
    expect(session.error).toBeNull()
  })

  it('indexes the marks by sheet id when the part reads back', () => {
    const session = redactionSessionFor(
      { status: 'ok', states: [{ sheetName: 'Customers', marks: [MARK] }] },
      SHEETS,
    )
    expect(session.error).toBeNull()
    expect(session.index.labelAt('sh1', 1, 1)).toBe('客户电话')
    expect(session.index.labelAt('sh2', 1, 1)).toBeNull()
  })

  it('withholds every cell when the part could not be read, and says why', () => {
    const session = redactionSessionFor({ status: 'unreadable', error: 'version 2' }, SHEETS)
    expect(session.error).toBe('version 2')
    // Not empty: an empty index is what would let the reads through.
    expect(session.index.isEmpty).toBe(false)
    expect(session.index).toBe(WITHHELD_UNKNOWN)
  })

  it('shows the model a placeholder instead of the withheld value after a bad read', () => {
    // The end-to-end consequence, stated as a test because it is the whole
    // reason the state exists: with the part unreadable, read_cells hands over
    // a placeholder, never 13800138000.
    const session = redactionSessionFor({ status: 'unreadable', error: 'version 2' }, SHEETS)
    const result = readCells(context(session.index), ['B2'])
    expect(result.B2).toEqual({ value: '{{private}}' })
    expect(JSON.stringify(result)).not.toContain('13800138000')
  })

  it('leaves a workbook that withholds nothing exactly as it was', () => {
    const session = redactionSessionFor({ status: 'absent' }, SHEETS)
    const result = readCells(context(session.index), ['B2'])
    expect(result.B2).toEqual({ value: '13800138000' })
  })
})
