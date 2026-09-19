import { describe, expect, it } from 'vitest'

import { csvToXlsxBuffer } from '../src/gateway/csv-import'
import {
  applyCellEditsToXlsx,
  assertOnlyTouchedEntriesChanged,
  createBufferEntrySource,
  inventoryXlsx,
  readBasicWorkbook,
} from '../src/gateway/xlsx-gateway'

async function sheetCells(
  buffer: Buffer,
): Promise<Readonly<Record<string, { value: unknown; formula?: string | undefined }>>> {
  const imported = await readBasicWorkbook(buffer)
  const sheet = imported.snapshot.sheets[0]
  if (!sheet) throw new Error('Imported workbook has no sheets.')
  return sheet.cells
}

async function sheetXml(buffer: Buffer): Promise<string> {
  const source = await createBufferEntrySource(buffer)
  return source.readText('xl/worksheets/sheet1.xml')
}

function shaByPath(entries: readonly { path: string; sha256: string }[]): Map<string, string> {
  return new Map(entries.map((entry) => [entry.path, entry.sha256]))
}

describe('xlsx-gateway csv import round trip', () => {
  it('lands commas, quotes, and escaped quotes in the expected cells', async () => {
    const csv = 'name,note,age\n"Doe, John","says ""hi""",30\nAlice,plain,007'
    const cells = await sheetCells(await csvToXlsxBuffer(csv))

    expect(cells['A1']?.value).toBe('name')
    expect(cells['B1']?.value).toBe('note')
    expect(cells['C1']?.value).toBe('age')
    expect(cells['A2']?.value).toBe('Doe, John')
    expect(cells['B2']?.value).toBe('says "hi"')
    expect(cells['C2']?.value).toBe(30)
    expect(cells['A3']?.value).toBe('Alice')
    expect(cells['B3']?.value).toBe('plain')
    // Leading zeros survive the import as text.
    expect(cells['C3']?.value).toBe('007')
  })

  it('preserves unicode and multiline quoted fields', async () => {
    const csv =
      'id,comment\n1,"caf\u00E9 \u2603 \u2014 na\u00EFve"\n2,"line one\nline two"\n3,"emoji \u{1F389} test"'
    const cells = await sheetCells(await csvToXlsxBuffer(csv))

    expect(cells['B2']?.value).toBe('caf\u00E9 \u2603 \u2014 na\u00EFve')
    expect(cells['B3']?.value).toBe('line one\nline two')
    expect(cells['B4']?.value).toBe('emoji \u{1F389} test')
    expect(cells['A2']?.value).toBe(1)
    expect(cells['A3']?.value).toBe(2)
  })

  it('keeps numbers numeric without corrupting codes or long ids', async () => {
    const csv = 'code,value\n007,007\n00123,1.5\n12345678901234567890,-42\n1e3,1E-2\n-42.5,0'
    const cells = await sheetCells(await csvToXlsxBuffer(csv))

    expect(cells['A2']?.value).toBe('007')
    expect(cells['B2']?.value).toBe('007')
    expect(cells['A3']?.value).toBe('00123')
    expect(cells['B3']?.value).toBe(1.5)
    // Past Excel's 15-digit precision the id must stay text.
    expect(cells['A4']?.value).toBe('12345678901234567890')
    expect(cells['B4']?.value).toBe(-42)
    expect(cells['A5']?.value).toBe(1000)
    expect(cells['B5']?.value).toBe(0.01)
    expect(cells['A6']?.value).toBe(-42.5)
    expect(cells['B6']?.value).toBe(0)
  })

  it('keeps formula-like text as text instead of a stored formula', async () => {
    const csv = 'expr,note\n"=SUM(A1:A2)",plain'
    const cells = await sheetCells(await csvToXlsxBuffer(csv))

    expect(cells['A2']?.value).toBe('=SUM(A1:A2)')
    expect(cells['A2']?.formula).toBeUndefined()
    expect(cells['B2']?.value).toBe('plain')
  })

  it('is deterministic: the same csv yields identical entry bytes and cells', async () => {
    // No csv export path exists at this base, so the round-trip proof is
    // import determinism at the entry-payload level (zip container metadata
    // such as file timestamps is not part of the file-format promise).
    const csv =
      'name,note,age\n"Doe, John","says ""hi"" caf\u00E9",30\nAlice,"line one\nline two",007'
    const first = await csvToXlsxBuffer(csv)
    const second = await csvToXlsxBuffer(csv)

    const firstEntries = await inventoryXlsx(first)
    const secondEntries = await inventoryXlsx(second)
    expect(secondEntries.map((entry) => entry.path)).toEqual(
      firstEntries.map((entry) => entry.path),
    )
    expect(shaByPath(secondEntries)).toEqual(shaByPath(firstEntries))
    expect(await sheetXml(second)).toBe(await sheetXml(first))

    const firstCells = await sheetCells(first)
    const secondCells = await sheetCells(second)
    expect(secondCells).toEqual(firstCells)
    expect(secondCells['A2']?.value).toBe('Doe, John')
    expect(secondCells['B3']?.value).toBe('line one\nline two')
  })

  it('leaves untouched package entries byte-identical on a no-op save', async () => {
    const csv = 'name,note\n"Doe, John","caf\u00E9 \u2603"\nAlice,"line one\nline two"'
    const source = await csvToXlsxBuffer(csv)
    const before = await inventoryXlsx(source)
    const beforeByPath = shaByPath(before)

    const mutation = await applyCellEditsToXlsx(source, [])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()

    const touched = new Set(mutation.touchedEntries)
    for (const entry of mutation.afterEntries) {
      if (touched.has(entry.path)) continue
      expect(entry.sha256).toBe(beforeByPath.get(entry.path))
    }

    const beforeCells = await sheetCells(source)
    const afterCells = await sheetCells(mutation.buffer)
    expect(afterCells).toEqual(beforeCells)
  })
})
