import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyCellEditsToXlsx,
  assertOnlyTouchedEntriesChanged,
  toA1Address,
  type CellEdit,
} from '../src/gateway/xlsx-gateway'
import { blankXlsxBuffer } from '../src/gateway/csv-import'
import { buildEditFixture } from './fixture-builder'

describe('toA1Address', () => {
  it('converts zero-based coordinates to A1 notation', () => {
    expect(toA1Address(0, 0)).toBe('A1')
    expect(toA1Address(2, 1)).toBe('B3')
    expect(toA1Address(0, 25)).toBe('Z1')
    expect(toA1Address(0, 26)).toBe('AA1')
    expect(toA1Address(9, 701)).toBe('ZZ10')
    expect(toA1Address(9, 702)).toBe('AAA10')
  })
})

describe('applyCellEditsToXlsx', () => {
  it('keeps the style index when overwriting a styled shared-string cell', async () => {
    const worksheet = await editedWorksheet([edit(0, 0, { value: 'World' })])
    expect(worksheet).toContain(
      '<c r="A1" s="1" t="inlineStr"><is><t xml:space="preserve">World</t></is></c>',
    )
    expect(worksheet).not.toContain('t="s"')
  })

  it('leaves sharedStrings and unrelated parts byte-identical', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      edit(0, 0, { value: 'World' }),
    ])
    expect(mutation.touchedEntries).toEqual(['xl/workbook.xml', 'xl/worksheets/sheet1.xml'])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const before = new Map(mutation.beforeEntries.map((entry) => [entry.path, entry.sha256]))
    const after = new Map(mutation.afterEntries.map((entry) => [entry.path, entry.sha256]))
    expect(after.get('xl/sharedStrings.xml')).toBe(before.get('xl/sharedStrings.xml'))
    expect(after.get('customXml/item1.xml')).toBe(before.get('customXml/item1.xml'))
    expect(after.get('xl/charts/chart1.xml')).toBe(before.get('xl/charts/chart1.xml'))
  })

  it('inserts a new cell in column order within an existing row', async () => {
    const worksheet = await editedWorksheet([edit(0, 1, { value: 7 })])
    const row = /<row r="1">([\s\S]*?)<\/row>/.exec(worksheet)?.[1] ?? ''
    const columns = [...row.matchAll(/r="([A-Z]+)1"/g)].map((match) => match[1])
    expect(columns).toEqual(['A', 'B', 'C'])
  })

  it('inserts a new row between existing rows in row order', async () => {
    const worksheet = await editedWorksheet([edit(1, 0, { value: 'between' })])
    const rows = [...worksheet.matchAll(/<row r="([0-9]+)"/g)].map((match) => Number(match[1]))
    expect(rows).toEqual([1, 2, 3, 5])
  })

  it('expands a self-closing row and keeps its attributes', async () => {
    const worksheet = await editedWorksheet([edit(4, 0, { value: 9 })])
    expect(worksheet).toContain('<row r="5" ht="20" customHeight="1"><c r="A5"><v>9</v></c></row>')
  })

  it('keeps a styled cell as an empty styled cell when cleared', async () => {
    const worksheet = await editedWorksheet([edit(0, 0, { value: null })])
    expect(worksheet).toContain('<c r="A1" s="1"/>')
  })

  it('removes an unstyled cell entirely when cleared', async () => {
    const worksheet = await editedWorksheet([edit(0, 2, { value: null })])
    expect(worksheet).not.toContain('r="C1"')
  })

  it('replaces a formula and drops the stale cached value', async () => {
    const worksheet = await editedWorksheet([edit(2, 1, { value: null, formula: '=SUM(C1:C3)' })])
    expect(worksheet).toContain('<c r="B3"><f>SUM(C1:C3)</f></c>')
  })

  it('does not expand replacement-pattern text like $& into backreferences', async () => {
    const worksheet = await editedWorksheet([edit(0, 0, { value: 'cost $1 & $& more' })])
    expect(worksheet).toContain('cost $1 &amp; $&amp; more')
  })

  it('forces a full recalculation on the next Excel open', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      edit(0, 2, { value: 6 }),
    ])
    const workbookXml = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbookXml).toContain('</definedNames><calcPr fullCalcOnLoad="1"/>')
  })

  it('writes boolean and numeric scalars with the right cell types', async () => {
    const worksheet = await editedWorksheet([
      edit(9, 0, { value: true }),
      edit(9, 1, { value: 3.5 }),
    ])
    expect(worksheet).toContain('<c r="A10" t="b"><v>1</v></c>')
    expect(worksheet).toContain('<c r="B10"><v>3.5</v></c>')
  })

  it('fails closed when the sheet does not exist', async () => {
    await expect(
      applyCellEditsToXlsx(await buildEditFixture(), [
        { sheetName: 'Missing', row: 0, column: 0, writeValue: true, cell: { value: 1 } },
      ]),
    ).rejects.toThrow('was not found')
  })
})

describe('applyCellEditsToXlsx style edits', () => {
  it('creates and registers a stylesheet when the workbook has none', async () => {
    const mutation = await applyCellEditsToXlsx(await blankXlsxBuffer(), [
      {
        sheetName: 'Sheet1',
        row: 0,
        column: 0,
        writeValue: true,
        cell: { value: 'Header' },
        style: { bold: true, fillColor: '#D9EAF7' },
      },
    ])

    expect(mutation.addedEntries).toContain('xl/styles.xml')
    expect(await entryText(mutation.buffer, 'xl/styles.xml')).toContain('<cellXfs count="2">')
    expect(await entryText(mutation.buffer, 'xl/_rels/workbook.xml.rels')).toContain(
      'relationships/styles',
    )
    expect(await entryText(mutation.buffer, '[Content_Types].xml')).toContain(
      'PartName="/xl/styles.xml"',
    )
  })

  it('persists a styled table written into an empty workbook', async () => {
    const values = [
      ['Name', 'Gender', 'Age', 'Phone'],
      ['Zhang San', 'Male', 28, '138-0000-0001'],
      ['Li Si', 'Female', 32, '138-0000-0002'],
    ]
    const edits: CellEdit[] = values.flatMap((row, rowIndex) =>
      row.map((value, column) => ({
        sheetName: 'Sheet1',
        row: rowIndex,
        column,
        writeValue: true,
        cell: { value },
        ...(rowIndex === 0 ? { style: { bold: true, fillColor: '#D9EAF7' } } : {}),
      })),
    )
    const mutation = await applyCellEditsToXlsx(await blankXlsxBuffer(), edits)
    const worksheet = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')

    expect(worksheet).toContain('<dimension ref="A1:D3"/>')
    for (const [rowIndex, row] of values.entries()) {
      for (const [column, value] of row.entries()) {
        const address = `${String.fromCharCode(65 + column)}${rowIndex + 1}`
        expect(worksheet, address).toContain(`r="${address}"`)
        expect(worksheet, String(value)).toContain(
          typeof value === 'number' ? `<v>${value}</v>` : `<t xml:space="preserve">${value}</t>`,
        )
      }
    }
  })

  it('repairs an understated dimension from the real used range when growing it', async () => {
    // Minimal writers ship refs like A1:A1 over a populated sheet. An edit
    // beyond the stated ref must not shrink-wrap the dimension to just the
    // mutation set — pre-existing cells past both would drop out of the
    // streamed viewport on reopen.
    const zip = await JSZip.loadAsync(await buildEditFixture())
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="3"><c r="C3"><v>9</v></c></row>
  </sheetData>
</worksheet>`,
    )
    const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    const mutation = await applyCellEditsToXlsx(buffer as Buffer, [
      { sheetName: 'Data', row: 1, column: 1, writeValue: true, cell: { value: 'x' } },
    ])
    const worksheet = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<dimension ref="A1:C3"/>')
  })

  it('bolds a styled cell via a new deduped xf without touching its content', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      styleEdit(0, 0, { bold: true }),
      styleEdit(0, 2, { bold: true }),
    ])
    expect(mutation.touchedEntries).toEqual([
      'xl/styles.xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
    ])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const worksheet = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')
    // Content untouched: A1 stays a shared-string cell, C1 keeps its value.
    // Both cells dedupe onto one new xf: bolding font 0 produces the existing
    // bold font 1, so A1 (already font 1) and C1 resolve identically.
    expect(worksheet).toContain('<c r="A1" t="s" s="2"><v>0</v></c>')
    expect(worksheet).toContain('<c r="C1" s="2"><v>5</v></c>')
    const styles = await entryText(mutation.buffer, 'xl/styles.xml')
    expect(styles).toContain('<fonts count="2">')
    expect(styles).toContain('<cellXfs count="3">')
  })

  it('applies fill, font color, and size as a solid fill plus new font', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      styleEdit(0, 2, { fillColor: '#FFF2CC', fontColor: '#C00000', fontSize: 14 }),
    ])
    const styles = await entryText(mutation.buffer, 'xl/styles.xml')
    expect(styles).toContain('<fgColor rgb="FFFFF2CC"/>')
    expect(styles).toContain('<color rgb="FFC00000"/>')
    expect(styles).toContain('<sz val="14"/>')
    expect(styles).toContain('applyFill="1"')
  })

  it('removes bold from an already-bold cell', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      styleEdit(0, 0, { bold: false }),
    ])
    const styles = await entryText(mutation.buffer, 'xl/styles.xml')
    const worksheet = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')
    const newXfIndex = /<c r="A1"[^>]*\bs="([0-9]+)"/.exec(worksheet)?.[1]
    expect(newXfIndex).toBe('2')
    const fonts = [...styles.matchAll(/<font\/>|<font>[\s\S]*?<\/font>/g)].map((m) => m[0])
    const cellXfsInner = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? ''
    const xfs = [...cellXfsInner.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)].map(
      (m) => m[0],
    )
    const newXf = xfs[Number(newXfIndex)]
    const fontId = Number(/fontId="([0-9]+)"/.exec(newXf ?? '')?.[1])
    expect(fonts[fontId]).not.toContain('<b')
  })

  it('maps builtin number formats and allocates custom ids from 164', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      styleEdit(0, 2, { numberFormat: '0.00' }),
      styleEdit(2, 1, { numberFormat: '$#,##0.00' }),
    ])
    const styles = await entryText(mutation.buffer, 'xl/styles.xml')
    expect(styles).toContain(
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="$#,##0.00"/></numFmts>',
    )
    expect(styles).toContain('numFmtId="2"')
    // numFmts must precede fonts per the schema sequence.
    expect(styles.indexOf('<numFmts')).toBeLessThan(styles.indexOf('<fonts'))
  })

  it('writes alignment and wrap into the new xf', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      styleEdit(0, 2, {
        horizontalAlignment: 'center',
        verticalAlignment: 'center',
        wrapText: true,
      }),
    ])
    const styles = await entryText(mutation.buffer, 'xl/styles.xml')
    expect(styles).toContain('<alignment horizontal="center" vertical="center" wrapText="1"/>')
    expect(styles).toContain('applyAlignment="1"')
  })

  it('styles a previously missing cell as an empty styled cell', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      styleEdit(9, 0, { fillColor: '#FFF2CC' }),
    ])
    const worksheet = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(worksheet).toMatch(/<row r="10"><c r="A10" s="[0-9]+"\/><\/row>/)
  })

  it('combines a value edit with a style delta in one cell', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 2,
        writeValue: true,
        cell: { value: 99 },
        style: { bold: true },
      },
    ])
    const worksheet = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(worksheet).toMatch(/<c r="C1" s="[0-9]+"><v>99<\/v><\/c>/)
  })
})

function styleEdit(row: number, column: number, style: NonNullable<CellEdit['style']>): CellEdit {
  return { sheetName: 'Data', row, column, writeValue: false, cell: { value: null }, style }
}

describe('assertOnlyTouchedEntriesChanged', () => {
  it('rejects a mutation that modified an untouched entry', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      edit(0, 0, { value: 'World' }),
    ])
    const tampered = {
      ...mutation,
      afterEntries: mutation.afterEntries.map((entry) =>
        entry.path === 'customXml/item1.xml' ? { ...entry, sha256: '0'.repeat(64) } : entry,
      ),
    }
    expect(() => assertOnlyTouchedEntriesChanged(tampered)).toThrow('unexpectedly modify')
  })

  it('rejects a mutation that dropped an entry', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      edit(0, 0, { value: 'World' }),
    ])
    const tampered = {
      ...mutation,
      afterEntries: mutation.afterEntries.filter((entry) => entry.path !== 'customXml/item1.xml'),
    }
    expect(() => assertOnlyTouchedEntriesChanged(tampered)).toThrow('package structure')
  })
})

function edit(
  row: number,
  column: number,
  cell: { value: string | number | boolean | null; formula?: string },
): CellEdit {
  return { sheetName: 'Data', row, column, writeValue: true, cell }
}

async function editedWorksheet(edits: readonly CellEdit[]): Promise<string> {
  const mutation = await applyCellEditsToXlsx(await buildEditFixture(), edits)
  assertOnlyTouchedEntriesChanged(mutation)
  return entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')
}

async function entryText(buffer: Buffer, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  if (!entry) throw new Error(`Missing ${path}`)
  return entry.async('text')
}

describe('rich-text run save', () => {
  it('writes inline-string runs with per-run properties', async () => {
    const worksheet = await editedWorksheet([
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: true,
        cell: { value: 'Hello World' },
        rich: [
          {
            text: 'Hello',
            bold: true,
            italic: false,
            underline: false,
            strikethrough: false,
            color: '#FF0000',
          },
          {
            text: ' World',
            bold: false,
            italic: true,
            underline: false,
            strikethrough: false,
            size: 14,
            family: 'Arial',
          },
        ],
      },
    ])
    expect(worksheet).toContain(
      '<c r="A1" s="1" t="inlineStr"><is>' +
        '<r><rPr><b/><color rgb="FFFF0000"/></rPr><t xml:space="preserve">Hello</t></r>' +
        '<r><rPr><i/><sz val="14"/><rFont val="Arial"/></rPr><t xml:space="preserve"> World</t></r>' +
        '</is></c>',
    )
  })
})

describe('style reset (Clear Formats / Clear All)', () => {
  it('resets a styled cell to the default xf without touching its content', async () => {
    const worksheet = await editedWorksheet([
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        styleReset: true,
      },
    ])
    expect(worksheet).toContain('<c r="A1" t="s" s="0"><v>0</v></c>')
  })

  it('derives a reset-then-format style from the default xf, not the old one', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        styleReset: true,
        style: { italic: true },
      },
    ])
    assertOnlyTouchedEntriesChanged(mutation)
    const worksheet = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')
    const styles = await entryText(mutation.buffer, 'xl/styles.xml')
    // A new xf derives from xf 0: italic only, without xf 1's bold font.
    expect(worksheet).toContain('<c r="A1" t="s" s="2">')
    expect(styles).toContain('<font><i/></font>')
  })

  it('Clear All empties the cell and resets its style', async () => {
    const worksheet = await editedWorksheet([
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: true,
        cell: { value: null },
        styleReset: true,
      },
    ])
    expect(worksheet).toContain('<c r="A1" s="0"/>')
  })
})

describe('large edit batches', () => {
  // Bulk actions (paste, sort, move-range) journal every affected cell, so a
  // single save can carry hundreds of thousands of edits. The save path must
  // stay a single pass over the worksheet — a per-edit whole-XML rewrite
  // would take hours at this scale.
  it('applies 100k inserts and 100k overwrites in one save each', async () => {
    const rowCount = 2_000
    const columnCount = 50
    const inserts: CellEdit[] = []
    for (let row = 0; row < rowCount; row += 1) {
      for (let column = 0; column < columnCount; column += 1) {
        inserts.push({
          sheetName: 'Sheet1',
          row,
          column,
          writeValue: true,
          cell: { value: row * columnCount + column },
        })
      }
    }
    const started = performance.now()
    const insertMutation = await applyCellEditsToXlsx(await blankXlsxBuffer(), inserts)
    const worksheet = await entryText(insertMutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<dimension ref="A1:AX2000"/>')
    expect(worksheet).toContain('<c r="A1"><v>0</v></c>')
    expect(worksheet).toContain(`<c r="AX2000"><v>${rowCount * columnCount - 1}</v></c>`)
    expect([...worksheet.matchAll(/<row /g)]).toHaveLength(rowCount)

    const overwrites: CellEdit[] = inserts.map((edit) => ({
      ...edit,
      cell: { value: `text ${edit.row}:${edit.column}` },
    }))
    const overwriteMutation = await applyCellEditsToXlsx(insertMutation.buffer, overwrites)
    const rewritten = await entryText(overwriteMutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(rewritten).toContain('<t xml:space="preserve">text 0:0</t>')
    expect(rewritten).toContain(
      `<t xml:space="preserve">text ${rowCount - 1}:${columnCount - 1}</t>`,
    )
    expect(rewritten).not.toContain('<v>0</v>')

    // Generous CI budget; the quadratic path this guards against took hours.
    expect(performance.now() - started).toBeLessThan(30_000)
  }, 60_000)
})

describe('text rotation and double underline save', () => {
  it('writes textRotation into the alignment and u val=double into the font', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { textRotation: 45, underline: true, underlineStyle: 'double' },
      },
    ])
    assertOnlyTouchedEntriesChanged(mutation)
    const styles = await entryText(mutation.buffer, 'xl/styles.xml')
    expect(styles).toContain('textRotation="45"')
    expect(styles).toContain('<u val="double"/>')
  })

  it('writes indent and protection flags into the xf', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { indent: 2, protectionLocked: false, protectionHidden: true },
      },
    ])
    assertOnlyTouchedEntriesChanged(mutation)
    const styles = await entryText(mutation.buffer, 'xl/styles.xml')
    expect(styles).toContain('indent="2"')
    expect(styles).toContain('applyProtection="1"')
    expect(styles).toContain('<protection locked="0" hidden="1"/>')
  })

  it('omits protection at the OOXML defaults', async () => {
    const mutation = await applyCellEditsToXlsx(await buildEditFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { protectionLocked: true, protectionHidden: false },
      },
    ])
    const styles = await entryText(mutation.buffer, 'xl/styles.xml')
    expect(styles).not.toContain('<protection')
  })

  it('clears a rotation with textRotation 0', async () => {
    const first = await applyCellEditsToXlsx(await buildEditFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { textRotation: 135 },
      },
    ])
    const styles = await entryText(first.buffer, 'xl/styles.xml')
    expect(styles).toContain('textRotation="135"')
    const cleared = await applyCellEditsToXlsx(await buildEditFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { textRotation: 0, bold: true },
      },
    ])
    const clearedStyles = await entryText(cleared.buffer, 'xl/styles.xml')
    expect(clearedStyles).not.toContain('textRotation')
  })
})
