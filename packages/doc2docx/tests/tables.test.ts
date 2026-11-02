import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  LegacyDocError,
  convertLegacyDocToDocx,
  generateTableXml,
  readLegacyDocBody,
} from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodePapxFkp,
  encodeSprm,
  encodeTapSprm,
  encodeTableCellPapx,
  encodeTableRowPapx,
  type SyntheticPapxEntry,
  type SyntheticTableRow,
} from './helpers/synthetic-doc'

function fixtureBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(fileURLToPath(new URL('./fixtures/' + name, import.meta.url))))
}

async function documentXml(docx: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(docx)
  return zip.file('word/document.xml')!.async('string')
}

function errorCode(run: () => unknown): string {
  try {
    run()
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyDocError)
    return (error as LegacyDocError).code
  }
  throw new Error('expected a LegacyDocError')
}

/**
 * Build a synthetic .doc whose body text holds the given table rows: each
 * cell ends with a cell mark and each row with the extra row-end mark, and the
 * corresponding PAPX entries carry the cell/row TAP sprms.
 */
function syntheticTableDoc(options: {
  rows: SyntheticTableRow[]
  afterText?: string
  itap?: number
  rowPapxOverride?: Array<Uint8Array | undefined>
}): Uint8Array {
  const itap = options.itap ?? 1
  let text = ''
  const entries: SyntheticPapxEntry[] = []
  options.rows.forEach((row, rowIndex) => {
    const cellCount = row.gridWidths.length
    for (let cellIndex = 0; cellIndex < cellCount; cellIndex++) {
      const startCp = text.length
      const cellText = row.texts?.[cellIndex] ?? 'C' + (rowIndex + 1) + (cellIndex + 1)
      text += cellText + '\u0007'
      entries.push({
        fcStart: SYNTHETIC_TEXT_OFFSET + startCp * 2,
        fcEnd: 0,
        istd: 0,
        grpprl: encodeTableCellPapx(itap),
      })
    }
    const rowEndCp = text.length
    text += '\u0007'
    entries.push({
      fcStart: SYNTHETIC_TEXT_OFFSET + rowEndCp * 2,
      fcEnd: 0,
      istd: 0,
      grpprl: options.rowPapxOverride?.[rowIndex] ?? encodeTableRowPapx(row, itap),
    })
  })
  text += options.afterText ?? 'After table\r'
  entries.sort((left, right) => left.fcStart - right.fcStart)
  entries.forEach((entry, index) => {
    entry.fcEnd =
      index + 1 < entries.length
        ? entries[index + 1]!.fcStart
        : SYNTHETIC_TEXT_OFFSET + text.length * 2
  })
  return buildSyntheticDoc({ pieces: [{ text }], papxFkpPages: [encodePapxFkp(entries)] })
}

describe('table.doc: reconstruction and writer output', () => {
  const body = readLegacyDocBody(fixtureBytes('table.doc'))

  it('reconstructs the 2x3 table with equal 2880-twip columns', () => {
    expect(body.blocks.map((block) => block.type)).toEqual(['table', 'paragraph', 'paragraph'])
    const table = body.blocks[0]!
    if (table.type !== 'table') throw new Error('expected a table block')
    expect(table.table.gridWidths).toEqual([2880, 2880, 2880])
    expect(table.table.rows).toHaveLength(2)
    expect(table.table.rows.map((row) => row.cells.length)).toEqual([3, 3])
    expect(table.table.rows[0]!.cells.map((cell) => cell.paragraphs[0]!.text)).toEqual([
      'C11',
      'C12',
      'C13',
    ])
    expect(table.table.rows[1]!.cells.map((cell) => cell.paragraphs[0]!.text)).toEqual([
      'C21',
      'C22',
      'C23',
    ])
    expect(table.table.rows[0]!.cells.every((cell) => cell.borders?.top === true)).toBe(true)
    expect(table.table.borders.top).toBe(true)
    expect(table.table.rows[0]!.header).toBe(false)
    expect(table.table.cellPadding).toEqual({ left: 108, right: 108 })
    expect(body.warnings).toContain(
      'table cell paragraph formatting is flattened to plain paragraphs',
    )
  })

  it('writes a real w:tbl at the table position, then the paragraph after it', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(fixtureBytes('table.doc'))
    const xml = await documentXml(docx)
    expect(xml).toContain('<w:tbl>')
    expect(xml).toContain(
      '<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>',
    )
    expect(xml).toContain('<w:tblBorders>')
    expect(xml.match(/<w:tr>/g)).toHaveLength(2)
    expect(xml.match(/<w:tc>/g)).toHaveLength(6)
    const cellTexts = [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((match) => match[1])
    expect(cellTexts.slice(0, 6)).toEqual(['C11', 'C12', 'C13', 'C21', 'C22', 'C23'])
    expect(xml.indexOf('After table')).toBeGreaterThan(xml.indexOf('</w:tbl>'))
    expect(warnings).toContain('table cell paragraph formatting is flattened to plain paragraphs')
    expect(warnings.some((warning) => warning.includes('tables are flattened'))).toBe(false)
    const parsed = await parseDocx(docx)
    expect(parsed.blocks.some((block) => block.type === 'table')).toBe(true)
    expect(JSON.stringify(parsed.blocks)).toContain('C11')
  })
})

describe('rich.doc: merged first row plus kept honest warnings', () => {
  const body = readLegacyDocBody(fixtureBytes('rich.doc'))

  it('reconstructs a 3x3 table whose first row merges the first two columns', () => {
    const tableBlock = body.blocks.find((block) => block.type === 'table')
    if (!tableBlock || tableBlock.type !== 'table') throw new Error('expected a table block')
    const table = tableBlock.table
    expect(table.gridWidths).toEqual([2976, 2976, 2976])
    expect(table.alignment).toBe('center')
    expect(table.rows).toHaveLength(3)
    expect(table.rows[0]!.cells.map((cell) => cell.gridSpan)).toEqual([2, 1])
    expect(table.rows[0]!.cells.map((cell) => cell.widthTwips)).toEqual([5952, 2976])
    expect(table.rows[0]!.cells[0]!.paragraphs[0]!.text).toBe('Merged header')
    expect(table.rows[0]!.cells[1]!.paragraphs[0]!.text).toBe('R1C3')
    expect(table.rows[1]!.cells.map((cell) => cell.paragraphs[0]!.text)).toEqual([
      'R2C1',
      'R2C2',
      'R2C3',
    ])
  })

  it('writes gridSpan 2, keeps the text order, and reports the remaining gaps', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(fixtureBytes('rich.doc'))
    const xml = await documentXml(docx)
    expect(xml).toContain('<w:gridSpan w:val="2"/>')
    expect(xml.indexOf('Merged header')).toBeLessThan(xml.indexOf('R1C3'))
    expect(xml.indexOf('R1C3')).toBeLessThan(xml.indexOf('R2C1'))
    expect(xml.indexOf('R3C3')).toBeLessThan(xml.indexOf('Image below:'))
    expect(warnings.some((warning) => warning.includes('images are not converted'))).toBe(true)
    expect(warnings.some((warning) => warning.includes('hyperlinks are not converted'))).toBe(true)
    expect(warnings.some((warning) => warning.includes('lists are not converted'))).toBe(true)
    expect(warnings.some((warning) => warning.includes('tables are flattened'))).toBe(false)
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })
})

describe('synthetic tables', () => {
  it('reads widths, alignment and borders from the TAP', () => {
    const bytes = syntheticTableDoc({
      rows: [
        {
          gridWidths: [1200, 2400],
          cells: [{}, {}],
          texts: ['A', 'B'],
          alignment: 'center',
          tableBorders: true,
          cantSplit: true,
        },
      ],
    })
    const body = readLegacyDocBody(bytes)
    const tableBlock = body.blocks[0]!
    if (tableBlock.type !== 'table') throw new Error('expected a table block')
    const table = tableBlock.table
    expect(table.gridWidths).toEqual([1200, 2400])
    expect(table.alignment).toBe('center')
    expect(table.borders.top).toBe(true)
    expect(table.rows[0]!.cantSplit).toBe(true)
    expect(table.rows[0]!.cells.map((cell) => cell.widthTwips)).toEqual([1200, 2400])
    expect(table.rows[0]!.cells.map((cell) => cell.paragraphs[0]!.text)).toEqual(['A', 'B'])
  })

  it('emits header rows, cantSplit and row heights in the writer', async () => {
    const bytes = syntheticTableDoc({
      rows: [
        {
          gridWidths: [1000, 1000],
          cells: [{}, {}],
          texts: ['H1', 'H2'],
          header: true,
          cantSplit: true,
          rowHeightTwips: 720,
          tableBorders: true,
        },
      ],
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = await documentXml(docx)
    expect(xml).toContain('<w:tblHeader/>')
    expect(xml).toContain('<w:cantSplit/>')
    expect(xml).toContain('<w:trHeight w:val="720" w:hRule="atLeast"/>')
    expect(xml).toContain('<w:tblBorders>')
  })

  it('folds fMerged continuation cells into the first cell gridSpan', () => {
    const bytes = syntheticTableDoc({
      rows: [
        {
          gridWidths: [1000, 2000, 3000],
          cells: [{ rgf: 0x0001 }, { rgf: 0x0002 }, {}],
          texts: ['Merged', '', 'C'],
        },
      ],
    })
    const body = readLegacyDocBody(bytes)
    const tableBlock = body.blocks[0]!
    if (tableBlock.type !== 'table') throw new Error('expected a table block')
    expect(tableBlock.table.rows[0]!.cells.map((cell) => cell.gridSpan)).toEqual([2, 1])
    expect(tableBlock.table.rows[0]!.cells.map((cell) => cell.widthTwips)).toEqual([3000, 3000])
    expect(tableBlock.table.rows[0]!.cells[0]!.paragraphs[0]!.text).toBe('Merged')
  })

  it('writes vertical merges as vMerge restart/continue', async () => {
    const bytes = syntheticTableDoc({
      rows: [
        { gridWidths: [2000], cells: [{ rgf: 0x0040 }], texts: ['V'] },
        { gridWidths: [2000], cells: [{ rgf: 0x0020 }], texts: [''] },
      ],
    })
    const body = readLegacyDocBody(bytes)
    const tableBlock = body.blocks[0]!
    if (tableBlock.type !== 'table') throw new Error('expected a table block')
    expect(tableBlock.table.rows[0]!.cells[0]!.verticalMerge).toBe('restart')
    expect(tableBlock.table.rows[1]!.cells[0]!.verticalMerge).toBe('continue')
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = await documentXml(docx)
    expect(xml).toContain('<w:vMerge w:val="restart"/>')
    expect(xml).toContain('<w:vMerge/>')
  })

  it('omits borders when no border is visible', () => {
    const bytes = syntheticTableDoc({
      rows: [
        {
          gridWidths: [1000, 1000],
          cells: [
            { borders: { top: false, left: false, bottom: false, right: false } },
            { borders: { top: false, left: false, bottom: false, right: false } },
          ],
          texts: ['A', 'B'],
        },
      ],
    })
    const body = readLegacyDocBody(bytes)
    const tableBlock = body.blocks[0]!
    if (tableBlock.type !== 'table') throw new Error('expected a table block')
    expect(tableBlock.table.borders).toEqual({
      top: false,
      left: false,
      bottom: false,
      right: false,
      insideHorizontal: false,
      insideVertical: false,
    })
    expect(tableBlock.table.rows[0]!.cells[0]!.borders).toBeUndefined()
    expect(generateTableXml(tableBlock.table)).not.toContain('<w:tblBorders>')
  })

  it('flattens cell borders into tcBorders but keeps the table border set', async () => {
    const bytes = syntheticTableDoc({
      rows: [
        {
          gridWidths: [1000],
          cells: [{ borders: { top: true, left: false, bottom: false, right: true } }],
          texts: ['A'],
        },
      ],
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = await documentXml(docx)
    expect(xml).toContain('<w:tcBorders><w:top')
    expect(xml).toContain('<w:right')
    expect(xml).not.toContain(
      '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders><w:p>',
    )
  })

  it('keeps multiple paragraphs inside one cell', async () => {
    const bytes = syntheticTableDoc({
      rows: [{ gridWidths: [2000], cells: [{}], texts: ['A\rB'] }],
    })
    const body = readLegacyDocBody(bytes)
    const tableBlock = body.blocks[0]!
    if (tableBlock.type !== 'table') throw new Error('expected a table block')
    expect(
      tableBlock.table.rows[0]!.cells[0]!.paragraphs.map((paragraph) => paragraph.text),
    ).toEqual(['A', 'B'])
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = await documentXml(docx)
    const tableXml = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'))
    expect(tableXml.match(/<w:p>/g) ?? []).toHaveLength(2)
  })

  it('warns once about nested tables (sprmPItap > 1)', () => {
    const bytes = syntheticTableDoc({
      itap: 2,
      rows: [{ gridWidths: [1000], cells: [{}], texts: ['A'] }],
    })
    const body = readLegacyDocBody(bytes)
    expect(
      body.warnings.filter((warning) => warning.includes('nested tables are not supported')),
    ).toHaveLength(1)
  })

  it('rejects a truncated TDefTable', () => {
    const truncated = encodeTapSprm(0x08, [1, 0, 0])
    const sprms = new Uint8Array([
      ...encodeSprm(0x2416, [1]),
      ...encodeSprm(0x2417, [1]),
      ...encodeSprm(0x6649, [1, 0, 0, 0]),
      ...truncated,
    ])
    const bytes = syntheticTableDoc({
      rows: [{ gridWidths: [1000], cells: [{}], texts: ['A'] }],
      rowPapxOverride: [sprms],
    })
    expect(errorCode(() => readLegacyDocBody(bytes))).toBe('bad_table')
  })

  it('rejects a row with no TDefTable', () => {
    const sprms = new Uint8Array([
      ...encodeSprm(0x2416, [1]),
      ...encodeSprm(0x2417, [1]),
      ...encodeSprm(0x6649, [1, 0, 0, 0]),
    ])
    const bytes = syntheticTableDoc({
      rows: [{ gridWidths: [1000], cells: [{}], texts: ['A'] }],
      rowPapxOverride: [sprms],
    })
    expect(errorCode(() => readLegacyDocBody(bytes))).toBe('bad_table')
  })

  it('rejects a row whose cell count does not match its TDefTable', () => {
    const sprms = encodeTableRowPapx({ gridWidths: [1000, 1000], cells: [{}, {}] })
    const bytes = syntheticTableDoc({
      rows: [{ gridWidths: [1000], cells: [{}], texts: ['A'] }],
      rowPapxOverride: [sprms],
    })
    expect(errorCode(() => readLegacyDocBody(bytes))).toBe('bad_table')
  })

  it('rejects unordered cell boundaries', () => {
    const operand = new Uint8Array(1 + 2 * 2 + 20)
    operand[0] = 1 // itcMac
    new DataView(operand.buffer).setInt16(1, 500, true)
    new DataView(operand.buffer).setInt16(3, 100, true)
    const sprms = new Uint8Array([
      ...encodeSprm(0x2416, [1]),
      ...encodeSprm(0x2417, [1]),
      ...encodeSprm(0x6649, [1, 0, 0, 0]),
      ...encodeTapSprm(0x08, operand),
    ])
    const bytes = syntheticTableDoc({
      rows: [{ gridWidths: [1000], cells: [{}], texts: ['A'] }],
      rowPapxOverride: [sprms],
    })
    expect(errorCode(() => readLegacyDocBody(bytes))).toBe('bad_table')
  })
})
