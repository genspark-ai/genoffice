import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compareDocuments, parseDocx } from '../src/index'
import { applyTrackedChangesXml } from '../src/revision-xml'
import { buildDocx } from './helpers/build-docx'
import {
  SHAPES,
  baseGrid,
  gridShape,
  gridXml,
  withColumn,
  withRow,
  withoutColumn,
  withoutRow,
  xmlShape,
  type Grid,
} from './helpers/merged-tables'

const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const around = (grid: Grid) =>
  para('This agreement is made between the supplier and the customer.') +
  gridXml(grid) +
  para('Signed by both parties.')

async function compare(original: Grid, revised: Grid) {
  const result = await compareDocuments(
    await parseDocx(await buildDocx({ bodyXml: around(original) })),
    await parseDocx(await buildDocx({ bodyXml: around(revised) })),
    { author: 'Rae Visor', date: '2026-11-24T00:00:00Z' },
  )
  // the comparison as the file that is written and read back
  const reopened = await parseDocx(result.bytes)
  const xml =
    (await (await JSZip.loadAsync(result.bytes)).file('word/document.xml')?.async('string')) ?? ''
  expect(reopened.internal.documentXml).toBe(xml)
  const resolved = (mode: 'accept' | 'reject') => {
    const out = applyTrackedChangesXml(xml, mode)
    expect(out.remaining).toBe(0)
    return xmlShape(out.xml)
  }
  return { xml, summary: result.summary, resolved }
}

const WHOLE = /<w:(?:del|ins) [^>]*><w:tbl>/

describe('Compare: a column added to or taken out of a table with merged cells', () => {
  for (const [name, merges] of Object.entries(SHAPES)) {
    const original = baseGrid(merges)
    for (let at = 0; at <= 4; at++) {
      const { grid: revised, reaches } = withColumn(original, at)
      it(`${name}; a column comes in at ${at}${reaches ? ' (inside the merged cell)' : ''}`, async () => {
        const { xml, summary, resolved } = await compare(original, revised)
        expect(resolved('accept')).toEqual(gridShape(revised))
        expect(resolved('reject')).toEqual(gridShape(original))
        // a merged cell that reaches across the column: not compared by column
        expect(/<w:cell(?:Ins|Del) /.test(xml)).toBe(!reaches)
        if (reaches) return
        expect(WHOLE.test(xml)).toBe(false)
        expect(xml.match(/<w:tbl>/g)).toHaveLength(1)
        // one cell of the new column per row it starts in
        const starts = gridShape(revised).filter((row) =>
          row.some(([text]) => text.startsWith('Added')),
        )
        expect(xml.match(/<w:cellIns /g)).toHaveLength(4)
        expect(starts).toHaveLength(4)
        expect(xml).not.toContain('<w:cellDel ')
        expect(summary).toMatchObject({ insertions: 1, deletions: 0 })
        expect(xml.match(/<w:gridCol /g)).toHaveLength(5)
      })
    }
    for (let at = 0; at < 4; at++) {
      const { grid: revised, reaches } = withoutColumn(original, at)
      it(`${name}; column ${at} goes${reaches ? ' (out of the merged cell)' : ''}`, async () => {
        const { xml, summary, resolved } = await compare(original, revised)
        expect(resolved('accept')).toEqual(gridShape(revised))
        expect(resolved('reject')).toEqual(gridShape(original))
        // a merged cell that loses a column: replaced, the table whole or the cell's columns
        if (reaches) return
        expect(WHOLE.test(xml)).toBe(false)
        expect(xml.match(/<w:tbl>/g)).toHaveLength(1)
        expect(xml).not.toContain('<w:cellIns ')
        // the column, and a row whose text no longer pairs it with its twin
        expect(summary.deletions).toBeGreaterThanOrEqual(1)
        expect(xml.match(/<w:gridCol /g)).toHaveLength(4)
      })
    }
  }

  it('the reverse comparison of every pair resolves the other way round', async () => {
    for (const merges of Object.values(SHAPES)) {
      const original = baseGrid(merges)
      for (let at = 0; at <= 4; at++) {
        const { grid: wider } = withColumn(original, at)
        const { resolved } = await compare(wider, original)
        expect(resolved('accept')).toEqual(gridShape(original))
        expect(resolved('reject')).toEqual(gridShape(wider))
      }
    }
  })

  it('a row that goes while a column comes in beside a cell merged across: the row is given a cell for the column', async () => {
    const original = baseGrid([{ r0: 0, r1: 0, c0: 0, c1: 1 }])
    const { grid: wider } = withColumn(original, 2)
    const revised: Grid = { texts: wider.texts.filter((_, r) => r !== 2), merges: wider.merges }
    const { xml, resolved } = await compare(original, revised)
    expect(WHOLE.test(xml)).toBe(false)
    expect(xml.match(/<w:trPr><w:del /g)).toHaveLength(1)
    // three cells of the new column and the one the struck row is given
    expect(xml.match(/<w:cellIns /g)).toHaveLength(4)
    expect(xml).toMatch(/<w:tcPr><w:tcW [^>]*\/><w:cellIns [^>]*\/><\/w:tcPr><w:p\/><\/w:tc>/)
    expect(resolved('accept')).toEqual(gridShape(revised))
    expect(resolved('reject')).toEqual(gridShape(original))
  })

  it('a row that goes in a table with a cell merged down: the column and the row are tracked, the merge kept', async () => {
    const original = baseGrid([{ r0: 0, r1: 1, c0: 3, c1: 3 }])
    const { grid: narrower } = withoutColumn(original, 1)
    const revised: Grid = {
      texts: narrower.texts.filter((_, r) => r !== 3),
      merges: narrower.merges,
    }
    const { xml, resolved } = await compare(original, revised)
    expect(WHOLE.test(xml)).toBe(false)
    expect(xml.match(/<w:trPr><w:del /g)).toHaveLength(1)
    // the struck row's own cell of the column is marked like the others
    expect(xml.match(/<w:cellDel /g)).toHaveLength(4)
    expect(resolved('accept')).toEqual(gridShape(revised))
    expect(resolved('reject')).toEqual(gridShape(original))
  })

  it('a row that reads the same in both tables still gets its cells for a column that goes and one that comes in', async () => {
    // two columns under one merged cell go, two come in at the end; the second row is the same XML on both sides
    const same = ['n/a', 'n/a', 'n/a', 'n/a', 'n/a']
    const original: Grid = {
      texts: [
        ['Item', 'Stock', '', 'Price', 'Origin'],
        same,
        ['Bolt', 'many', 'on hand', 'ten dollars', 'Leeds'],
        ['Nut', 'few', 'on order', 'two dollars', 'Turin'],
      ],
      merges: [{ r0: 0, r1: 0, c0: 1, c1: 2 }],
    }
    const revised: Grid = {
      texts: [
        ['Item', 'Price', 'Origin', 'Maker', 'Since'],
        same,
        ['Bolt', 'ten dollars', 'Leeds', 'Acme', 'last year'],
        ['Nut', 'two dollars', 'Turin', 'Borg', 'this year'],
      ],
      merges: [],
    }
    const { xml, resolved } = await compare(original, revised)
    expect(xml.match(/<w:tbl>/g)).toHaveLength(1)
    expect(xml).toContain('<w:cellDel ')
    expect(xml).toContain('<w:cellIns ')
    // every row covers the whole grid
    const grid = xml.match(/<w:gridCol /g)!.length
    const widths = xmlShape(xml).map((row) => row.reduce((sum, [, span]) => sum + span, 0))
    expect(grid).toBe(7)
    expect(widths.length).toBeGreaterThanOrEqual(4)
    expect(new Set(widths)).toEqual(new Set([7]))
    expect(resolved('accept')).toEqual(gridShape(revised))
    expect(resolved('reject')).toEqual(gridShape(original))
  })

  it('cells merged differently in a column both tables have leave the table replaced whole', async () => {
    const original = baseGrid([{ r0: 1, r1: 2, c0: 0, c1: 0 }])
    const { grid: wider } = withColumn(baseGrid([]), 4)
    const { xml, resolved } = await compare(original, wider)
    expect(WHOLE.test(xml)).toBe(true)
    expect(resolved('accept')).toEqual(gridShape(wider))
    expect(resolved('reject')).toEqual(gridShape(original))
  })
})

describe('Compare: rows added or removed among cells merged down', () => {
  const variants = (original: Grid) => {
    const columns: Array<[string, Grid]> = [['the same columns', original]]
    for (let at = 0; at <= 4; at++) {
      const { grid, reaches } = withColumn(original, at)
      if (!reaches) columns.push([`a column in at ${at}`, grid])
    }
    for (let at = 0; at < 4; at++) {
      const { grid, reaches } = withoutColumn(original, at)
      if (!reaches) columns.push([`column ${at} out`, grid])
    }
    const out: Array<[string, Grid]> = []
    for (const [name, grid] of columns) {
      for (let at = 0; at <= 4; at++) out.push([`${name}, a row in at ${at}`, withRow(grid, at)])
      for (let at = 0; at < 4; at++) out.push([`${name}, row ${at} out`, withoutRow(grid, at)])
    }
    return out
  }

  // whole: paired rows merged differently (a row that started or ended a merge went), or tables too far apart to be paired
  const WHOLES: Record<string, number> = {
    'a heading cell across the first two columns': 0,
    'a cell across the last two columns of a middle row': 0,
    'a cell down two rows of the first column': 20,
    'a cell down the whole last column': 30,
    'a block two columns wide and two rows high': 65,
    'a cell across and another down': 28,
  }

  for (const [name, merges] of Object.entries(SHAPES)) {
    it(`${name}: accept all gives the revised table and reject all the original, either way round`, async () => {
      const original = baseGrid(merges)
      let whole = 0
      for (const [change, revised] of variants(original)) {
        for (const [from, to] of [
          [original, revised],
          [revised, original],
        ] as const) {
          const { xml, resolved } = await compare(from, to)
          expect(resolved('accept'), change).toEqual(gridShape(to))
          expect(resolved('reject'), change).toEqual(gridShape(from))
          if (WHOLE.test(xml)) whole++
          else expect(xml.match(/<w:tbl>/g), change).toHaveLength(1)
        }
      }
      expect(whole).toBe(WHOLES[name])
    }, 60_000)
  }

  it('a row that comes in where a column with a cell merged down goes is given a cell that continues the merge', async () => {
    const original = baseGrid([{ r0: 0, r1: 3, c0: 3, c1: 3 }])
    const revised = withRow(withoutColumn(original, 3).grid, 2)
    const { xml, resolved } = await compare(original, revised)
    expect(WHOLE.test(xml)).toBe(false)
    expect(xml.match(/<w:trPr><w:ins /g)).toHaveLength(1)
    expect(xml).toMatch(
      /<w:tcPr><w:tcW [^>]*\/><w:vMerge\/><w:cellDel [^>]*\/><\/w:tcPr><w:p\/><\/w:tc>/,
    )
    // while the changes are pending the merge runs through the new row: five rows, one cell merged down all of them
    const last = xmlShape(xml).map((row) => row.at(-1)![2])
    expect(last).toEqual(['restart', 'continue', 'continue', 'continue', 'continue'])
    expect(resolved('accept')).toEqual(gridShape(revised))
    expect(resolved('reject')).toEqual(gridShape(original))
  })

  it('a row of one table that starts a merge the rows below continue leaves the table replaced whole', async () => {
    const original = baseGrid([{ r0: 1, r1: 3, c0: 0, c1: 0 }])
    const { grid: wider } = withColumn(original, 4)
    const revised: Grid = {
      merges: wider.merges,
      texts: wider.texts.map((row, r) =>
        r === 1 ? ['Rivet', 'a lot', 'nine dollars', 'Bonn', 'Added tin'] : row,
      ),
    }
    const { xml, resolved } = await compare(original, revised)
    expect(WHOLE.test(xml)).toBe(true)
    expect(resolved('accept')).toEqual(gridShape(revised))
    expect(resolved('reject')).toEqual(gridShape(original))
  })
})
