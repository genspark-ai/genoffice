import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { StylesheetEditor } from '../src/gateway/xlsx-styles'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { workbookRangeResultSchema } from '../src/shared/desktop-api'
import { fromNeutralStyle, toNeutralStyle } from '../src/renderer/edit-journal'

const STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="2"><border/><border diagonalUp="1"><left style="thin"/><diagonal style="hair"><color rgb="FF888888"/></diagonal></border></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="2"><xf/><xf borderId="1" applyBorder="1" fillId="1"/></cellXfs>
</styleSheet>`

describe('StylesheetEditor borders', () => {
  it('derives a new border with edges in schema order and dedupes', () => {
    const editor = new StylesheetEditor(STYLES)
    const first = editor.resolveStyle(0, {
      borderBottom: { style: 'medium', color: '#FF0000' },
      borderTop: { style: 'thin' },
    })
    const again = editor.resolveStyle(0, {
      borderBottom: { style: 'medium', color: '#FF0000' },
      borderTop: { style: 'thin' },
    })
    expect(again).toBe(first)
    const xml = editor.serialize()
    expect(xml).toContain(
      '<border><top style="thin"/><bottom style="medium"><color rgb="FFFF0000"/></bottom></border>',
    )
    expect(xml).toContain('<borders count="3">')
    expect(xml).toContain('borderId="2" applyBorder="1"')
  })

  it('keeps untouched edges, diagonals, and border attributes', () => {
    const editor = new StylesheetEditor(STYLES)
    editor.resolveStyle(1, { borderRight: { style: 'double' } })
    const xml = editor.serialize()
    expect(xml).toContain(
      '<border diagonalUp="1"><left style="thin"/><right style="double"/>' +
        '<diagonal style="hair"><color rgb="FF888888"/></diagonal></border>',
    )
  })

  it('removes an edge with null and clears fills back to none', () => {
    const editor = new StylesheetEditor(STYLES)
    editor.resolveStyle(1, { borderLeft: null, fillColor: null })
    const xml = editor.serialize()
    expect(xml).toContain('<border diagonalUp="1"><left/><diagonal style="hair">')
    expect(xml).toContain('fillId="0"')
  })
})

describe('journal border mapping', () => {
  it('maps Univer bd deltas to neutral border edges and back', () => {
    const neutral = toNeutralStyle({
      bd: {
        t: { s: 1, cl: { rgb: '#FF0000' } },
        b: null,
        tl_br: null,
      },
    })
    expect(neutral).toEqual({
      borderTop: { style: 'thin', color: '#FF0000' },
      borderBottom: null,
    })
    expect(fromNeutralStyle(neutral ?? {})).toEqual({
      bd: {
        t: { s: 1, cl: { rgb: '#FF0000' } },
        b: null,
      },
    })
  })

  it('maps fill clearing through bg null', () => {
    const neutral = toNeutralStyle({ bg: null })
    expect(neutral).toEqual({ fillColor: null })
    // The overlay replays a clear with the same empty-rgb sentinel the
    // installer uses: bg: null would be stripped by the mutation's
    // removeNull and let a <col style=> fill compose back through.
    expect(fromNeutralStyle(neutral ?? {})).toEqual({ bg: { rgb: '' } })
    expect(toNeutralStyle(fromNeutralStyle(neutral ?? {}))).toEqual({ fillColor: null })
  })
})

// The sidecar used to drop value-less styled cells, losing borders
// on blank cells and around merged ranges.
describe('sidecar read side keeps styled blanks', () => {
  it('returns value-less bordered/filled cells with their style index', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    )
    zip.file(
      'xl/styles.xml',
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font/></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thick"/></border></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf borderId="1" applyBorder="1"/><xf fillId="2" applyFill="1"/></cellXfs>
</styleSheet>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"/><c r="C1" s="1"></c><c r="D1" s="0"/></row>
<row r="2"><c r="A2" s="2"/><c r="B2" s="2"/></row>
</sheetData>
<mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>
</worksheet>`,
    )

    const directory = await mkdtemp(join(tmpdir(), 'xlsx-borders-read-'))
    const path = join(directory, 'styled-blanks.xlsx')
    await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    let sessionId: string | null = null
    try {
      const opened = (await client.open(path)) as {
        sessionId: string
        styles: { borderBottom?: { style: string }; fillColor?: string }[]
      }
      sessionId = opened.sessionId
      // Merges only surface once the worksheet part is fully indexed; an
      // immediate read races the streaming indexer on slow runners.
      let result
      for (;;) {
        result = workbookRangeResultSchema.parse(
          await client.readRange({
            sessionId,
            sheetId: 'sheet-1',
            range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 3 },
          }),
        )
        if (result.indexingComplete) break
        await new Promise((resolve) => setTimeout(resolve, 5))
      }
      const shape = result.cells
        .map((cell) => ({ row: cell.row, column: cell.column, styleIndex: cell.styleIndex }))
        .sort((a, b) => a.row - b.row || a.column - b.column)
      expect(shape).toEqual([
        // No s= is cellXfs[0], not "no style".
        { row: 0, column: 0, styleIndex: 0 },
        { row: 0, column: 1, styleIndex: 1 },
        { row: 0, column: 2, styleIndex: 1 },
        { row: 1, column: 0, styleIndex: 2 },
        { row: 1, column: 1, styleIndex: 2 },
      ])
      expect(opened.styles[1]?.borderBottom).toEqual({ style: 'thick' })
      expect(opened.styles[2]?.fillColor).toBe('#FFCC00')
      expect(result.merges).toEqual([{ startRow: 1, startColumn: 0, endRow: 1, endColumn: 1 }])
    } finally {
      if (sessionId) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}
