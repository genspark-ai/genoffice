import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'
import { part } from './xlsx-helpers'

const sidecar = Boolean(xlsxSidecarPath())

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

/**
 * An openpyxl-shaped workbook: the worksheet's comments and VML drawing
 * relationships use absolute OPC targets (`/xl/comments/comment1.xml`), which
 * are valid (ECMA-376 Part 2: a leading '/' means the package root) but broke
 * the IronCalc import behind `sheet read` (#2306). The formula's cached value
 * is stripped so the read has to evaluate it through the engine.
 */
async function absoluteRelationshipsWorkbook(): Promise<string> {
  const dir = tempDir()
  const table = join(dir, 'table.json')
  writeFileSync(table, JSON.stringify([[10], [20], ['=SUM(A1:A2)']]))
  const path = join(dir, 'absolute-rels.xlsx')
  expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', path])).code).toBe(0)
  const zip = await JSZip.loadAsync(readFileSync(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file('xl/worksheets/sheet1.xml', sheet.replace(/(<f>[^<]*<\/f>)<v>[^<]*<\/v>/g, '$1'))
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${RELS_NS}">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="/xl/comments/comment1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="/xl/drawings/commentsDrawing1.vml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/comments/comment1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<authors><author>openpyxl</author></authors>' +
      '<commentList><comment ref="A1" authorId="0"><text><t>note</t></text></comment></commentList>' +
      '</comments>',
  )
  zip.file(
    'xl/drawings/commentsDrawing1.vml',
    '<xml xmlns:v="urn:schemas-microsoft-com:vml"></xml>',
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
  return path
}

describe('sheet read with absolute OPC relationship targets (#2306)', () => {
  it.skipIf(!sidecar)('reads cells and evaluates the uncached formula', async () => {
    const path = await absoluteRelationshipsWorkbook()
    // the fixture really does carry the absolute targets openpyxl writes
    expect(await part(path, 'xl/worksheets/_rels/sheet1.xml.rels')).toContain(
      'Target="/xl/comments/comment1.xml"',
    )
    const r = await run(['sheet', 'read', path, '--json'])
    expect(r.code).toBe(0)
    const detail = r.json().detail
    expect(detail.rows).toEqual([[10], [20], [30]])
    expect(detail.formulas).toEqual({ A3: '=SUM(A1:A2)' })
    // the comment behind the absolute target is still read (not lost to the fixup)
    expect(detail.features.notes).toBe(1)
  })
})
