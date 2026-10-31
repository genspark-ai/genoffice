import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import {
  createBufferEntrySource,
  planCellEditsToXlsx,
  type CellEdit,
  type ExternalLinkOps,
  type MutationPlan,
} from '../src/gateway/xlsx-gateway'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const STYLES =
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

/// The dependent workbook of the CLI suite: link [1] points at a missing
/// `gone.xlsx`, link [2] at the relative `source.xlsx`, and the sheet holds
/// two formulas on [2] and one on [1].
async function dependentBuffer(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    XML_DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/><Override PartName="/xl/externalLinks/externalLink2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    XML_DECL +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId4"/><externalReference r:id="rId5"/></externalReferences></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink2.xml"/></Relationships>',
  )
  zip.file('xl/styles.xml', XML_DECL + STYLES)
  for (const n of [1, 2]) {
    zip.file(
      `xl/externalLinks/externalLink${n}.xml`,
      XML_DECL +
        `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>1</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>`,
    )
    zip.file(
      `xl/externalLinks/_rels/externalLink${n}.xml.rels`,
      XML_DECL +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${n === 1 ? 'gone.xlsx' : 'source.xlsx'}" TargetMode="External"/></Relationships>`,
    )
  }
  zip.file(
    'xl/worksheets/sheet1.xml',
    XML_DECL +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><f>[2]Sheet1!A1*2</f><v>84</v></c><c r="B1"><f>MATCH(20,[2]Sheet1!A1:A3,0)</f><v>2</v></c></row><row r="2"><c r="A2" t="str"><f>[1]Sheet1!A1&amp;"!"</f><v>gone!</v></c></row></sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/// planCellEditsToXlsx takes one long ordered list; these slices spell out
/// the slots around externalLinkOps (29th) so the test call sites stay
/// position-safe. Slot comments follow the signature order.
const HEAD: unknown[] = [
  [], // structuralOps
  [], // chartEdits
  undefined, // sheetPlan
  [], // filterStates
  [], // hyperlinkEdits
  [], // cfStates
  [], // dvStates
  [], // sheetProtections
  null, // definedNamesState
  [], // visualAdditions
  [], // pageSetupStates
  [], // noteStates
  [], // tableAdditions
  [], // pivotAdditions
  [], // pivotCacheRefreshPaths
  [], // pivotRefreshUpdates
  [], // visualEdits
  [], // sparklineAdditions
  [], // formulaValues
]

const TAIL = (ops: ExternalLinkOps): unknown[] => [
  null, // themeState
  null, // workbookProtectionState
  [], // protectedRangeStates
  [], // bulkConstantFills
  [], // tabColorStates
  [], // tableEdits
  [], // cellImageAdditions
  ops, // externalLinkOps
  null, // redactionStates
]

async function plan(buffer: Buffer, edits: readonly CellEdit[], ops: ExternalLinkOps) {
  // the planner's long positional signature is spelled out in HEAD/TAIL
  const call = planCellEditsToXlsx as unknown as (
    source: Awaited<ReturnType<typeof createBufferEntrySource>>,
    edits: readonly CellEdit[],
    ...rest: unknown[]
  ) => Promise<MutationPlan>
  return call(await createBufferEntrySource(buffer), edits, ...HEAD, ...TAIL(ops))
}

/// Reassembles a mutation plan into a zip buffer, the way the in-memory
/// assembler would apply it.
async function apply(buffer: Buffer, mutation: MutationPlan): Promise<JSZip> {
  const zip = await JSZip.loadAsync(buffer)
  for (const entry of mutation.removedEntries) zip.remove(entry)
  for (const [path, content] of mutation.replaced) zip.file(path, content)
  for (const [path, content] of mutation.added) zip.file(path, content)
  return zip
}

describe('planCellEditsToXlsx externalLinkOps', () => {
  it('breaks a link on the final contents, renumbering survivors', async () => {
    const source = await dependentBuffer()
    // the renderer journals the broken cell as a literal before saving
    const edits: CellEdit[] = [
      { sheetName: 'Sheet1', row: 1, column: 0, writeValue: true, cell: { value: 'gone!' } },
    ]
    const zip = await apply(source, await plan(source, edits, { broken: [1], retargets: [] }))
    const workbook = await zip.file('xl/workbook.xml')!.async('string')
    expect(workbook).toContain('<externalReference r:id="rId5"/>')
    expect(workbook).not.toContain('rId4')
    expect(Object.keys(zip.files)).not.toContain('xl/externalLinks/externalLink1.xml')
    expect(Object.keys(zip.files)).toContain('xl/externalLinks/externalLink2.xml')
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
    expect(contentTypes).not.toContain('externalLink1.xml')
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    // survivors renumber [2]→[1]; the journaled cell is a literal, not a formula
    expect(sheet).toContain('<f>[1]Sheet1!A1*2</f>')
    expect(sheet).not.toContain('[2]Sheet1')
    expect(sheet).not.toContain('<f>[1]Sheet1!A1&amp;')
    expect(sheet).toContain('gone!')
  })

  it('breaks a link whose formulas are still on the file, turning them literal', async () => {
    const source = await dependentBuffer()
    const zip = await apply(source, await plan(source, [], { broken: [1], retargets: [] }))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    // the [1] formula became its cached value; the string result is an inlineStr
    expect(sheet).not.toContain('<f>[1]Sheet1!A1&amp;')
    expect(sheet).toMatch(
      /<c r="A2" t="inlineStr"><is><t xml:space="preserve">gone!<\/t><\/is><\/c>/,
    )
  })

  it('retargets a link by rewriting only its own rels', async () => {
    const source = await dependentBuffer()
    const zip = await apply(
      source,
      await plan(source, [], {
        broken: [],
        retargets: [{ index: 2, target: 'file:///C:/data/new-source.xlsx' }],
      }),
    )
    const rels = await zip.file('xl/externalLinks/_rels/externalLink2.xml.rels')!.async('string')
    expect(rels).toContain('Target="file:///C:/data/new-source.xlsx"')
    expect(rels).toContain('TargetMode="External"')
    // formulas, sheet names and the other link stay untouched
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<f>[2]Sheet1!A1*2</f><v>84</v>')
    expect(Object.keys(zip.files)).toContain('xl/externalLinks/externalLink1.xml')
    const workbook = await zip.file('xl/workbook.xml')!.async('string')
    expect(workbook).toContain('<externalReference r:id="rId4"/>')
  })

  it('fails a retarget naming an index the workbook does not have', async () => {
    const source = await dependentBuffer()
    await expect(
      plan(source, [], { broken: [], retargets: [{ index: 7, target: 'x.xlsx' }] }),
    ).rejects.toThrow('no external link 7')
  })

  it('fails a break naming an index the workbook does not have', async () => {
    const source = await dependentBuffer()
    await expect(plan(source, [], { broken: [4], retargets: [] })).rejects.toThrow(
      'no external link 4',
    )
  })
})
