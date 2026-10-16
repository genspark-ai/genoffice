import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyCellEditsToXlsx,
  assertOnlyTouchedEntriesChanged,
  type CellEdit,
  type SheetStructuralOps,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import {
  REDACTION_MIRROR_PREFIX,
  REDACTION_PART_PATH,
  parseRedactionNameMirror,
  parseRedactionPart,
  type SheetRedactionState,
} from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`

/// A one-sheet workbook whose B2 holds the value the reader wants withheld.
async function buildFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>13800138000</t></is></c></row></sheetData></worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const EDIT: CellEdit = {
  sheetName: 'Sheet1',
  row: 0,
  column: 0,
  writeValue: true,
  cell: { value: 'edited' },
}

const MARK = { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }

const STATES: SheetRedactionState[] = [{ sheetName: 'Sheet1', marks: [MARK] }]

/**
 * `applyCellEditsToXlsx` takes one long positional list and the redaction
 * states sit at its end. Building the list by name here keeps the call sites
 * from becoming a wall of `undefined` that a future signature change would
 * silently re-target.
 */
function argumentsFor(
  source: Buffer,
  redactionStates: readonly SheetRedactionState[],
  plan?: SheetEditPlan,
  structuralOps: readonly SheetStructuralOps[] = [],
) {
  return [
    source,
    [EDIT],
    structuralOps, // structuralOps
    [] as never[], // chartEdits
    plan, // sheetPlan
    [] as never[], // filterStates
    [] as never[], // hyperlinkEdits
    [] as never[], // cfStates
    [] as never[], // dvStates
    [] as never[], // sheetProtections
    null, // definedNamesState
    [] as never[], // pageSetupStates
    [] as never[], // noteStates
    [] as never[], // formulaValues
    redactionStates,
  ] as const
}

const save = async (
  redactionStates: readonly SheetRedactionState[],
  plan?: SheetEditPlan,
  source?: Buffer,
) => applyCellEditsToXlsx(...argumentsFor(source ?? (await buildFixture()), redactionStates, plan))

async function savedParts(buffer: Buffer): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(buffer)
  const parts = new Map<string, string>()
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue
    parts.set(name, await entry.async('string'))
  }
  return parts
}

/// A part the test is about to assert on, named explicitly so a missing entry
/// says which one rather than failing on a `.map` of undefined.
function requiredPart(parts: Map<string, string>, name: string): string {
  const text = parts.get(name)
  if (text === undefined) throw new Error(`${name} is missing from the saved package`)
  return text
}

describe('the redaction part survives a real save', () => {
  it('lands in the package, declared, and reads back', async () => {
    const mutation = await save(STATES)
    // The save's own guard: an undeclared or unexpected part change aborts here.
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()

    const parts = await savedParts(mutation.buffer)
    const raw = parts.get(REDACTION_PART_PATH)
    expect(raw, 'the part is missing from the saved package').toBeDefined()
    expect(parseRedactionPart(raw!)).toEqual(STATES)

    // A part is only properly part of the package once both declarations exist.
    expect(parts.get('xl/_rels/workbook.xml.rels')).toContain('Target="gxRedactions.json"')
    expect(parts.get('[Content_Types].xml')).toContain(`PartName="/${REDACTION_PART_PATH}"`)
  })

  it('leaves the withheld value itself untouched in the worksheet', async () => {
    // A mark must never cost the reader their data. The cell keeps its value;
    // only the model's view of it changes.
    const mutation = await save(STATES)
    const parts = await savedParts(mutation.buffer)
    expect(parts.get('xl/worksheets/sheet1.xml')).toContain('13800138000')
  })

  it('rewrites the part in place on a second save instead of failing', async () => {
    const first = await save(STATES)
    // Re-saving the already-marked file: the part exists, so it must be a
    // replacement. Treating it as an addition trips assertManifestPreserved's
    // "should have created … but it already existed".
    const second = await applyCellEditsToXlsx(...argumentsFor(first.buffer, STATES))
    expect(() => assertOnlyTouchedEntriesChanged(second)).not.toThrow()
    const parts = await savedParts(second.buffer)
    expect(parseRedactionPart(parts.get(REDACTION_PART_PATH)!)).toEqual(STATES)
  })

  it('carries the marks across a sheet rename done in the same save', async () => {
    // The part is keyed by sheet name, so a rename that did not move the marks
    // would orphan them — and the withheld values would go back to the model
    // with nothing on screen saying so.
    const plan: SheetEditPlan = {
      renames: [{ sheetName: 'Sheet1', newName: 'Q3' }],
      removals: [],
      additions: [],
      order: ['Q3'],
    }
    const mutation = await save(STATES, plan)
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const parts = await savedParts(mutation.buffer)
    expect(parseRedactionPart(parts.get(REDACTION_PART_PATH)!)).toEqual([
      { ...STATES[0], sheetName: 'Q3' },
    ])
  })

  it('adds no part at all when nothing is withheld', async () => {
    const mutation = await save([])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const parts = await savedParts(mutation.buffer)
    expect(parts.has(REDACTION_PART_PATH)).toBe(false)
    expect(parts.get('xl/_rels/workbook.xml.rels')).toBe(WORKBOOK_RELS)
    expect(parts.get('[Content_Types].xml')).toBe(CONTENT_TYPES)
  })
})

/**
 * The part is the record, but it is not one every program keeps: LibreOffice
 * rebuilds the package and drops parts it does not model, along with their
 * relationships and content-type overrides. These tests take the part away and
 * check that the marks are still readable — the mirror in workbook.xml is what
 * makes that true, and it is written on the same save as the part.
 */
describe('the defined-name mirror survives a reader that drops the part', () => {
  it('lands in workbook.xml on every save that carries marks', async () => {
    const mutation = await save(STATES)
    // The save's own guard: the mirror rewrites workbook.xml, so it has to be
    // accounted for rather than slipping in as an unexplained change.
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const workbookXml = requiredPart(await savedParts(mutation.buffer), 'xl/workbook.xml')
    expect(workbookXml).toContain(REDACTION_MIRROR_PREFIX)
    expect(parseRedactionNameMirror(workbookXml, ['Sheet1'])).toEqual(STATES)
  })

  it('still names the marks when the part is taken away', async () => {
    // LibreOffice's exact shape, reproduced without LibreOffice: workbook.xml
    // is the file it keeps, and the part, the relationship, and the content-type
    // override are the three it loses.
    const mutation = await save(STATES)
    const zip = await JSZip.loadAsync(mutation.buffer)
    zip.remove(REDACTION_PART_PATH)
    zip.remove('xl/_rels/workbook.xml.rels')
    const rebuilt = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    const parts = await savedParts(rebuilt)

    expect(parts.has(REDACTION_PART_PATH)).toBe(false)
    expect(parts.has('xl/_rels/workbook.xml.rels')).toBe(false)
    // The point of the whole mirror: the marks are still there.
    expect(parseRedactionNameMirror(requiredPart(parts, 'xl/workbook.xml'), ['Sheet1'])).toEqual(
      STATES,
    )
  })

  it('follows a sheet rename done in the same save', async () => {
    const plan: SheetEditPlan = {
      renames: [{ sheetName: 'Sheet1', newName: 'Q3' }],
      removals: [],
      additions: [],
      order: ['Q3'],
    }
    const mutation = await save(STATES, plan)
    const workbookXml = requiredPart(await savedParts(mutation.buffer), 'xl/workbook.xml')
    expect(parseRedactionNameMirror(workbookXml, ['Q3'])).toEqual([
      { ...STATES[0], sheetName: 'Q3' },
    ])
  })

  it('is removed with the marks, so an un-hidden cell cannot come back', async () => {
    const marked = await save(STATES)
    const cleared = await applyCellEditsToXlsx(
      ...argumentsFor(marked.buffer, [{ sheetName: 'Sheet1', marks: [] }]),
    )
    expect(() => assertOnlyTouchedEntriesChanged(cleared)).not.toThrow()
    const workbookXml = requiredPart(await savedParts(cleared.buffer), 'xl/workbook.xml')
    expect(workbookXml).not.toContain(REDACTION_MIRROR_PREFIX)
    expect(parseRedactionNameMirror(workbookXml, ['Sheet1'])).toEqual([])
  })

  it('leaves a workbook that withholds nothing without a mirror', async () => {
    const mutation = await save([])
    const workbookXml = requiredPart(await savedParts(mutation.buffer), 'xl/workbook.xml')
    expect(workbookXml).not.toContain(REDACTION_MIRROR_PREFIX)
  })
})

describe('the mirror survives a later save', () => {
  it('reads back after a structural edit rewrote the defined names', async () => {
    // `shiftDefinedNames` rewrites every definedName body so the reader's own
    // names follow the rows a save inserts. The mirror is written through that
    // pass on the next save, and it has to come out the other side readable —
    // a mangled payload would look exactly like "nothing is withheld".
    const marked = await save(STATES)
    const second = await applyCellEditsToXlsx(
      ...argumentsFor(marked.buffer, STATES, undefined, [
        { sheetName: 'Sheet1', ops: [{ kind: 'insert-rows', index: 0, count: 2 }] },
      ]),
    )
    const parts = await savedParts(second.buffer)
    const workbookXml = requiredPart(parts, 'xl/workbook.xml')
    // The structural pass moved the part's mark to row 3; the mirror must agree,
    // or the two copies would withhold different cells after a row insert.
    const moved = [{ ...MARK, startRow: 3, endRow: 3 }]
    const expected = [{ sheetName: 'Sheet1', marks: moved }]
    expect(parseRedactionPart(requiredPart(parts, REDACTION_PART_PATH))).toEqual(expected)
    expect(parseRedactionNameMirror(workbookXml, ['Sheet1'])).toEqual(expected)
  })
})
