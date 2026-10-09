import { describe, expect, it } from 'vitest'

import {
  REDACTION_MIRROR_PREFIX,
  REDACTION_PART_PATH,
  RedactionError,
  applyRedactionNameMirror,
  applyRedactionPart,
  parseRedactionNameMirror,
  parseRedactionPart,
  rekeyRedactionStates,
  serializeRedactionPart,
  type RedactionMark,
  type RedactionPackage,
  type SheetRedactionState,
} from '../src/gateway/xlsx-redaction'
import { shiftDefinedNames } from '../src/gateway/xlsx-structure'

const MARK = { startRow: 1, endRow: 3, startColumn: 2, endColumn: 2, label: 'client phone' }

const STATES: SheetRedactionState[] = [{ sheetName: 'Sheet1', marks: [MARK] }]

/// A stand-in for PackageEditor that records what the save asked for, so the
/// tests can assert on the plan rather than on a real package.
function fakePackage(initial: Record<string, string> = {}): RedactionPackage & {
  added: string[]
  removed: string[]
  written: string[]
  touched: Set<string>
  files: Map<string, string>
} {
  const files = new Map(Object.entries(initial))
  const added: string[] = []
  const removed: string[] = []
  const written: string[] = []
  return {
    files,
    added,
    removed,
    written,
    touched: new Set<string>(),
    has: (path) => Promise.resolve(files.has(path)),
    readText: (path) => {
      const content = files.get(path)
      if (content === undefined) throw new Error(`Workbook is missing ${path}.`)
      return Promise.resolve(content)
    },
    write(path, content) {
      written.push(path)
      files.set(path, content)
    },
    add(path, content) {
      added.push(path)
      files.set(path, content)
    },
    remove(path) {
      removed.push(path)
      files.delete(path)
    },
  }
}

const RELS =
  '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="x" Target="a.xml"/></Relationships>'
const TYPES = '<Types xmlns="x"><Override PartName="/xl/workbook.xml" ContentType="y"/></Types>'

describe('redaction part round trip', () => {
  it('reads back exactly what it wrote, label byte-exact', () => {
    // "Client Phone" is the case a defined-name carrier cannot express: Excel's
    // name grammar allows no space. The label is the only string the model is
    // permitted to see, so it has to survive verbatim.
    const label = 'Client Phone / Primary Contact'
    const states: SheetRedactionState[] = [{ sheetName: 'Sheet 2', marks: [{ ...MARK, label }] }]
    expect(parseRedactionPart(serializeRedactionPart(states))).toEqual(states)
  })

  it('serializes the same bytes whatever order the marks arrive in', () => {
    // Order-independent output is what lets an unchanged workbook re-save
    // without the part showing up as a touched entry every time.
    const a: SheetRedactionState = {
      sheetName: 'A',
      marks: [MARK, { startRow: 0, endRow: 0, startColumn: 5, endColumn: 5, label: 'x' }],
    }
    const reversed = [...a.marks].reverse()
    const b: SheetRedactionState = { sheetName: 'A', marks: reversed }
    const c: SheetRedactionState = { sheetName: 'A', marks: [a.marks[0]!, a.marks[1]!] }
    expect(serializeRedactionPart([a])).toBe(serializeRedactionPart([b]))
    expect(serializeRedactionPart([c])).toBe(serializeRedactionPart([a]))
  })

  it('drops sheets that hold no marks', () => {
    const text = serializeRedactionPart([
      { sheetName: 'Empty', marks: [] },
      { sheetName: 'Full', marks: [MARK] },
    ])
    expect(parseRedactionPart(text).map((s) => s.sheetName)).toEqual(['Full'])
  })
})

describe('redaction part parsing fails closed', () => {
  // Every one of these is a part that records which values the reader hid.
  // Degrading any of them to "no marks" would hand the model exactly the
  // values that were withheld, so each must refuse the save instead.
  const broken: [string, string][] = [
    ['not JSON at all', 'not JSON at all'],
    ['a JSON array', '[]'],
    ['a missing version', JSON.stringify({ sheets: [] })],
    ['an unknown version', JSON.stringify({ version: 99, sheets: [] })],
    ['a missing sheets array', JSON.stringify({ version: 1 })],
    [
      'a mark with no label',
      JSON.stringify({
        version: 1,
        sheets: [{ name: 'S', marks: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }] }],
      }),
    ],
    [
      'a negative row',
      JSON.stringify({
        version: 1,
        sheets: [
          {
            name: 'S',
            marks: [{ startRow: -1, endRow: 0, startColumn: 0, endColumn: 0, label: 'l' }],
          },
        ],
      }),
    ],
    [
      'a fractional row',
      JSON.stringify({
        version: 1,
        sheets: [
          {
            name: 'S',
            marks: [{ startRow: 0.5, endRow: 1, startColumn: 0, endColumn: 0, label: 'l' }],
          },
        ],
      }),
    ],
    [
      'a mark ending before it starts',
      JSON.stringify({
        version: 1,
        sheets: [
          {
            name: 'S',
            marks: [{ startRow: 5, endRow: 2, startColumn: 0, endColumn: 0, label: 'l' }],
          },
        ],
      }),
    ],
    ['a sheet with no name', JSON.stringify({ version: 1, sheets: [{ marks: [] }] })],
  ]

  for (const [what, text] of broken) {
    it(`refuses ${what}`, () => {
      expect(() => parseRedactionPart(text)).toThrow(RedactionError)
    })
  }

  it('accepts the part it just wrote', () => {
    expect(() => parseRedactionPart(serializeRedactionPart(STATES))).not.toThrow()
  })
})

describe('rekeying for sheet edits done in the same save', () => {
  it('carries the marks onto a renamed sheet', () => {
    // The part is keyed by sheet name. A rename that did not move the marks
    // would orphan them, and the reader's withheld values would become
    // visible to the model with no warning anywhere.
    expect(rekeyRedactionStates(STATES, [{ sheetName: 'Sheet1', newName: 'Q3' }], [])).toEqual([
      { sheetName: 'Q3', marks: [MARK] },
    ])
  })

  it('drops the marks of a removed sheet', () => {
    expect(rekeyRedactionStates(STATES, [], ['Sheet1'])).toEqual([])
  })

  it('moves the marks with the rows the same save inserts', () => {
    // The repro: mark B2, insert two rows above it, save. The grid XML the
    // same save writes has the marked cell at row 3, so a part still saying
    // row 1 describes whatever shifted down into its place.
    expect(
      rekeyRedactionStates(
        STATES,
        [],
        [],
        [{ sheetName: 'Sheet1', ops: [{ kind: 'insert-rows', index: 0, count: 2 }] }],
      ),
    ).toEqual([
      {
        sheetName: 'Sheet1',
        marks: [{ ...MARK, startRow: MARK.startRow + 2, endRow: MARK.endRow + 2 }],
      },
    ])
  })

  it('drops a mark whose rows the same save deletes outright', () => {
    // The value it withheld is gone; a rectangle clamped onto a neighbour
    // would withhold something the reader never hid. A delete that only takes
    // part of the rectangle keeps the part that survives.
    expect(
      rekeyRedactionStates(
        STATES,
        [],
        [],
        [{ sheetName: 'Sheet1', ops: [{ kind: 'remove-rows', index: 0, count: MARK.endRow + 1 }] }],
      ),
    ).toEqual([{ sheetName: 'Sheet1', marks: [] }])
  })

  it('leaves the marks alone for a sheet with no structural ops', () => {
    expect(
      rekeyRedactionStates(
        STATES,
        [],
        [],
        [{ sheetName: 'Other', ops: [{ kind: 'insert-rows', index: 0, count: 5 }] }],
      ),
    ).toEqual(STATES)
  })

  it('leaves unrelated sheets alone', () => {
    expect(rekeyRedactionStates(STATES, [{ sheetName: 'Other', newName: 'X' }], ['Other'])).toEqual(
      STATES,
    )
  })
})

describe('writing the part into the package', () => {
  it('adds the part, a workbook relationship and a content-type override', () => {
    const pkg = fakePackage({
      'xl/_rels/workbook.xml.rels': RELS,
      '[Content_Types].xml': TYPES,
    })
    return applyRedactionPart(pkg, pkg.touched, STATES).then(() => {
      expect(pkg.added).toEqual([REDACTION_PART_PATH])
      // rId1 is taken, so the declaration must not collide with it.
      expect(pkg.files.get('xl/_rels/workbook.xml.rels')).toContain('Id="rId2"')
      expect(pkg.files.get('xl/_rels/workbook.xml.rels')).toContain('Target="gxRedactions.json"')
      expect(pkg.files.get('[Content_Types].xml')).toContain('PartName="/xl/gxRedactions.json"')
      expect([...pkg.touched]).toContain(REDACTION_PART_PATH)
    })
  })

  it('rewrites rather than re-adds on a second save', () => {
    // `add` would land the path in MutationPlan.added and trip
    // assertManifestPreserved's "should have created … but it already existed".
    const pkg = fakePackage({
      'xl/_rels/workbook.xml.rels': RELS,
      '[Content_Types].xml': TYPES,
      [REDACTION_PART_PATH]: serializeRedactionPart(STATES),
    })
    return applyRedactionPart(pkg, pkg.touched, STATES).then(() => {
      expect(pkg.added).toEqual([])
      expect(pkg.written).toContain(REDACTION_PART_PATH)
    })
  })

  it('does not declare the relationship or the override twice', () => {
    const pkg = fakePackage({
      'xl/_rels/workbook.xml.rels': RELS,
      '[Content_Types].xml': TYPES,
      [REDACTION_PART_PATH]: serializeRedactionPart(STATES),
    })
    return applyRedactionPart(pkg, pkg.touched, STATES)
      .then(() => applyRedactionPart(pkg, pkg.touched, STATES))
      .then(() => {
        const rels = pkg.files.get('xl/_rels/workbook.xml.rels') ?? ''
        expect(rels.match(/gxRedactions\.xml/g)?.length ?? 0).toBeLessThanOrEqual(1)
        const types = pkg.files.get('[Content_Types].xml') ?? ''
        expect(types.match(/gxRedactions\.json/g)?.length ?? 0).toBe(1)
      })
  })

  it('removes the part when the last mark is cleared', async () => {
    // The mirror of the case above, and the one that used to do nothing: an
    // early return left the part, so the next open read the stale marks and
    // the span came back after the reader had un-hidden it.
    const pkg = fakePackage({
      'xl/gxRedactions.json': '{"version":1,"sheets":[]}',
      'xl/_rels/workbook.xml.rels':
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="a.xml"/></Relationships>',
      '[Content_Types].xml':
        '<Types xmlns="x"><Override PartName="/xl/gxRedactions.json" ContentType="y"/></Types>',
    })
    await applyRedactionPart(pkg, pkg.touched, [{ sheetName: 'Sheet1', marks: [] }])
    expect(pkg.removed).toContain('xl/gxRedactions.json')
    expect(pkg.files.has('xl/gxRedactions.json')).toBe(false)
  })

  it('removes the relationship and the override with it', async () => {
    // An orphan override is worse than an orphan part: Excel complains about the
    // file on open, so all three have to go or the reader's file is damaged.
    const pkg = fakePackage({
      'xl/gxRedactions.json': '{"version":1,"sheets":[]}',
      'xl/_rels/workbook.xml.rels':
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId9" Type="' +
        'https://schemas.genspark.ai/genoffice/2026/relationships/redactions' +
        '" Target="gxRedactions.json"/></Relationships>',
      '[Content_Types].xml':
        '<Types xmlns="x"><Override PartName="/xl/gxRedactions.json" ContentType="y"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="y"/></Types>',
    })
    await applyRedactionPart(pkg, pkg.touched, [])
    expect(pkg.files.get('xl/_rels/workbook.xml.rels')).not.toContain('gxRedactions')
    const types = pkg.files.get('[Content_Types].xml') ?? ''
    expect(types).not.toContain('gxRedactions.json')
    // and the entries that were not ours are untouched
    expect(types).toContain('/xl/workbook.xml')
  })

  it('leaves a workbook with nothing withheld untouched', () => {
    // Growing an empty part would be a permanent difference from the file the
    // reader opened, for a workbook that never hid anything.
    const pkg = fakePackage({
      'xl/_rels/workbook.xml.rels': RELS,
      '[Content_Types].xml': TYPES,
    })
    return applyRedactionPart(pkg, pkg.touched, [{ sheetName: 'Sheet1', marks: [] }]).then(() => {
      expect(pkg.added).toEqual([])
      expect(pkg.written).toEqual([])
      expect(pkg.touched.size).toBe(0)
    })
  })

  it('overwrites a damaged part with the editor state rather than refusing', () => {
    // The editor's in-memory marks are authoritative for this session, so a
    // part that got damaged on disk is replaced rather than blocking the save.
    // The fail-closed rule belongs on the load path, where there is no editor
    // state to fall back on — see the parsing suite above.
    const pkg = fakePackage({
      'xl/_rels/workbook.xml.rels': RELS,
      '[Content_Types].xml': TYPES,
      [REDACTION_PART_PATH]: 'corrupted',
    })
    return applyRedactionPart(pkg, pkg.touched, STATES).then(() => {
      expect(parseRedactionPart(pkg.files.get(REDACTION_PART_PATH) ?? '')).toEqual(STATES)
    })
  })
})

const WORKBOOK_XML =
  '<workbook xmlns:r="urn:r"><sheets>' +
  '<sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Q3" sheetId="2" r:id="rId2"/>' +
  '</sheets></workbook>'
const SHEET_ORDER = ['Data', 'Q3']
const DATA_STATES: SheetRedactionState[] = [
  {
    sheetName: 'Data',
    marks: [
      { startRow: 1, endRow: 3, startColumn: 2, endColumn: 2, label: 'Client Phone' },
      { startRow: 7, endRow: 7, startColumn: 0, endColumn: 4, label: 'Payroll' },
    ],
  },
]

/// What a reader that rebuilds the package leaves behind: `xl/workbook.xml`
/// and nothing else the save added. Dropping the part, its relationship, and
/// its content-type override is the whole failure this mirror exists for.
const AFTER_THE_PART_IS_DROPPED = (workbookXml: string): string => workbookXml

/// The marks a mirror yields, failing with a readable message rather than an
/// "undefined" three assertions later.
function mirrorMarks(xml: string, sheetOrder: readonly string[] = SHEET_ORDER): RedactionMark[] {
  const recovered = parseRedactionNameMirror(xml, sheetOrder).flatMap((state) => state.marks)
  if (recovered.length === 0) throw new Error('the mirror produced no mark')
  return recovered
}

function mirrorMark(xml: string, sheetOrder: readonly string[] = SHEET_ORDER): RedactionMark {
  const [mark] = mirrorMarks(xml, sheetOrder)
  if (mark === undefined) throw new Error('the mirror produced no mark')
  return mark
}

describe('the defined-name mirror', () => {
  it('rebuilds every mark, the label byte-exact', () => {
    // "Client Phone" is the case a defined-name *name* cannot express — Excel's
    // name grammar allows no space. The label rides in the formula instead.
    const written = applyRedactionNameMirror(WORKBOOK_XML, DATA_STATES, SHEET_ORDER)
    expect(parseRedactionNameMirror(written, SHEET_ORDER)).toEqual(DATA_STATES)
  })

  it('survives a reader that dropped the part — the LibreOffice case', () => {
    // Save writes both copies. `soffice --convert-to xlsx` rebuilds the package
    // and keeps the second, so the marks are still there on the next open.
    const saved = applyRedactionNameMirror(WORKBOOK_XML, DATA_STATES, SHEET_ORDER)
    expect(AFTER_THE_PART_IS_DROPPED(saved)).toContain(REDACTION_MIRROR_PREFIX)
    expect(parseRedactionNameMirror(AFTER_THE_PART_IS_DROPPED(saved), SHEET_ORDER)).toEqual(
      DATA_STATES,
    )
  })

  it('round-trips a label with quotes, an ampersand and non-Latin text', () => {
    // A formula string literal doubles its quotes and XML escapes five of
    // them, so a label like this exercises three escaping layers at once.
    const label = '雇主 "Acme & Co." 的电话 +86-13800138000'
    const states: SheetRedactionState[] = [{ sheetName: 'Q3', marks: [{ ...MARK, label }] }]
    const written = applyRedactionNameMirror(WORKBOOK_XML, states, SHEET_ORDER)
    expect(mirrorMark(written, SHEET_ORDER).label).toBe(label)
  })

  it('records the sheet by position, so a rename cannot orphan a mark', () => {
    // The accepted cost of the mirror: it names the sheet by its place in
    // workbook order. Renaming that sheet leaves the mark exactly where it was.
    const written = applyRedactionNameMirror(WORKBOOK_XML, DATA_STATES, SHEET_ORDER)
    expect(mirrorMarks(written, ['Data renamed', 'Q3'])).toHaveLength(2)
  })

  it('writes hidden, workbook-scoped names', () => {
    // Hidden keeps them out of the Name Manager and out of
    // applyDefinedNamesState's rewrite; workbook-scoped keeps a deleted sheet
    // from renumbering every localSheetId behind it.
    const written = applyRedactionNameMirror(WORKBOOK_XML, DATA_STATES, SHEET_ORDER)
    expect(written).toContain('hidden="1"')
    expect(written).not.toContain('localSheetId')
    expect([...written.matchAll(/name="(_gxRedaction_\d+)"/g)].map((m) => m[1])).toEqual([
      `${REDACTION_MIRROR_PREFIX}1`,
      `${REDACTION_MIRROR_PREFIX}2`,
    ])
  })

  it('leaves a workbook with nothing withheld untouched', () => {
    expect(
      applyRedactionNameMirror(WORKBOOK_XML, [{ sheetName: 'Data', marks: [] }], SHEET_ORDER),
    ).toBe(WORKBOOK_XML)
  })

  it("leaves the reader's own defined names alone", () => {
    const withNames = WORKBOOK_XML.replace(
      '</sheets>',
      '</sheets><definedNames><definedName name="Total">Data!$C$1</definedName></definedNames>',
    )
    const written = applyRedactionNameMirror(withNames, DATA_STATES, SHEET_ORDER)
    expect(written).toContain('<definedName name="Total">Data!$C$1</definedName>')
    // ...and clearing the marks must take only ours, never theirs.
    const cleared = applyRedactionNameMirror(
      written,
      [{ sheetName: 'Data', marks: [] }],
      SHEET_ORDER,
    )
    expect(cleared).toContain('<definedName name="Total">Data!$C$1</definedName>')
    expect(cleared).not.toContain(REDACTION_MIRROR_PREFIX)
  })

  it('clearing the last mark leaves nothing behind', () => {
    // The mirror of the part's "clearing has to undo what setting did": a
    // stale entry here would bring the span back after the reader un-hid it.
    const written = applyRedactionNameMirror(WORKBOOK_XML, DATA_STATES, SHEET_ORDER)
    expect(applyRedactionNameMirror(written, [], SHEET_ORDER)).not.toContain(
      REDACTION_MIRROR_PREFIX,
    )
  })

  it('re-saving an unchanged workbook produces identical bytes', () => {
    // Otherwise every save would show the workbook as touched and "did
    // anything change?" would stop being answerable.
    const once = applyRedactionNameMirror(WORKBOOK_XML, DATA_STATES, SHEET_ORDER)
    expect(applyRedactionNameMirror(once, DATA_STATES, SHEET_ORDER)).toBe(once)
  })

  it('keeps a mark whose label is too long for one entry, and says it was cut', () => {
    // Excel rejects a defined name over 255 characters, so the entry cannot
    // hold the whole label. Dropping the entry instead would drop the mark,
    // and the mark is what withholds the cell — so the label gives way, with
    // an ellipsis so nothing passes itself off as the whole string.
    const label = 'x'.repeat(400)
    const states: SheetRedactionState[] = [{ sheetName: 'Data', marks: [{ ...MARK, label }] }]
    const written = applyRedactionNameMirror(WORKBOOK_XML, states, SHEET_ORDER)
    const formula = /<definedName\b[^>]*>[^<]*<\/definedName>/.exec(written)?.[0] ?? ''
    expect(Array.from(formula)).not.toHaveLength(0)
    const mark = mirrorMark(written, SHEET_ORDER)
    expect(mark).toMatchObject({
      startRow: MARK.startRow,
      startColumn: MARK.startColumn,
      endRow: MARK.endRow,
      endColumn: MARK.endColumn,
    })
    expect(mark.label.endsWith('…')).toBe(true)
    expect(mark.label.length).toBeLessThan(label.length)
  })

  it('drops marks for a sheet the file no longer has', () => {
    const written = applyRedactionNameMirror(WORKBOOK_XML, DATA_STATES, ['Q3'])
    expect(parseRedactionNameMirror(written, ['Q3'])).toEqual([])
  })
})

describe('reading the mirror', () => {
  const mirrored = (label = 'Client Phone') =>
    applyRedactionNameMirror(
      WORKBOOK_XML,
      [{ sheetName: 'Data', marks: [{ ...MARK, label }] }],
      SHEET_ORDER,
    )

  it('reads nothing as nothing withheld', () => {
    expect(parseRedactionNameMirror(WORKBOOK_XML, SHEET_ORDER)).toEqual([])
  })

  const refused: readonly (readonly [string, string])[] = [
    [
      'an entry that is not the mirror at all',
      applyRedactionNameMirror(WORKBOOK_XML, [], SHEET_ORDER).replace(
        '</sheets>',
        '</sheets><definedNames><definedName name="_gxRedaction_1" hidden="1">"nonsense"</definedName></definedNames>',
      ),
    ],
    ['a future mirror version', mirrored().replace('{""v"":1', '{""v"":2')],
    [
      'a sheet position the workbook does not have',
      applyRedactionNameMirror(
        WORKBOOK_XML,
        [{ sheetName: 'Data', marks: [MARK] }],
        SHEET_ORDER,
      ).replace('""sheet"":0', '""sheet"":9'),
    ],
    ['a rectangle that ends before it starts', mirrored().replace('""to"":[3,2]', '""to"":[0,0]')],
    ['a mark with no label', mirrored().replace('""label"":""Client Phone""}', '""label"":""""}')],
  ]
  for (const [what, xml] of refused) {
    it(`refuses ${what} rather than read it as nothing withheld`, () => {
      expect(() => parseRedactionNameMirror(xml, SHEET_ORDER)).toThrow(RedactionError)
    })
  }
})

describe('the mirror against the defined-name rewriter', () => {
  it('is left byte-identical by a structural edit', () => {
    // Every save that shifts rows rewrites all definedName bodies so the
    // reader's own names follow. The mirror is written through that pass, and
    // it has to survive it: a mangled payload is indistinguishable from "this
    // workbook withholds nothing" to whoever reads it next.
    const written = applyRedactionNameMirror(WORKBOOK_XML, DATA_STATES, SHEET_ORDER)
    const shifted = shiftDefinedNames(written, 'Data', [
      { kind: 'insert-rows', index: 0, count: 2 },
    ])
    expect(shifted).toBe(written)
  })
})
