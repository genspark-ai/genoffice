import { describe, expect, it } from 'vitest'

import {
  REDACTION_PART_PATH,
  RedactionError,
  applyRedactionPart,
  parseRedactionPart,
  rekeyRedactionStates,
  serializeRedactionPart,
  type RedactionPackage,
  type SheetRedactionState,
} from '../src/gateway/xlsx-redaction'

const MARK = { startRow: 1, endRow: 3, startColumn: 2, endColumn: 2, label: '客户电话' }

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
    const label = 'Client Phone / 主要客户'
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
