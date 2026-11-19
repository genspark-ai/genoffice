import { describe, expect, it } from 'vitest'

import { dateStyleIndexes } from '../src/gateway/xlsx-gateway'
import { parseStylesheetFormats } from '../src/gateway/xlsx-style-read'
import { resolvedStyleEntries, styleEntrySlots } from '../src/gateway/xlsx-style-entries'
import { StylesheetEditor } from '../src/gateway/xlsx-styles'

// Shapes Hancom Cell (HCell 11.0) writes into xl/styles.xml: an entry is
// wrapped in mc:AlternateContent with xmlns:mc on the wrapper and xmlns:hs
// on mc:Choice, so the hs: prefix is bound only inside the Choice branch.
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const HS = 'http://schemas.haansoft.com/office/spreadsheet/8.0'
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

const wrap = (choice: string, fallback: string): string =>
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:hs="${HS}" Requires="hs">${choice}</mc:Choice>` +
  `<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`

const HANCOM_DXFS =
  `<styleSheet xmlns="${SML}">` +
  `<fonts count="1"><font><sz val="11"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border/></borders>` +
  `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>` +
  `<dxfs count="3">` +
  wrap(
    '<dxf hs:applyExtension="1"><font hs:extension="1"><b/><hs:ratio val="0"/></font></dxf>',
    '<dxf><font><b/></font></dxf>',
  ) +
  `<dxf><font><i/></font></dxf>` +
  wrap('<dxf hs:applyExtension="1"><font><u/></font></dxf>', '<dxf><font><u/></font></dxf>') +
  `</dxfs>` +
  `<tableStyles count="1"><tableStyle name="T" count="2"><tableStyleElement type="wholeTable" dxfId="1"/><tableStyleElement type="headerRow" dxfId="2"/></tableStyle></tableStyles>` +
  `</styleSheet>`

const HANCOM_FONTS_XFS =
  `<styleSheet xmlns="${SML}">` +
  `<fonts count="3"><font><sz val="11"/></font>` +
  wrap('<font hs:extension="1"><b/><sz val="20"/></font>', '<font><b/><sz val="18"/></font>') +
  `<font><sz val="9"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border/></borders>` +
  `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
  wrap(
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1" hs:applyExtension="1"/>',
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>',
  ) +
  `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/></cellXfs>` +
  `</styleSheet>`

/// Minimal namespace-scope check: every prefixed element or attribute name
/// must be bound by an xmlns:prefix on itself or an open ancestor — what a
/// namespace-aware parser (the sidecar's roxmltree, Excel) enforces.
function unboundPrefixes(xml: string): string[] {
  const unbound: string[] = []
  const scopes: Set<string>[] = []
  for (const match of xml.matchAll(/<(\/?)([^\s/>!?]+)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g)) {
    const [, closing, name, attributes, selfClosing] = match
    if (closing) {
      scopes.pop()
      continue
    }
    const declared = new Set(
      [...attributes!.matchAll(/\sxmlns:([\w.-]+)\s*=/g)].map((declaration) => declaration[1]!),
    )
    const inScope = (prefix: string): boolean =>
      prefix === 'xml' || declared.has(prefix) || scopes.some((scope) => scope.has(prefix))
    const used = [
      name!,
      ...[...attributes!.matchAll(/\s([\w.-]+:[\w.-]+)\s*=/g)].map((attribute) => attribute[1]!),
    ]
    for (const qualified of used) {
      const colon = qualified.indexOf(':')
      if (colon <= 0) continue
      const prefix = qualified.slice(0, colon)
      if (prefix !== 'xmlns' && !inScope(prefix)) unbound.push(qualified)
    }
    if (!selfClosing) scopes.push(declared)
  }
  return unbound
}

describe('StylesheetEditor on mc:AlternateContent entries (Hancom Cell)', () => {
  it('keeps a Hancom stylesheet well-formed after a style edit', () => {
    const editor = new StylesheetEditor(HANCOM_DXFS)
    editor.resolveStyle(0, { bold: true })
    const xml = editor.serialize()
    // Before the fix the Choice dxf was hoisted out of its wrapper, leaving
    // hs:applyExtension unbound: "an unknown namespace prefix 'hs'".
    expect(unboundPrefixes(xml)).toEqual([])
  })

  it('counts a wrapped entry once, so later dxfIds keep pointing at the same format', () => {
    const editor = new StylesheetEditor(HANCOM_DXFS)
    editor.resolveStyle(0, { bold: true })
    const xml = editor.serialize()
    expect(xml).toContain('<dxfs count="3">')
    // dxfId 1 is still the italic entry the table style references.
    const dxfs = /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(xml)?.[1] ?? ''
    expect(dxfs).toContain('<dxf><font><i/></font></dxf>')
    expect(dxfs.match(/<mc:AlternateContent\b/g)).toHaveLength(2)
    expect(editor.internDxf('<dxf><fill/></dxf>')).toBe(3)
  })

  it('keeps wrapped entries byte for byte', () => {
    const editor = new StylesheetEditor(HANCOM_FONTS_XFS)
    editor.resolveStyle(0, { italic: true })
    const xml = editor.serialize()
    expect(xml).toContain(
      wrap('<font hs:extension="1"><b/><sz val="20"/></font>', '<font><b/><sz val="18"/></font>'),
    )
    expect(unboundPrefixes(xml)).toEqual([])
  })

  it('derives from the indexed entry, not from a shifted neighbour', () => {
    const editor = new StylesheetEditor(HANCOM_FONTS_XFS)
    // xf 2 is the 9pt one; flattening put the Fallback xf (18pt bold) there.
    const index = editor.resolveStyle(2, { italic: true })
    expect(index).toBe(3)
    const xml = editor.serialize()
    expect(xml).toContain('<cellXfs count="4">')
    expect(xml).toContain('<fonts count="4">')
    const fonts = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(xml)?.[1] ?? ''
    expect(fonts.endsWith('<font><i/><sz val="9"/></font>')).toBe(true)
  })

  it('derives a wrapped base from its Fallback, so the new entry carries no hs: markup', () => {
    const editor = new StylesheetEditor(HANCOM_FONTS_XFS)
    editor.resolveStyle(1, { italic: true })
    const xml = editor.serialize()
    const fonts = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(xml)?.[1] ?? ''
    expect(fonts.endsWith('<font><i/><b/><sz val="18"/></font>')).toBe(true)
    expect(unboundPrefixes(xml)).toEqual([])
  })

  it('keeps namespace declarations and extension attributes on the section start tag', () => {
    const source = HANCOM_FONTS_XFS.replace(
      '<cellXfs count="3">',
      `<cellXfs count="3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" x14ac:knownFonts="1">`,
    )
    const editor = new StylesheetEditor(source)
    editor.resolveStyle(0, { italic: true })
    const xml = editor.serialize()
    expect(xml).toContain('<cellXfs count="4" xmlns:x14ac=')
    expect(xml).toContain('x14ac:knownFonts="1"')
    expect(unboundPrefixes(xml)).toEqual([])
  })
})

describe('style entry slots', () => {
  it('reads a Choice-only block as its Choice and keeps its place', () => {
    const inner =
      `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="zz"><dxf><font><i/></font></dxf></mc:Choice></mc:AlternateContent>` +
      '<dxf><font><b/></font></dxf>'
    expect(styleEntrySlots(inner, 'dxf')).toHaveLength(2)
    expect(resolvedStyleEntries(inner, 'dxf')).toEqual([
      '<dxf><font><i/></font></dxf>',
      '<dxf><font><b/></font></dxf>',
    ])
  })

  it('unwraps a block nested inside another', () => {
    const inner =
      '<font><sz val="1"/></font>' +
      wrap(
        '<font><sz val="9"/></font>',
        wrap('<font><sz val="3"/></font>', '<font><sz val="2"/></font>'),
      ) +
      '<font><sz val="4"/></font>'
    expect(styleEntrySlots(inner, 'font')).toHaveLength(3)
    expect(resolvedStyleEntries(inner, 'font')).toEqual([
      '<font><sz val="1"/></font>',
      '<font><sz val="2"/></font>',
      '<font><sz val="4"/></font>',
    ])
  })

  it('refuses a block that does not stand for exactly one entry', () => {
    const source = HANCOM_FONTS_XFS.replace(
      '<font><sz val="9"/></font></fonts>',
      `<font><sz val="9"/></font><mc:AlternateContent xmlns:mc="${MC}"><mc:Fallback><font/><font/></mc:Fallback></mc:AlternateContent></fonts>`,
    )
    expect(() => new StylesheetEditor(source)).toThrow(/style edits cannot be saved/)
  })

  it('does not reuse a numFmtId taken by either branch of a wrapped format', () => {
    const source = HANCOM_FONTS_XFS.replace(
      '<fonts count="3">',
      `<numFmts count="1">${wrap('<numFmt numFmtId="170" formatCode="0.0" hs:x="1"/>', '<numFmt numFmtId="165" formatCode="0.0"/>')}</numFmts><fonts count="3">`,
    )
    const editor = new StylesheetEditor(source)
    editor.resolveStyle(0, { numberFormat: '0.000' })
    const xml = editor.serialize()
    expect(xml).toContain('<numFmts count="2">')
    expect(xml).toContain('<numFmt numFmtId="171" formatCode="0.000"/>')
    expect(unboundPrefixes(xml)).toEqual([])
  })
})

describe('read-only stylesheet views on wrapped entries', () => {
  it('index colors and fonts the way the sidecar does', () => {
    const formats = parseStylesheetFormats(HANCOM_FONTS_XFS)
    expect(formats.xfs).toEqual([
      { fillId: 0, fontId: 0 },
      { fillId: 0, fontId: 1 },
      { fillId: 0, fontId: 2 },
    ])
    expect(formats.fontColors).toHaveLength(3)
  })

  it('finds date cell formats at their own index', () => {
    const source = HANCOM_FONTS_XFS.replace(
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/></cellXfs>',
      '<xf numFmtId="14" fontId="2" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>',
    )
    expect([...dateStyleIndexes(source)]).toEqual([2])
  })
})

describe('date cell formats behind a wrapped numFmt', () => {
  // One numFmtId, two codes: the sidecar shows the Fallback's, so the text
  // read must not call the cell a date because the Choice branch is one.
  const withNumFmt = (choiceCode: string, fallbackCode: string): string =>
    HANCOM_FONTS_XFS.replace(
      '<fonts count="3">',
      `<numFmts count="1">${wrap(
        `<numFmt numFmtId="164" formatCode="${choiceCode}" hs:x="1"/>`,
        `<numFmt numFmtId="164" formatCode="${fallbackCode}"/>`,
      )}</numFmts><fonts count="3">`,
    ).replace(
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/></cellXfs>',
      '<xf numFmtId="164" fontId="2" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>',
    )

  it('reads the Fallback code, not the Choice one', () => {
    expect([...dateStyleIndexes(withNumFmt('yyyy-mm-dd', '0.00'))]).toEqual([])
    expect([...dateStyleIndexes(withNumFmt('0.00', 'yyyy-mm-dd'))]).toEqual([2])
  })

  it('ignores numFmt elements outside the numFmts section', () => {
    const source = HANCOM_FONTS_XFS.replace(
      '</cellXfs>',
      '</cellXfs><dxfs count="1"><dxf><numFmt numFmtId="0" formatCode="yyyy-mm-dd"/></dxf></dxfs>',
    )
    expect([...dateStyleIndexes(source)]).toEqual([])
  })
})
