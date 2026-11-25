import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  computeListMarkerInfos,
  computeListMarkers,
  customEnumItems,
  customLevelFromNumberingLevel,
  customLevelXml,
  formatNumber,
  markerTabAdvance,
  mergeLevelXml,
  parseDocx,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import { unlinkListStyleXml } from '../src/patch'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const MULTILEVEL_NUMBERING =
  XML_DECL +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>' +
  '<w:lvl w:ilvl="2"><w:start w:val="3"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%3)"/></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="0"/>' +
  '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' +
  '<w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="5"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%2:"/></w:lvl></w:lvlOverride>' +
  '</w:num>' +
  '</w:numbering>'

describe('numbering definitions (word/numbering.xml)', () => {
  it('parses per-level numFmt / lvlText / start and lvlOverrides', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      numberingXml: MULTILEVEL_NUMBERING,
    })
    const doc = await parseDocx(bytes)

    const num1 = doc.numbering.get('1')!
    expect(num1.abstractNumId).toBe('0')
    expect(num1.levels[0]).toEqual({ numFmt: 'decimal', lvlText: '%1.', start: 1 })
    expect(num1.levels[1]).toEqual({ numFmt: 'decimal', lvlText: '%1.%2', start: 1 })
    expect(num1.levels[2]).toEqual({ numFmt: 'lowerLetter', lvlText: '%3)', start: 3 })
    expect(num1.startOverrides).toEqual({})

    const num2 = doc.numbering.get('2')!
    expect(num2.startOverrides).toEqual({ 0: 1 })
    // A full w:lvl inside lvlOverride replaces that level's definition; other levels
    // follow the abstractNum
    expect(num2.levels[1]).toEqual({ numFmt: 'upperRoman', lvlText: '%2:', start: 5 })
    expect(num2.levels[0]).toEqual({ numFmt: 'decimal', lvlText: '%1.', start: 1 })
  })

  it('reads level w:ind as logical start/end with an absent left = 0, and the level w:bidi', async () => {
    // Word probe 2026-10-20: a bidi level `right=360 hanging=360` hangs the marker
    // 18pt outside the start (right) margin and indents the end (left) side by 18pt
    const numbering =
      XML_DECL +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
      '<w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:lvlJc w:val="right"/>' +
      '<w:pPr><w:bidi/><w:ind w:right="360" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
      '<w:numFmt w:val="bullet"/><w:lvlText w:val="-"/>' +
      '<w:pPr><w:ind w:start="720" w:end="120" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
      '</w:numbering>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' +
          '<w:jc w:val="left"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: numbering,
      }),
    )
    const rtl = doc.numbering.get('1')?.levels[0]
    expect(rtl).toMatchObject({ indentLeft: 0, indentRight: 360, hanging: 360, bidi: true })
    expect(doc.numbering.get('2')?.levels[0]).toMatchObject({ indentLeft: 720, indentRight: 120 })
    expect(doc.numbering.get('2')?.levels[0]?.bidi).toBeUndefined()
    // the level's w:bidi makes the paragraph bidi: w:jc left is the start edge (visual right)
    const item = doc.blocks.find((b) => b.type === 'listItem')
    expect(item && 'format' in item ? item.format?.align : undefined).toBe('right')
  })

  it('keeps the bullet/ordered classification for list blocks', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>b</w:t></w:r></w:p>',
      withNumbering: true,
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].list?.kind).toBe('bullet')
    expect(doc.blocks[1].list?.kind).toBe('ordered')
    expect(doc.numbering.get('2')?.levels[0]?.lvlText).toBe('%1.')
  })
})

describe('mixed multilevel list kind', () => {
  const MIXED_NUMBERING =
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="9">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="7"><w:abstractNumId w:val="9"/></w:num>' +
    '</w:numbering>'

  it('classifies each level by its own numFmt, not level 0', async () => {
    const li = (ilvl: number, text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="7"/></w:numPr></w:pPr>` +
      `<w:r><w:t>${text}</w:t></w:r></w:p>`
    const doc = await parseDocx(
      await buildDocx({ bodyXml: li(0, 'top') + li(1, 'sub'), numberingXml: MIXED_NUMBERING }),
    )
    expect(doc.blocks[0].list).toMatchObject({ kind: 'ordered', ilvl: 0 })
    expect(doc.blocks[1].list).toMatchObject({ kind: 'bullet', ilvl: 1 })
  })

  it('falls back to the level-0 classification for undefined levels', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:numPr><w:ilvl w:val="4"/><w:numId w:val="7"/></w:numPr></w:pPr>' +
          '<w:r><w:t>deep</w:t></w:r></w:p>',
        numberingXml: MIXED_NUMBERING,
      }),
    )
    expect(doc.blocks[0].list).toMatchObject({ kind: 'ordered', ilvl: 4 })
  })
})

describe('mergeLevelXml keeps foreign-namespace children', () => {
  it('keeps w14 extension children of rPr and an mc:AlternateContent sibling', () => {
    // the child scan matched w: names only, so a w14:textOutline vanished from
    // the marker's rPr and the mc:AlternateContent was replaced by its own child
    const outline =
      '<w14:textOutline w14:w="9525"><w14:solidFill><w14:srgbClr w14:val="FF0000"/>' +
      '</w14:solidFill></w14:textOutline>'
    const alt =
      '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing/></mc:Choice></mc:AlternateContent>'
    const existing =
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>' +
      `<w:rPr><w:rFonts w:ascii="Symbol"/>${outline}</w:rPr>${alt}</w:lvl>`

    const merged = mergeLevelXml(existing, { numFmt: 'bullet', lvlText: '–', indentLeft: 720 }, 0)
    expect(merged).toContain(`<w:rPr><w:rFonts w:ascii="Symbol"/>${outline}</w:rPr>`)
    expect(merged).toContain(alt)
    expect(merged).toContain('<w:lvlText w:val="–"/>')

    // an rPr holding nothing but a w14 child survives instead of being deleted
    const extOnly = mergeLevelXml(
      `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:rPr>${outline}</w:rPr></w:lvl>`,
      { numFmt: 'bullet', lvlText: '–', indentLeft: 720 },
      0,
    )
    expect(extOnly).toContain(`<w:rPr>${outline}</w:rPr>`)
  })
})

describe('missing w:start default', () => {
  const NO_START_NUMBERING =
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'

  it('a w:lvl without w:start starts at 0, as Word renders it', async () => {
    const li = (text: string) =>
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
      `<w:r><w:t>${text}</w:t></w:r></w:p>`
    const doc = await parseDocx(
      await buildDocx({ bodyXml: li('a') + li('b'), numberingXml: NO_START_NUMBERING }),
    )
    expect(doc.numbering.get('1')!.levels[0].start).toBe(0)
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 0 },
      ],
      doc.numbering,
    )
    expect(markers).toEqual(['0.', '1.'])
  })

  it('letter formats clamp a missing or zero w:start to the first letter instead of an empty marker', async () => {
    const letters = NO_START_NUMBERING.replace('w:val="decimal"', 'w:val="lowerLetter"')
    const li = (text: string) =>
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
      `<w:r><w:t>${text}</w:t></w:r></w:p>`
    const doc = await parseDocx(
      await buildDocx({ bodyXml: li('a') + li('b'), numberingXml: letters }),
    )
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 0 },
      ],
      doc.numbering,
    )
    // the counter still starts at 0 (as for decimal); only the glyph is floored, like toRoman
    expect(markers[0]).toBe('a.')
    expect(markers.every((m) => m !== '.')).toBe(true)
    expect(formatNumber(0, 'lowerLetter')).toBe('a')
    expect(formatNumber(-3, 'upperLetter')).toBe('A')
  })
})

describe('greek letter formats', () => {
  it('lowerGreek walks the 24-letter alphabet without final sigma', () => {
    expect(formatNumber(1, 'lowerGreek')).toBe('α')
    expect(formatNumber(17, 'lowerGreek')).toBe('ρ')
    expect(formatNumber(18, 'lowerGreek')).toBe('σ')
    expect(formatNumber(24, 'lowerGreek')).toBe('ω')
    expect(formatNumber(25, 'lowerGreek')).toBe('αα')
    expect(formatNumber(49, 'lowerGreek')).toBe('ααα')
  })

  it('upperGreek skips the unassigned U+03A2 slot', () => {
    expect(formatNumber(1, 'upperGreek')).toBe('Α')
    expect(formatNumber(17, 'upperGreek')).toBe('Ρ')
    expect(formatNumber(18, 'upperGreek')).toBe('Σ')
    expect(formatNumber(24, 'upperGreek')).toBe('Ω')
    expect(formatNumber(25, 'upperGreek')).toBe('ΑΑ')
  })

  it('renders greek markers through lvlText', async () => {
    const numbering =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerGreek"/><w:lvlText w:val="%1."/></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '</w:numbering>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>', numberingXml: numbering }),
    )
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 0 },
      ],
      doc.numbering,
    )
    expect(markers).toEqual(['α.', 'β.'])
  })
})

describe('w:suff', () => {
  const SUFF_NUMBERING =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="space"/><w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="nothing"/><w:lvlText w:val="%2."/></w:lvl>' +
    '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%3."/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'

  it('parses the marker suffix kind, leaving the tab default unset', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: SUFF_NUMBERING,
      }),
    )
    const levels = doc.numbering.get('1')!.levels
    expect(levels[0].suff).toBe('space')
    expect(levels[1].suff).toBe('nothing')
    expect(levels[2].suff).toBeUndefined()
  })
})

describe('style-level numId="0" cancels inherited numbering', () => {
  const STYLES =
    '<w:style w:type="paragraph" w:styleId="NumberedBase"><w:name w:val="Numbered Base"/>' +
    '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="NoNumChild"><w:name w:val="No Num Child"/>' +
    '<w:basedOn w:val="NumberedBase"/>' +
    '<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="PlainChild"><w:name w:val="Plain Child"/>' +
    '<w:basedOn w:val="NumberedBase"/></w:style>'

  it('blocks basedOn numPr inheritance for the cancelling style only', async () => {
    const p = (style: string, text: string) =>
      `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: p('NumberedBase', 'a') + p('NoNumChild', 'b') + p('PlainChild', 'c'),
        extraStylesXml: STYLES,
        withNumbering: true,
      }),
    )
    expect(doc.blocks[0].list).toMatchObject({ numId: '2', ilvl: 0 })
    expect(doc.blocks[1].list).toBeUndefined()
    expect(doc.blocks[2].list).toMatchObject({ numId: '2', ilvl: 0 })
  })
})

describe('w:numStyleLink indirection', () => {
  const LINKED_NUMBERING =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="ListNumberStyle"/></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="ListNumberStyle"/>' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'

  it('resolves levels through numStyleLink -> styleLink', async () => {
    const li = (ilvl: number, text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
      `<w:r><w:t>${text}</w:t></w:r></w:p>`
    const doc = await parseDocx(
      await buildDocx({ bodyXml: li(0, 'a') + li(1, 'b'), numberingXml: LINKED_NUMBERING }),
    )
    expect(doc.numbering.get('1')!.levels[0]).toMatchObject({ numFmt: 'decimal', lvlText: '%1)' })
    expect(doc.blocks[0].list?.kind).toBe('ordered')
    expect(doc.blocks[1].list?.kind).toBe('bullet')
    expect(computeListMarkers([{ numId: '1', ilvl: 0 }], doc.numbering)).toEqual(['1)'])
  })

  it('survives a numStyleLink cycle', async () => {
    const cyclic =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="A"/><w:styleLink w:val="B"/></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="1"><w:numStyleLink w:val="B"/><w:styleLink w:val="A"/></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '</w:numbering>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>', numberingXml: cyclic }),
    )
    expect(doc.numbering.get('1')!.levels).toEqual({})
  })
})

describe('named list styles', () => {
  const LIST_STYLES =
    '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/></w:style>' +
    '<w:style w:type="numbering" w:styleId="Clauses" w:customStyle="1"><w:name w:val="Clauses"/>' +
    '<w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>'
  const STYLED_NUMBERING =
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="Clauses"/></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="Clauses"/>' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
    '</w:numbering>'
  const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
  const originals = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
    doc.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
  const LEVELS = [
    { numFmt: 'decimal', lvlText: '%1)', indentLeft: 360, hanging: 360 },
    { numFmt: 'lowerLetter', lvlText: '%2)', indentLeft: 720, hanging: 360 },
  ]

  it('lists the styles of type numbering with the levels of their abstractNum', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: P('x'),
        extraStylesXml: LIST_STYLES,
        numberingXml: STYLED_NUMBERING,
      }),
    )
    expect(doc.listStyles?.map((s) => s.styleId)).toEqual(['NoList', 'Clauses'])
    const [none, clauses] = doc.listStyles!
    expect(none).toMatchObject({ name: 'No List', isDefault: true, levels: {} })
    expect(none.numId).toBeUndefined()
    expect(clauses).toMatchObject({ name: 'Clauses', numId: '2', custom: true })
    expect(clauses.levels[0]).toMatchObject({ numFmt: 'upperRoman', lvlText: '%1.' })
    expect(doc.numbering.get('1')!.listStyleId).toBe('Clauses')
    expect(doc.numbering.get('2')!.listStyleId).toBe('Clauses')
    expect(doc.styles.has('Clauses')).toBe(false)
  })

  it('styleNums: a new w:num over an abstractNum that only links to the existing style', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: P('x'),
        extraStylesXml: LIST_STYLES,
        numberingXml: STYLED_NUMBERING,
      }),
    )
    const saved = await saveDocx(doc, originals(doc), {
      numbering: {
        styleNums: [
          {
            numId: '7',
            styleId: 'Clauses',
            define: { name: 'Clauses', numId: '8', levels: LEVELS },
          },
        ],
      },
    })
    const zip = await JSZip.loadAsync(saved)
    const numXml = await zip.file('word/numbering.xml')!.async('string')
    expect(numXml).toContain(
      '<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="multilevel"/><w:numStyleLink w:val="Clauses"/></w:abstractNum>',
    )
    expect(numXml).toContain('<w:num w:numId="7"><w:abstractNumId w:val="2"/></w:num>')
    // the style is already there: nothing is defined a second time
    expect(numXml).not.toContain('w:numId="8"')
    expect(numXml.match(/<w:styleLink /g)).toHaveLength(1)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml.match(/w:styleId="Clauses"/g)).toHaveLength(1)
    const reparsed = await parseDocx(saved)
    expect(reparsed.numbering.get('7')).toMatchObject({ listStyleId: 'Clauses' })
    expect(reparsed.numbering.get('7')!.levels[0]).toMatchObject({ numFmt: 'upperRoman' })
  })

  it('a list style name and level text holding `$&` are written as typed', async () => {
    const name = "Cost $& Tax $' $$"
    const levels = [{ ...LEVELS[0], lvlText: '$&%1$$' }, ...LEVELS.slice(1)]
    const doc = await parseDocx(await buildDocx({ bodyXml: P('x') }))
    const saved = await saveDocx(doc, originals(doc), {
      numbering: { styleDefs: [{ styleId: 'Costs', name, numId: '5', levels }] },
    })
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml.match(/<\/w:styles>/g)).toHaveLength(1)
    expect(
      (await zip.file('word/numbering.xml')!.async('string')).match(/<\/w:numbering>/g),
    ).toHaveLength(1)
    const style = (await parseDocx(saved)).listStyles!.find((s) => s.styleId === 'Costs')!
    expect(style.name).toBe(name)
    expect(style.levels[0].lvlText).toBe('$&%1$$')
  })

  it('a level edit keeps a literal `$$` in w:lvlText', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: P('x'), numberingXml: MULTILEVEL_NUMBERING }),
    )
    const saved = await saveDocx(doc, originals(doc), {
      numbering: {
        levelEdits: [
          {
            abstractNumId: '0',
            ilvl: 0,
            level: { numFmt: 'decimal', lvlText: '$$%1', indentLeft: 720 },
          },
        ],
      },
    })
    const numXml = await JSZip.loadAsync(saved).then((zip) =>
      zip.file('word/numbering.xml')!.async('string'),
    )
    expect(numXml).toContain('<w:lvlText w:val="$$%1"/>')
    expect((await parseDocx(saved)).numbering.get('1')!.levels[0].lvlText).toBe('$$%1')
  })

  it('a deleted list style leaves `$$` level text whole in the lists that copied it', () => {
    const xml =
      '<w:numbering><w:abstractNum w:abstractNumId="1"><w:styleLink w:val="S"/>' +
      '<w:lvl w:ilvl="0"><w:lvlText w:val="$$%1 $&amp;"/></w:lvl></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="S"/></w:abstractNum></w:numbering>'
    expect(unlinkListStyleXml(xml, 'S')).toContain(
      '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:lvlText w:val="$$%1 $&amp;"/></w:lvl></w:abstractNum>',
    )
  })

  it('styleDefs: a user-made list style is written though no list follows it, once', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: P('x') }))
    const saved = await saveDocx(doc, originals(doc), {
      numbering: {
        styleDefs: [{ styleId: 'MyList', name: 'My List', numId: '5', levels: LEVELS }],
        styleNums: [
          {
            numId: '6',
            styleId: 'MyList',
            define: { name: 'My List', numId: '7', levels: LEVELS, custom: true },
          },
        ],
      },
    })
    const zip = await JSZip.loadAsync(saved)
    const numXml = await zip.file('word/numbering.xml')!.async('string')
    expect(numXml.match(/<w:styleLink w:val="MyList"\/>/g)).toHaveLength(1)
    expect(numXml).not.toContain('w:numId="7"')
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml).toContain(
      '<w:style w:type="numbering" w:customStyle="1" w:styleId="MyList"><w:name w:val="My List"/>',
    )
    const reparsed = await parseDocx(saved)
    const style = reparsed.listStyles!.find((s) => s.styleId === 'MyList')!
    expect(style).toMatchObject({ name: 'My List', numId: '5', custom: true })
    expect(style.levels[1]).toMatchObject({ numFmt: 'lowerLetter', lvlText: '%2)' })
    expect(reparsed.numbering.get('6')!.listStyleId).toBe('MyList')
  })

  it('removedStyles: the style goes, and the lists that followed it keep its levels as their own', async () => {
    const item =
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: item,
        extraStylesXml: LIST_STYLES,
        numberingXml: STYLED_NUMBERING,
      }),
    )
    const saved = await saveDocx(doc, originals(doc), {
      numbering: { removedStyles: ['Clauses'] },
    })
    const zip = await JSZip.loadAsync(saved)
    const numXml = await zip.file('word/numbering.xml')!.async('string')
    expect(numXml).not.toContain('tyleLink')
    expect(numXml.match(/<w:lvl w:ilvl="0">/g)).toHaveLength(2)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml).not.toContain('Clauses')
    expect(stylesXml).toContain('w:styleId="NoList"')
    const reparsed = await parseDocx(saved)
    expect(reparsed.listStyles!.map((s) => s.styleId)).toEqual(['NoList'])
    expect(reparsed.numbering.get('1')!.listStyleId).toBeUndefined()
    expect(reparsed.numbering.get('1')!.levels[0]).toMatchObject({ numFmt: 'upperRoman' })
    expect(computeListMarkers([{ numId: '1', ilvl: 0 }], reparsed.numbering)).toEqual(['I.'])
  })

  it('styleNums with define: writes the style, its abstractNum and the linking list into a document with neither', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: P('x') }))
    const saved = await saveDocx(doc, originals(doc), {
      numbering: {
        styleNums: [
          {
            numId: '10',
            styleId: '1ai',
            define: { name: 'Outline List 1', numId: '11', levels: LEVELS },
          },
        ],
      },
    })
    const zip = await JSZip.loadAsync(saved)
    const numXml = await zip.file('word/numbering.xml')!.async('string')
    const defining =
      /<w:abstractNum w:abstractNumId="(\d+)"><w:multiLevelType w:val="multilevel"\/><w:styleLink w:val="1ai"\/><w:lvl /.exec(
        numXml,
      )
    const linking =
      /<w:abstractNum w:abstractNumId="(\d+)"><w:multiLevelType w:val="multilevel"\/><w:numStyleLink w:val="1ai"\/><\/w:abstractNum>/.exec(
        numXml,
      )
    expect(defining).not.toBeNull()
    expect(linking).not.toBeNull()
    expect(numXml).toContain(
      `<w:num w:numId="11"><w:abstractNumId w:val="${defining![1]}"/></w:num>`,
    )
    expect(numXml).toContain(
      `<w:num w:numId="10"><w:abstractNumId w:val="${linking![1]}"/></w:num>`,
    )
    expect(numXml.lastIndexOf('<w:abstractNum ')).toBeLessThan(numXml.indexOf('<w:num '))
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml).toContain(
      '<w:style w:type="numbering" w:styleId="1ai"><w:name w:val="Outline List 1"/>',
    )
    expect(stylesXml).toContain('<w:pPr><w:numPr><w:numId w:val="11"/></w:numPr></w:pPr></w:style>')

    const reparsed = await parseDocx(saved)
    const style = reparsed.listStyles!.find((s) => s.styleId === '1ai')!
    expect(style).toMatchObject({ name: 'Outline List 1', numId: '11' })
    expect(style.custom).toBeUndefined()
    expect(style.levels[1]).toMatchObject({ numFmt: 'lowerLetter', lvlText: '%2)' })
    expect(reparsed.numbering.get('10')!.listStyleId).toBe('1ai')
    expect(
      computeListMarkers(
        [
          { numId: '10', ilvl: 0 },
          { numId: '10', ilvl: 1 },
          { numId: '10', ilvl: 0 },
        ],
        reparsed.numbering,
      ),
    ).toEqual(['1)', 'a)', '2)'])
  })

  it('a package without a styles part gets one, with its relationship and content type', async () => {
    const zip = await JSZip.loadAsync(await buildDocx({ bodyXml: P('x') }))
    zip.remove('word/styles.xml')
    for (const [name, pattern] of [
      ['word/_rels/document.xml.rels', /<Relationship [^>]*Target="styles\.xml"[^>]*\/>/],
      ['[Content_Types].xml', /<Override [^>]*PartName="\/word\/styles\.xml"[^>]*\/>/],
    ] as const)
      zip.file(name, (await zip.file(name)!.async('string')).replace(pattern, ''))
    const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    const saved = await saveDocx(doc, originals(doc), {
      numbering: {
        styleNums: [
          {
            numId: '10',
            styleId: '1ai',
            define: { name: 'Outline List 1', numId: '11', levels: LEVELS },
          },
        ],
      },
    })
    const out = await JSZip.loadAsync(saved)
    expect(await out.file('word/_rels/document.xml.rels')!.async('string')).toMatch(
      /Type="[^"]*\/styles" Target="styles\.xml"|Target="styles\.xml"[^>]*Type="[^"]*\/styles"/,
    )
    expect(await out.file('[Content_Types].xml')!.async('string')).toContain(
      'PartName="/word/styles.xml"',
    )
    expect((await parseDocx(saved)).listStyles?.map((s) => s.styleId)).toEqual(['1ai'])
  })

  it('two lists of one undefined style define it once', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: P('x') }))
    const define = { name: 'Mine', numId: '11', levels: LEVELS }
    const saved = await saveDocx(doc, originals(doc), {
      numbering: {
        styleNums: [
          { numId: '10', styleId: 'Mine', define },
          { numId: '12', styleId: 'Mine', define },
        ],
      },
    })
    const zip = await JSZip.loadAsync(saved)
    const numXml = await zip.file('word/numbering.xml')!.async('string')
    expect(numXml.match(/<w:styleLink /g)).toHaveLength(1)
    expect(numXml.match(/<w:numStyleLink /g)).toHaveLength(2)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml.match(/w:styleId="Mine"/g)).toHaveLength(1)
  })
})

describe('numFmt "none" with empty lvlText', () => {
  const NONE_NUMBERING =
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'

  it('emits an explicit empty marker, not null (null re-enables counter fallbacks)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: NONE_NUMBERING,
      }),
    )
    expect(doc.numbering.get('1')!.levels[0].numFmt).toBe('none')
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 0 },
      ],
      doc.numbering,
    )
    expect(markers).toEqual(['', ''])
  })
})

describe('empty lvlText', () => {
  const EMPTY_TEXT_NUMBERING =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="nothing"/><w:lvlText w:val=""/>' +
    '<w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'

  it('a decimal level whose lvlText is empty shows nothing, not the counter fallback', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: EMPTY_TEXT_NUMBERING,
      }),
    )
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 0 },
      ],
      doc.numbering,
    )
    expect(markers).toEqual(['', ''])
  })
})

describe('w14 custom numFmt (mc:AlternateContent)', () => {
  const alternate = (choiceFmt: string, format: string) =>
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/>' +
    '<mc:AlternateContent><mc:Choice Requires="w14">' +
    `<w:numFmt w:val="${choiceFmt}" w:format="${format}"/>` +
    '</mc:Choice><mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback></mc:AlternateContent>' +
    '<w:lvlText w:val="%1."/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'

  it('parses the enumerated custom format from mc:Choice and cycles it', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: alternate('custom', 'α, β, γ, ...'),
      }),
    )
    const level = doc.numbering.get('1')!.levels[0]
    expect(level.numFmt).toBe('custom')
    expect(level.customFormat).toBe('α, β, γ, ...')
    const ref = { numId: '1', ilvl: 0 }
    expect(computeListMarkers([ref, ref, ref, ref], doc.numbering)).toEqual([
      'α.',
      'β.',
      'γ.',
      'α.',
    ])
  })

  it('falls back to mc:Fallback when the custom format is not an enumeration', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: alternate('custom', 'mystery'),
      }),
    )
    expect(doc.numbering.get('1')!.levels[0].numFmt).toBe('decimal')
    expect(computeListMarkers([{ numId: '1', ilvl: 0 }], doc.numbering)).toEqual(['1.'])
  })

  it('customEnumItems accepts comma enumerations and rejects everything else', () => {
    expect(customEnumItems('α, β, γ, ...')).toEqual(['α', 'β', 'γ'])
    expect(customEnumItems('01, 02, 03')).toEqual(['01', '02', '03'])
    expect(customEnumItems('mystery')).toBeNull()
    expect(customEnumItems('a, ')).toBeNull()
  })

  it('formatNumber cycles custom enumerations', () => {
    expect(formatNumber(2, 'custom', 'α, β, γ, ...')).toBe('β')
    expect(formatNumber(5, 'custom', 'α, β, γ, ...')).toBe('β')
    expect(formatNumber(3, 'custom')).toBe('3')
  })

  // genoffice#2121: the dialog shape dropped the enumeration, so a reused or
  // re-saved custom level came back as plain digits ("Za 1:")
  it('customLevelFromNumberingLevel carries the enumeration', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: alternate('custom', 'α, β, γ, ...'),
      }),
    )
    const custom = customLevelFromNumberingLevel(doc.numbering.get('1')!.levels[0])
    expect(custom.numFmt).toBe('custom')
    expect(custom.customFormat).toBe('α, β, γ, ...')
  })

  it('a new custom level is written as mc:AlternateContent and parses back', async () => {
    const lvl = customLevelXml(
      { numFmt: 'custom', customFormat: 'Za prvé, Za druhé, ...', lvlText: '%1:', indentLeft: 720 },
      0,
    )
    expect(lvl).not.toMatch(/<w:numFmt w:val="custom"\/>/)
    expect(lvl).toMatch(/<w:start[^>]*\/><mc:AlternateContent[^>]*><mc:Choice Requires="w14">/)
    expect(lvl).toContain('<mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback>')
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml:
          XML_DECL +
          '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          `<w:abstractNum w:abstractNumId="0">${lvl}</w:abstractNum>` +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
      }),
    )
    expect(doc.numbering.get('1')!.levels[0].customFormat).toBe('Za prvé, Za druhé, ...')
    expect(
      computeListMarkers(
        [
          { numId: '1', ilvl: 0 },
          { numId: '1', ilvl: 0 },
        ],
        doc.numbering,
      ),
    ).toEqual(['Za prvé:', 'Za druhé:'])
  })

  it('mergeLevelXml replaces the custom numFmt in place instead of adding a bare one beside it', () => {
    const existing =
      '<w:lvl w:ilvl="0"><w:start w:val="1"/>' +
      '<mc:AlternateContent><mc:Choice Requires="w14"><w:numFmt w:val="custom" w:format="α, β, γ, ..."/></mc:Choice>' +
      '<mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback></mc:AlternateContent>' +
      '<w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
    const kept = mergeLevelXml(
      existing,
      { numFmt: 'custom', customFormat: 'α, β, γ, ...', lvlText: '%1.', indentLeft: 1440 },
      0,
    )
    expect(kept.match(/<w:numFmt\b/g)).toHaveLength(2)
    expect(kept.match(/<mc:AlternateContent\b/g)).toHaveLength(1)
    expect(kept).toMatch(/<w:start[^>]*\/><mc:AlternateContent/)
    expect(kept).toContain('w:format="α, β, γ, ..."')
    expect(kept).toContain('<w:ind w:left="1440"')

    const plain = mergeLevelXml(
      existing,
      { numFmt: 'upperRoman', lvlText: '%1.', indentLeft: 720 },
      0,
    )
    expect(plain).not.toContain('AlternateContent')
    expect(plain).toContain('<w:start w:val="1"/><w:numFmt w:val="upperRoman"/>')
  })
})

describe('numbering level positive w:firstLine', () => {
  const FIRSTLINE_NUMBERING =
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1"/>' +
    '<w:pPr><w:ind w:left="432" w:firstLine="135"/></w:pPr></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'

  it('keeps the positive first-line indent (marker sits right of the text indent)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: FIRSTLINE_NUMBERING,
      }),
    )
    expect(doc.numbering.get('1')!.levels[0]).toMatchObject({
      indentLeft: 432,
      firstLine: 135,
    })
    expect(doc.numbering.get('1')!.levels[0].hanging).toBeUndefined()
  })
})

describe('markerTabAdvance (default tab after the marker)', () => {
  it('returns null when the marker fits the hanging area', () => {
    expect(markerTabAdvance(0, 300, 360)).toBeNull()
    expect(markerTabAdvance(0, 360, 360)).toBeNull()
  })

  it('a custom stop between the marker end and the hanging indent catches the tab', () => {
    // numbered heading: level left=720 hanging=720, paragraph stops 450/1350/3060,
    // one-digit marker ~160 twips -> Word starts the text at 450, not at 720
    expect(markerTabAdvance(0, 160, 720, 720, [450, 1350, 3060, 720])).toBe(450)
    // stops at or past the indent change nothing
    expect(markerTabAdvance(0, 160, 720, 720, [720, 1350])).toBeNull()
    // a stop under the marker itself is skipped
    expect(markerTabAdvance(0, 500, 720, 720, [450])).toBeNull()
  })

  it('jumps to the next default-tab-grid stop when the marker overflows', () => {
    // "NEW-1-FORMAT" at ind left=360 hanging=360: marker 0..~1476 -> stop at 2160
    expect(markerTabAdvance(0, 1476, 360)).toBe(2160)
    expect(markerTabAdvance(0, 400, 360)).toBe(720)
  })

  it('handles markers starting past the text indent (positive firstLine)', () => {
    // 47_NumberingWOverrides "B": marker at 567, width ~160 -> stop at 1440
    expect(markerTabAdvance(567, 160, 432)).toBe(1440 - 567)
  })

  it('honors a custom default tab interval', () => {
    expect(markerTabAdvance(0, 400, 360, 708)).toBe(708)
  })

  it('a zero default tab grid puts the text right after the marker, custom stops still win', () => {
    expect(markerTabAdvance(0, 400, 360, 0)).toBe(400)
    expect(markerTabAdvance(0, 400, 360, 0, [1080])).toBe(1080)
  })
})

describe('picture bullets (w:numPicBullet / w:lvlPicBulletId)', () => {
  const NS =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:v="urn:schemas-microsoft-com:vml" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
  const VML_PIC =
    '<w:numPicBullet w:numPicBulletId="0"><w:pict>' +
    '<v:shape id="_x0000_i1029" type="#_x0000_t75" style="width:11.25pt;height:11.25pt" o:bullet="t">' +
    '<v:imagedata r:id="rId1" o:title="mso6BF"/></v:shape></w:pict></w:numPicBullet>'
  const BLIP_PIC =
    '<w:numPicBullet w:numPicBulletId="1"><w:drawing><wp:inline><a:graphic><a:graphicData>' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId2"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:numPicBullet>'
  // lvlText is Word's placeholder glyph (Symbol private-use bullet)
  const lvl = (ilvl: number, picId: number) =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\uF0B7"/>` +
    `<w:lvlPicBulletId w:val="${picId}"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>` +
    '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>'
  const numberingXml =
    XML_DECL +
    `<w:numbering ${NS}>${VML_PIC}${BLIP_PIC}` +
    `<w:abstractNum w:abstractNumId="0">${lvl(0, 0)}${lvl(1, 1)}${lvl(2, 7)}</w:abstractNum>` +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'
  const numberingRels =
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>' +
    '</Relationships>'
  const li = (ilvl: number, text: string) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
  const build = (withMedia = true) =>
    buildDocx({
      bodyXml: li(0, 'vml') + li(1, 'blip') + li(2, 'undefined id'),
      numberingXml,
      extraParts: [
        {
          path: 'word/_rels/numbering.xml.rels',
          xml: numberingRels,
          contentType: 'application/vnd.openxmlformats-package.relationships+xml',
        },
      ],
      binaryParts: withMedia
        ? ['word/media/image1.png', 'word/media/image2.png'].map((path) => ({
            path,
            base64: TINY_PNG_BASE64,
            extension: 'png',
            contentType: 'image/png',
          }))
        : [],
    })
  const PNG_URL = `data:image/png;base64,${TINY_PNG_BASE64}`

  it('resolves v:imagedata and a:blip picture bullets through numbering.xml.rels', async () => {
    const doc = await parseDocx(await build())
    const levels = doc.numbering.get('1')!.levels
    expect(levels[0].picBulletId).toBe(0)
    expect(levels[0].picBulletSrc).toBe(PNG_URL)
    expect(levels[1].picBulletId).toBe(1)
    expect(levels[1].picBulletSrc).toBe(PNG_URL)
    expect(levels[2]).toMatchObject({ picBulletId: 7 })
    expect(levels[2].picBulletSrc).toBeUndefined()

    const items = [0, 1, 2].map((ilvl) => ({ numId: '1', ilvl }))
    const infos = computeListMarkerInfos(items, doc.numbering)
    expect(infos[0]).toEqual({ text: '', picBulletSrc: PNG_URL })
    expect(infos[1]).toEqual({ text: '', picBulletSrc: PNG_URL })
    // no numPicBullet for the id: a plain bullet, never the placeholder glyph
    expect(infos[2]).toEqual({ text: '\u2022' })
  })

  it('falls back to a plain bullet when the media part is missing', async () => {
    const doc = await parseDocx(await build(false))
    const levels = doc.numbering.get('1')!.levels
    expect(levels[0].picBulletId).toBe(0)
    expect(levels[0].picBulletSrc).toBeUndefined()
    expect(computeListMarkers([{ numId: '1', ilvl: 0 }], doc.numbering)).toEqual(['\u2022'])
  })

  it('saving keeps numbering.xml, its rels and the bullet media byte-identical', async () => {
    const bytes = await build()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    const blocks: SaveBlock[] = visible.map((b, i) =>
      i === 0
        ? {
            kind: 'generated',
            block: { type: 'paragraph', runs: [{ text: 'edited', bold: false }] },
          }
        : { kind: 'original', docxIndex: b.docxIndex! },
    )
    const saved = await saveDocx(doc, blocks)
    const before = await JSZip.loadAsync(bytes)
    const after = await JSZip.loadAsync(saved)
    for (const name of [
      'word/numbering.xml',
      'word/_rels/numbering.xml.rels',
      'word/media/image1.png',
      'word/media/image2.png',
    ]) {
      expect(await after.file(name)?.async('uint8array'), name).toEqual(
        await before.file(name)!.async('uint8array'),
      )
    }
    const reparsed = await parseDocx(saved)
    expect(reparsed.numbering.get('1')!.levels[0].picBulletSrc).toBe(PNG_URL)
  })
})

describe('style numPr inherited per component along basedOn', () => {
  const STYLES =
    '<w:style w:type="paragraph" w:styleId="H1"><w:name w:val="H1"/>' +
    '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="H2"><w:name w:val="H2"/><w:basedOn w:val="H1"/>' +
    '<w:pPr><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="H3"><w:name w:val="H3"/><w:basedOn w:val="H2"/>' +
    '<w:pPr><w:numPr><w:ilvl w:val="2"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="H3Off"><w:name w:val="H3Off"/><w:basedOn w:val="H2"/>' +
    '<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="H4Off"><w:name w:val="H4Off"/><w:basedOn w:val="H3Off"/>' +
    '<w:pPr><w:numPr><w:ilvl w:val="3"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="LvlOnly"><w:name w:val="LvlOnly"/>' +
    '<w:pPr><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr></w:style>'
  const p = (style: string, text: string, numPr = '') =>
    `<w:p><w:pPr><w:pStyle w:val="${style}"/>${numPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

  it('takes numId from an ancestor when the style declares only w:ilvl', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: p('H1', 'a') + p('H2', 'b') + p('H3', 'c') + p('H2', 'd'),
        extraStylesXml: STYLES,
        numberingXml: MULTILEVEL_NUMBERING,
      }),
    )
    expect(doc.styles.get('H2')?.numPr).toEqual({ numId: '1', ilvl: 1 })
    expect(doc.styles.get('H3')?.numPr).toEqual({ numId: '1', ilvl: 2 })
    const refs = doc.blocks
      .filter((b) => b.list)
      .map((b) => ({ numId: b.list!.numId, ilvl: b.list!.ilvl }))
    expect(computeListMarkers(refs, doc.numbering)).toEqual(['1.', '1.1', 'c)', '1.2'])
  })

  it('numId 0 still cancels, also for descendants that re-declare only w:ilvl', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: p('H3Off', 'a') + p('H4Off', 'b') + p('LvlOnly', 'c'),
        extraStylesXml: STYLES,
        numberingXml: MULTILEVEL_NUMBERING,
      }),
    )
    expect(doc.styles.get('H3Off')?.numPr).toBe('none')
    expect(doc.styles.get('H4Off')?.numPr).toBe('none')
    expect(doc.styles.get('LvlOnly')?.numPr).toBeUndefined()
    expect(doc.blocks.every((b) => b.list === undefined)).toBe(true)
  })

  it('a paragraph-level w:ilvl alone overrides the level of the style numbering', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          p('H1', 'a', '<w:numPr><w:ilvl w:val="1"/></w:numPr>') +
          p('H2', 'b', '<w:numPr><w:numId w:val="0"/></w:numPr>'),
        extraStylesXml: STYLES,
        numberingXml: MULTILEVEL_NUMBERING,
      }),
    )
    expect(doc.blocks[0].list).toMatchObject({ numId: '1', ilvl: 1 })
    expect(doc.blocks[1].list).toBeUndefined()
  })
})

describe('tracked-deleted list paragraphs do not consume a number', () => {
  it('a deleted item previews the next number and leaves the counters untouched', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: MULTILEVEL_NUMBERING,
      }),
    )
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 0, deleted: true },
        { numId: '1', ilvl: 0, deleted: true },
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 1, deleted: true },
        { numId: '1', ilvl: 1 },
        { numId: '2', ilvl: 0, deleted: true },
        { numId: '2', ilvl: 0 },
        { numId: '2', ilvl: 0 },
      ],
      doc.numbering,
    )
    expect(markers).toEqual(['1.', '2.', '2.', '2.', '2.1', '2.1', '1.', '1.', '2.'])
  })

  it('parses the deleted paragraph mark of a list item', async () => {
    const li = (text: string, del = '') =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${del}</w:pPr>` +
      `<w:r><w:t>${text}</w:t></w:r></w:p>`
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          li('a') +
          li('b', '<w:rPr><w:del w:id="1" w:author="r" w:date="2020-01-01T00:00:00Z"/></w:rPr>'),
        numberingXml: MULTILEVEL_NUMBERING,
      }),
    )
    expect(doc.blocks[0].paraMarkDel).toBeUndefined()
    expect(doc.blocks[1].paraMarkDel).toMatchObject({ author: 'r' })
    expect(doc.blocks[1].list).toMatchObject({ numId: '1', ilvl: 0 })
  })
})

describe('w:isLgl', () => {
  const LEGAL_NUMBERING =
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="Article %1."/>' +
    '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimalZero"/><w:isLgl/><w:lvlText w:val="Section %1.%2"/></w:lvl>' +
    '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:isLgl/><w:lvlJc w:val="right"/><w:lvlText w:val="%1.%2.%3"/>' +
    '<w:rPr><w:color w:val="ffc000"/><w:sz w:val="52"/></w:rPr></w:lvl>' +
    '<w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:isLgl w:val="0"/><w:lvlText w:val="%4"/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'

  it('parses the flag, right alignment, marker color and explicit zero indents', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: LEGAL_NUMBERING,
      }),
    )
    const levels = doc.numbering.get('1')!.levels
    expect(levels[0].isLgl).toBeUndefined()
    expect(levels[0].indentLeft).toBe(0)
    expect(levels[0].firstLine).toBe(0)
    expect(levels[0].hanging).toBeUndefined()
    expect(levels[1].isLgl).toBe(true)
    expect(levels[2]).toMatchObject({
      isLgl: true,
      lvlJc: 'right',
      color: 'FFC000',
      szHalfPoints: 52,
    })
    expect(levels[3].isLgl).toBeUndefined()
  })

  it('renders every placeholder of a legal level as decimal, Arabic formats untouched', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        numberingXml: LEGAL_NUMBERING,
      }),
    )
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 1 },
        { numId: '1', ilvl: 2 },
        { numId: '1', ilvl: 2 },
        { numId: '1', ilvl: 3 },
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 1 },
      ],
      doc.numbering,
    )
    expect(markers).toEqual([
      'Article I.',
      'Section 1.01',
      '1.01.1',
      '1.01.2',
      'i',
      'Article II.',
      'Section 2.01',
    ])
  })
})

describe('enclosed, full-width and spelled-out number formats', () => {
  it('maps the enclosed decimal family onto the Unicode enclosed alphanumerics', () => {
    expect(formatNumber(1, 'decimalEnclosedFullstop')).toBe('⒈')
    expect(formatNumber(20, 'decimalEnclosedFullstop')).toBe('⒛')
    expect(formatNumber(21, 'decimalEnclosedFullstop')).toBe('21')
    expect(formatNumber(3, 'decimalEnclosedParen')).toBe('⑶')
    expect(formatNumber(7, 'decimalEnclosedCircleChinese')).toBe('⑦')
    expect(formatNumber(2, 'ideographEnclosedCircle')).toBe('㊁')
    expect(formatNumber(11, 'ideographEnclosedCircle')).toBe('11')
  })

  it('transliterates digit-wise formats and counts CJK legal / Korean forms', () => {
    expect(formatNumber(12, 'decimalFullWidth')).toBe('１２')
    expect(formatNumber(105, 'ideographDigital')).toBe('\u4e00\u3007\u4e94')
    expect(formatNumber(20, 'koreanDigital')).toBe('이영')
    expect(formatNumber(12, 'chineseLegalSimplified')).toBe('\u58f9\u62fe\u8d30')
    expect(formatNumber(12, 'ideographLegalTraditional')).toBe('\u58f9\u62fe\u8cb3')
    expect(formatNumber(12, 'taiwaneseCounting')).toBe('\u5341\u4e8c')
    expect(formatNumber(3, 'ideographTraditional')).toBe('\u4e19')
    expect(formatNumber(12, 'ideographZodiac')).toBe('\u4ea5')
    expect(formatNumber(5, 'aiueo')).toBe('ｵ')
    expect(formatNumber(2, 'irohaFullWidth')).toBe('ロ')
    expect(formatNumber(3, 'ganada')).toBe('다')
    expect(formatNumber(2, 'russianUpper')).toBe('Б')
    expect(formatNumber(9, 'thaiNumbers')).toBe('๙')
  })

  it('spells English ordinals and cardinals', () => {
    expect(
      [1, 2, 3, 4, 5, 8, 9, 12, 20, 21, 100, 112].map((n) => formatNumber(n, 'ordinalText')),
    ).toEqual([
      'First',
      'Second',
      'Third',
      'Fourth',
      'Fifth',
      'Eighth',
      'Ninth',
      'Twelfth',
      'Twentieth',
      'Twenty-First',
      'One Hundredth',
      'One Hundred Twelfth',
    ])
    expect(formatNumber(1234, 'cardinalText')).toBe('One Thousand Two Hundred Thirty-Four')
    expect([1, 2, 3, 11, 12, 13, 22, 101].map((n) => formatNumber(n, 'ordinal'))).toEqual([
      '1st',
      '2nd',
      '3rd',
      '11th',
      '12th',
      '13th',
      '22nd',
      '101st',
    ])
    expect(formatNumber(4, 'numberInDash')).toBe('- 4 -')
  })
})

describe('bullets declared in an ordinary text font', () => {
  it('keep the literal glyph and report the font; without a font the glyph substitutes', async () => {
    const numbering =
      XML_DECL +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl>' +
      '<w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '</w:numbering>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>', numberingXml: numbering }),
    )
    const infos = computeListMarkerInfos(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 1 },
        { numId: '1', ilvl: 2 },
      ],
      doc.numbering,
    )
    expect(infos[0]).toEqual({ text: 'o', font: 'Courier New' })
    expect(infos[1]).toEqual({ text: '◦' })
    // a private-use glyph is not literal text: the common bullet substitute
    expect(infos[2]).toEqual({ text: '\u2022' })
  })
})

describe('markerTabAdvance with custom tab stops', () => {
  it('lands on the first custom stop past the marker; the default grid only beyond the last one', () => {
    // marker at 720 (left 360 + firstLine 360), 200 wide, custom stop 4320
    expect(markerTabAdvance(720, 200, 360, 720, [4320])).toBe(3600)
    // marker at 360, custom stop 1080 replaces the 720 default
    expect(markerTabAdvance(360, 200, 360, 720, [1080])).toBe(720)
    // the hanging edge still wins when the marker fits
    expect(markerTabAdvance(360, 200, 720, 720, [4320])).toBeNull()
    // every custom stop is behind the marker: next default past both
    expect(markerTabAdvance(0, 1500, 0, 720, [720])).toBe(2160)
    expect(markerTabAdvance(0, 1500, 0, 720, [])).toBe(2160)
  })
})

describe('hostile numbering values', () => {
  it('bounds huge and non-finite values instead of hanging', () => {
    const start = Date.now()
    for (const fmt of [
      'upperRoman',
      'lowerRoman',
      'upperLetter',
      'lowerLetter',
      'upperGreek',
      'decimal',
    ]) {
      for (const v of [1e9, Infinity, -Infinity, NaN]) {
        const out = formatNumber(v, fmt)
        expect(typeof out).toBe('string')
        expect(out.length).toBeLessThan(100_000)
      }
    }
    expect(Date.now() - start).toBeLessThan(10000)
    expect(formatNumber(4, 'upperRoman')).toBe('IV')
    expect(formatNumber(27, 'upperLetter')).toBe('AA')
  })
})
