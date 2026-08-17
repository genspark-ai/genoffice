import { describe, expect, it } from 'vitest'
import {
  computeListMarkers,
  customEnumItems,
  formatNumber,
  markerTabAdvance,
  parseDocx,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

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
})
