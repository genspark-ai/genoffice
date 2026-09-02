import { describe, expect, it } from 'vitest'
import { mergePPrFormat, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * Character-unit indents (w:firstLineChars / w:hangingChars / w:leftChars /
 * w:rightChars). Every expectation below is a Word for Mac measurement
 * (probe 2026-09-02, docs-fidelity word-truths.md): first-line/hanging
 * units follow the first text run, left/right units follow the Normal style,
 * the character form supersedes the twips twin, zero means "absolute".
 */

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

/** docDefaults 10.5pt (the usual Chinese Normal), plus styles carrying character indents */
function stylesXml(normalRPr = ''): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>' +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>${normalRPr}</w:style>` +
    '<w:style w:type="paragraph" w:styleId="Body16"><w:name w:val="Body16"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:ind w:firstLineChars="200"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Body16Left"><w:name w:val="Body16Left"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:ind w:leftChars="200"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="BodyNoSz"><w:name w:val="BodyNoSz"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:ind w:firstLineChars="200"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>' +
    '</w:styles>'
  )
}

function run(text: string, sz?: number, extraRPr = ''): string {
  const rPr = (sz ? `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` : '') + extraRPr
  return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
}

function para(pPr: string, runs: string): string {
  return `<w:p><w:pPr>${pPr}</w:pPr>${runs}</w:p>`
}

/** parsed paragraph formats, without the `charIndents` marker (asserted separately) */
async function formats(bodyXml: string, styles = stylesXml()) {
  const doc = await parseDocx(await buildDocx({ bodyXml, stylesXml: styles }))
  return doc.blocks.map((b) => {
    if (!b.format) return b.format
    const { charIndents: _chars, ...rest } = b.format
    return Object.keys(rest).length > 0 ? rest : undefined
  })
}

describe('character-unit indents: first line and hanging follow the first text run', () => {
  it('firstLineChars scales with the run size (2 chars of 16pt = 640 twips, of 12pt = 480)', async () => {
    const f = await formats(
      para('<w:ind w:firstLineChars="200"/>', run('body', 32)) +
        para('<w:ind w:firstLineChars="200"/>', run('body', 24)) +
        // Latin text: the unit is the font size, not a glyph width
        para('<w:ind w:firstLineChars="200"/>', run('Latin text', 24)) +
        // no run size: docDefaults 10.5pt
        para('<w:ind w:firstLineChars="200"/>', run('body')),
    )
    expect(f[0]).toEqual({ indentFirstLine: 640 })
    expect(f[1]).toEqual({ indentFirstLine: 480 })
    expect(f[2]).toEqual({ indentFirstLine: 480 })
    expect(f[3]).toEqual({ indentFirstLine: 420 })
  })

  it('the character form supersedes a stale twips twin', async () => {
    const f = await formats(
      para('<w:ind w:firstLineChars="200" w:firstLine="100"/>', run('body', 24)),
    )
    expect(f[0]).toEqual({ indentFirstLine: 480 })
  })

  it('uses the first text run, not the largest run or the paragraph mark', async () => {
    const f = await formats(
      para('<w:ind w:firstLineChars="200"/>', run('small', 20) + run('LARGE TEXT', 40)) +
        para('<w:ind w:firstLineChars="200"/>', run('LARGE', 40) + run('small text', 20)) +
        // 20pt paragraph mark, 10pt runs: the mark does not count
        para('<w:ind w:firstLineChars="200"/><w:rPr><w:sz w:val="40"/></w:rPr>', run('body', 20)) +
        // an empty first run does not count either
        para('<w:ind w:firstLineChars="200"/>', run('', 20) + run('body', 40)),
    )
    expect(f[0]).toEqual({ indentFirstLine: 400 })
    expect(f[1]).toEqual({ indentFirstLine: 800 })
    expect(f[2]).toEqual({ indentFirstLine: 400 })
    expect(f[3]).toEqual({ indentFirstLine: 800 })
  })

  it('adds the first run letter spacing to the unit (12pt + 5pt spacing -> 17pt characters)', async () => {
    const f = await formats(
      para('<w:ind w:firstLineChars="200"/>', run('body', 24, '<w:spacing w:val="100"/>')),
    )
    expect(f[0]).toEqual({ indentFirstLine: 680 })
  })

  it('a run-less paragraph sizes the unit by its paragraph mark', async () => {
    const f = await formats(
      para('<w:ind w:firstLineChars="200"/><w:rPr><w:sz w:val="32"/></w:rPr>', ''),
    )
    expect(f[0]).toMatchObject({ indentFirstLine: 640 })
  })
})

describe('character-unit indents: left and right follow the Normal style', () => {
  it('leftChars uses the default paragraph size even for 12pt text (2 chars = 420 twips)', async () => {
    const f = await formats(
      para('<w:ind w:leftChars="200"/>', run('body', 24)) +
        // the twips twin is ignored
        para('<w:ind w:leftChars="200" w:left="1440"/>', run('body', 24)) +
        // strict attribute names
        para('<w:ind w:startChars="200"/>', run('body', 24)) +
        // negative left is legal
        para('<w:ind w:leftChars="-200"/>', run('body', 24)) +
        // run letter spacing does not affect the Normal unit
        para('<w:ind w:leftChars="200"/>', run('body', 24, '<w:spacing w:val="100"/>')),
    )
    expect(f[0]).toEqual({ indentLeft: 420 })
    expect(f[1]).toEqual({ indentLeft: 420 })
    expect(f[2]).toEqual({ indentLeft: 420 })
    expect(f[3]).toEqual({ indentLeft: -420 })
    expect(f[4]).toEqual({ indentLeft: 420 })
  })

  it('a Normal style size outranks docDefaults for the left unit (14pt Normal -> 560)', async () => {
    const f = await formats(
      para('<w:ind w:leftChars="200"/>', run('body', 24)) +
        para('<w:ind w:firstLineChars="200"/>', run('body')),
      stylesXml('<w:rPr><w:sz w:val="28"/></w:rPr>'),
    )
    expect(f[0]).toEqual({ indentLeft: 560 })
    // a run without its own size inherits Normal's 14pt for the first-line unit too
    expect(f[1]).toEqual({ indentFirstLine: 560 })
  })

  it('rightChars follows Normal and beats the twips twin', async () => {
    const f = await formats(
      para('<w:jc w:val="right"/><w:ind w:rightChars="200" w:right="1440"/>', run('body', 24)),
    )
    expect(f[0]).toEqual({ align: 'right', indentRight: 420 })
  })
})

describe('character-unit indents: hanging semantics and mixed units', () => {
  it('leftChars is the first-line position; body lines sit the hanging further in', async () => {
    const f = await formats(
      // leftChars=4 (Normal 10.5pt -> 840) + hangingChars=2 (12pt -> 480): first 840, body 1320
      para('<w:ind w:leftChars="400" w:hangingChars="200"/>', run('body', 24)) +
        // hangingChars beats a firstLineChars next to it
        para(
          '<w:ind w:leftChars="400" w:hangingChars="200" w:firstLineChars="300"/>',
          run('body', 24),
        ) +
        // twips hanging next to leftChars: same character-mode semantics
        para('<w:ind w:leftChars="400" w:hanging="240"/>', run('body', 24)) +
        // hangingChars alone: first line at the margin, body 2 chars in
        para('<w:ind w:hangingChars="200"/>', run('body', 24)) +
        // hangingChars drops a twips w:left (Word ignores it)
        para('<w:ind w:left="720" w:hangingChars="200"/>', run('body', 24)) +
        // ... even next to an explicit leftChars="0"
        para('<w:ind w:left="720" w:leftChars="0" w:hangingChars="200"/>', run('body', 24)) +
        // leftChars + firstLineChars: plain first-line indent from the character left
        para('<w:ind w:leftChars="400" w:firstLineChars="200"/>', run('body', 24)),
    )
    expect(f[0]).toEqual({ indentLeft: 1320, indentFirstLine: -480 })
    expect(f[1]).toEqual({ indentLeft: 1320, indentFirstLine: -480 })
    expect(f[2]).toEqual({ indentLeft: 1080, indentFirstLine: -240 })
    expect(f[3]).toEqual({ indentLeft: 480, indentFirstLine: -480 })
    expect(f[4]).toEqual({ indentLeft: 480, indentFirstLine: -480 })
    expect(f[5]).toEqual({ indentLeft: 480, indentFirstLine: -480 })
    expect(f[6]).toEqual({ indentLeft: 840, indentFirstLine: 480 })
  })

  it('firstLineChars alone keeps a twips w:left as the body indent and beats a twips hanging', async () => {
    const f = await formats(
      para('<w:ind w:left="720" w:firstLineChars="200"/>', run('body', 24)) +
        para('<w:ind w:left="720" w:firstLineChars="200" w:firstLine="480"/>', run('body', 24)) +
        para('<w:ind w:left="720" w:hanging="240" w:firstLineChars="200"/>', run('body', 24)),
    )
    expect(f[0]).toEqual({ indentLeft: 720, indentFirstLine: 480 })
    expect(f[1]).toEqual({ indentLeft: 720, indentFirstLine: 480 })
    expect(f[2]).toEqual({ indentLeft: 720, indentFirstLine: 480 })
  })

  it('a zero character value means "absolute": the twips attribute stays in force', async () => {
    const f = await formats(
      para('<w:ind w:firstLineChars="0" w:firstLine="480"/>', run('body', 24)) +
        para('<w:ind w:leftChars="0" w:left="720"/>', run('body', 24)) +
        para('<w:ind w:left="720" w:hanging="240" w:hangingChars="0"/>', run('body', 24)) +
        // twips-only paragraphs are untouched
        para('<w:ind w:hanging="240"/>', run('body', 24)),
    )
    expect(f[0]).toEqual({ indentFirstLine: 480 })
    expect(f[1]).toEqual({ indentLeft: 720 })
    expect(f[2]).toEqual({ indentLeft: 720, indentFirstLine: -240 })
    expect(f[3]).toEqual({ indentFirstLine: -240 })
  })
})

describe('character-unit indents: document grid and styles', () => {
  it('a linesAndChars grid widens both units by charSpace/4096 pt (snapToGrid off included)', async () => {
    const grid = '<w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="16384"/>'
    const f = await formats(
      para('<w:ind w:firstLineChars="200"/>', run('body', 24)) +
        para('<w:ind w:leftChars="200"/>', run('body', 24)) +
        para('<w:snapToGrid w:val="0"/><w:ind w:firstLineChars="200"/>', run('body', 24)) +
        `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${grid}</w:sectPr></w:pPr></w:p>` +
        // second section, plain lines grid: no delta
        para('<w:ind w:firstLineChars="200"/>', run('body', 24)),
    )
    // 12pt + 4pt = 16pt characters -> 640; Normal 10.5 + 4 = 14.5pt -> 580
    expect(f[0]).toEqual({ indentFirstLine: 640 })
    expect(f[1]).toEqual({ indentLeft: 580 })
    expect(f[2]).toMatchObject({ indentFirstLine: 640 })
    expect(f[4]).toEqual({ indentFirstLine: 480 })
  })

  it('a style firstLineChars resolves per paragraph against the paragraph text', async () => {
    const f = await formats(
      // style size 16pt, runs unsized: 640
      para('<w:pStyle w:val="Body16"/>', run('body')) +
        // runs 12pt directly: 480
        para('<w:pStyle w:val="Body16"/>', run('body', 24)) +
        // no size anywhere: docDefaults 10.5pt -> 420
        para('<w:pStyle w:val="BodyNoSz"/>', run('body')) +
        // a style leftChars still follows Normal, not the style's own 16pt
        para('<w:pStyle w:val="Body16Left"/>', run('body')) +
        // style firstLineChars + direct leftChars combine
        para('<w:pStyle w:val="Body16"/><w:ind w:leftChars="200"/>', run('body')) +
        // a direct twips special indent does NOT cancel the style's character one
        // (Word: separate properties; the character form keeps superseding)
        para('<w:pStyle w:val="Body16"/><w:ind w:firstLine="0"/>', run('body')) +
        para('<w:pStyle w:val="Body16"/><w:ind w:hanging="240"/>', run('body')) +
        // only a *Chars attribute does — the explicit zero Word writes for pt indents
        para(
          '<w:pStyle w:val="Body16"/><w:ind w:firstLineChars="0" w:firstLine="420"/>',
          run('body'),
        ) +
        para('<w:pStyle w:val="Body16"/><w:ind w:hangingChars="100"/>', run('body', 24)),
    )
    expect(f[0]).toEqual({ indentFirstLine: 640 })
    expect(f[1]).toEqual({ indentFirstLine: 480 })
    expect(f[2]).toEqual({ indentFirstLine: 420 })
    expect(f[3]).toEqual({ indentLeft: 420 })
    expect(f[4]).toEqual({ indentLeft: 420, indentFirstLine: 640 })
    expect(f[5]).toEqual({ indentFirstLine: 640 })
    expect(f[6]).toEqual({ indentFirstLine: 640 })
    expect(f[7]).toEqual({ indentFirstLine: 420 })
    expect(f[8]).toEqual({ indentLeft: 240, indentFirstLine: -240 })
  })

  it('a direct twips w:left / w:hanging leaves a style character indent in force', async () => {
    // Word for Mac probe 2026-09-02: PC = firstLineChars 200, PL = leftChars 200
    const f = await formats(
      // 12pt text: chars 24pt first line; twips left 36pt body -> 60 / 36
      para('<w:pStyle w:val="Body16"/><w:ind w:left="720" w:hanging="240"/>', run('body', 24)) +
        para('<w:pStyle w:val="Body16"/><w:ind w:left="720"/>', run('body', 24)) +
        // style leftChars: direct twips w:left is ignored (character-mode left edge)
        para('<w:pStyle w:val="Body16Left"/><w:ind w:left="1440"/>', run('body', 24)) +
        para(
          '<w:pStyle w:val="Body16Left"/><w:ind w:leftChars="0" w:left="1440"/>',
          run('body', 24),
        ),
    )
    expect(f[0]).toEqual({ indentLeft: 720, indentFirstLine: 480 })
    expect(f[1]).toEqual({ indentLeft: 720, indentFirstLine: 480 })
    expect(f[2]).toEqual({ indentLeft: 420 })
    expect(f[3]).toEqual({ indentLeft: 1440 })
  })

  it('a derived style keeps the parent character indent unless it sets the *Chars attribute', async () => {
    // Word for Mac probe 2026-09-02, 12pt text, Normal 10.5pt
    const chain =
      '<w:style w:type="paragraph" w:styleId="PC"><w:name w:val="PC"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:firstLineChars="200"/></w:pPr></w:style>' +
      // twips-only child: parent chars survive (24pt, not 21)
      '<w:style w:type="paragraph" w:styleId="CT"><w:name w:val="CT"/><w:basedOn w:val="PC"/><w:pPr><w:ind w:firstLine="420"/></w:pPr></w:style>' +
      // explicit zero (what Word writes for a pt indent): absolute, 21pt
      '<w:style w:type="paragraph" w:styleId="CZ"><w:name w:val="CZ"/><w:basedOn w:val="PC"/><w:pPr><w:ind w:firstLineChars="0" w:firstLine="420"/></w:pPr></w:style>' +
      // twips hanging child: the parent's firstLineChars still wins (60 / 36)
      '<w:style w:type="paragraph" w:styleId="CH"><w:name w:val="CH"/><w:basedOn w:val="PC"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="CC"><w:name w:val="CC"/><w:basedOn w:val="PC"/><w:pPr><w:ind w:firstLineChars="300"/></w:pPr></w:style>' +
      // two levels down, through a twips-only intermediate style
      '<w:style w:type="paragraph" w:styleId="GC"><w:name w:val="GC"/><w:basedOn w:val="CT"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="PL"><w:name w:val="PL"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:leftChars="200"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="CL"><w:name w:val="CL"/><w:basedOn w:val="PL"/><w:pPr><w:ind w:left="1440"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="CLZ"><w:name w:val="CLZ"/><w:basedOn w:val="PL"/><w:pPr><w:ind w:leftChars="0" w:left="1440"/></w:pPr></w:style>'
    const f = await formats(
      para('<w:pStyle w:val="CT"/>', run('body', 24)) +
        para('<w:pStyle w:val="CZ"/>', run('body', 24)) +
        para('<w:pStyle w:val="CH"/>', run('body', 24)) +
        para('<w:pStyle w:val="CC"/>', run('body', 24)) +
        para('<w:pStyle w:val="GC"/>', run('body', 24)) +
        para('<w:pStyle w:val="CL"/>', run('body', 24)) +
        para('<w:pStyle w:val="CLZ"/>', run('body', 24)),
      stylesXml().replace('</w:styles>', chain + '</w:styles>'),
    )
    expect(f[0]).toEqual({ indentFirstLine: 480 })
    // the style's twips indent renders through the style CSS, nothing to fold
    expect(f[1]).toBeUndefined()
    expect(f[2]).toEqual({ indentFirstLine: 480 })
    expect(f[3]).toEqual({ indentFirstLine: 720 })
    expect(f[4]).toEqual({ indentFirstLine: 480 })
    expect(f[5]).toEqual({ indentLeft: 420 })
    expect(f[6]).toBeUndefined()
  })

  it('a Normal-style firstLineChars reaches unstyled paragraphs but not list items', async () => {
    const styles = stylesXml().replace(
      '<w:name w:val="Normal"/>',
      '<w:name w:val="Normal"/><w:pPr><w:ind w:firstLineChars="200"/></w:pPr>',
    )
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          para('', run('body', 24)) +
          para('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>', run('item', 24)),
        stylesXml: styles,
        withNumbering: true,
      }),
    )
    expect(doc.blocks[0].format).toEqual({
      indentFirstLine: 480,
      charIndents: { firstLine: 200 },
    })
    expect(doc.blocks[1].type).toBe('listItem')
    expect(doc.blocks[1].format).toBeUndefined()
  })

  it('table cell paragraphs resolve their character indents too', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
          para('<w:ind w:firstLineChars="200"/>', run('cell', 24)) +
          '</w:tc></w:tr></w:tbl>',
        stylesXml: stylesXml(),
      }),
    )
    const table = doc.blocks[0]
    expect(table.type).toBe('table')
    expect(table.table!.rows[0][0].richParas![0].indentFirstLine).toBe(480)
  })

  it('the section grid pitch reaches table-cell and textbox paragraphs too', async () => {
    // Word for Mac probe 2026-09-02: under linesAndChars charSpace=16384 a
    // 12pt firstLineChars="200" lays at 32pt in the body, in a cell and in a textbox
    const cellPara = para('<w:ind w:firstLineChars="200"/>', run('cell', 24))
    const tbl =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
      cellPara +
      '</w:tc></w:tr></w:tbl>'
    const txbx =
      '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="3000000" cy="600000"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Box 1"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/>' +
      '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>' +
      '<wps:txbx><w:txbxContent>' +
      para('<w:ind w:firstLineChars="200"/>', run('box', 24)) +
      '</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>' +
      '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'
    const gridSect =
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '<w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="16384"/></w:sectPr></w:pPr></w:p>'
    const doc = await parseDocx(
      await buildDocx({
        // section 1 carries the character grid, section 2 (the default sectPr) does not
        bodyXml: tbl + txbx + gridSect + tbl + txbx,
        stylesXml: stylesXml(),
      }),
    )
    const cellIndent = (b: (typeof doc.blocks)[number]) =>
      b.table!.rows[0][0].richParas![0].indentFirstLine
    const boxIndent = (b: (typeof doc.blocks)[number]) => b.textboxes![0].paras[0].indentFirstLine
    expect(cellIndent(doc.blocks[0])).toBe(640)
    expect(boxIndent(doc.blocks[1])).toBe(640)
    expect(cellIndent(doc.blocks[3])).toBe(480)
    expect(boxIndent(doc.blocks[4])).toBe(480)
  })

  it('a Normal-style firstLineChars reaches pPr-less textbox paragraphs', async () => {
    const styles = stylesXml().replace(
      '<w:name w:val="Normal"/>',
      '<w:name w:val="Normal"/><w:pPr><w:ind w:firstLineChars="200"/></w:pPr>',
    )
    const txbx =
      '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="3000000" cy="600000"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Box 1"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/>' +
      '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>' +
      // no w:pPr at all on the box paragraph
      `<wps:txbx><w:txbxContent><w:p>${run('box', 24)}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
      '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: txbx, stylesXml: styles }))
    expect(doc.blocks[0].textboxes![0].paras[0].indentFirstLine).toBe(480)
  })
})

describe('character-unit indents: saving', () => {
  it('an unrelated format edit keeps the raw w:ind bytes when the indent model is unchanged', async () => {
    const bytes = await buildDocx({
      bodyXml: para('<w:ind w:firstLineChars="200"/>', run('body', 24)),
      stylesXml: stylesXml(),
    })
    const doc = await parseDocx(bytes)
    const block = doc.blocks[0]
    // the marker tells the save path which character indents an edit has to cancel
    expect(block.format).toEqual({ indentFirstLine: 480, charIndents: { firstLine: 200 } })
    // the editor centers the paragraph: only w:jc is rebuilt, w:ind stays character-based
    const merged = mergePPrFormat(
      block.rawPPr!,
      { indentFirstLine: 480, align: 'center' },
      block.format,
    )
    expect(merged).toContain('<w:ind w:firstLineChars="200"/>')
    expect(merged).toContain('<w:jc w:val="center"/>')
    expect(merged).not.toContain('w:firstLine=')
    // a real indent edit rebuilds w:ind in twips and cancels the character unit
    // the way Word writes pt indents, so the user's value wins on reload
    const edited = mergePPrFormat(block.rawPPr!, { indentFirstLine: 720 }, block.format)
    expect(edited).toContain('<w:ind w:firstLine="720" w:firstLineChars="0"/>')
  })

  it('an indent edit under a style character indent writes the cancel Word writes', async () => {
    // Word for Mac probe 2026-09-02: a direct twips w:firstLine does not beat a style
    // firstLineChars; only `w:firstLineChars="0"` does. Without it the edit is
    // superseded by the style on reload — in Word and in this parser.
    const styles = stylesXml().replace(
      '<w:name w:val="Normal"/>',
      '<w:name w:val="Normal"/><w:pPr><w:ind w:firstLineChars="200"/></w:pPr>',
    )
    const bytes = await buildDocx({
      bodyXml:
        // style indent only, raw pPr without w:ind
        para('<w:jc w:val="both"/>', run('body', 24)) +
        // style indent, direct twips left
        para('<w:ind w:left="720"/>', run('body', 24)) +
        // direct character left + hanging
        para('<w:ind w:leftChars="400" w:hangingChars="200"/>', run('body', 24)),
      stylesXml: styles,
    })
    const doc = await parseDocx(bytes)
    const [styled, withLeft, hanging] = doc.blocks
    expect(styled.format).toEqual({
      align: 'justify',
      indentFirstLine: 480,
      charIndents: { firstLine: 200 },
    })

    // first-line indent changed: w:ind is added with the cancel, the rest of the pPr kept
    const changed = mergePPrFormat(
      styled.rawPPr!,
      { align: 'justify', indentFirstLine: 720 },
      styled.format,
    )
    expect(changed).toBe(
      '<w:pPr><w:ind w:firstLine="720" w:firstLineChars="0"/><w:jc w:val="both"/></w:pPr>',
    )
    // first-line indent removed: nothing but the cancel is left to write
    const removed = mergePPrFormat(styled.rawPPr!, { align: 'justify' }, styled.format)
    expect(removed).toBe('<w:pPr><w:ind w:firstLineChars="0"/><w:jc w:val="both"/></w:pPr>')
    // unrelated edit: no cancel, the paragraph stays character-indented
    const centered = mergePPrFormat(
      styled.rawPPr!,
      { indentFirstLine: 480, align: 'center' },
      styled.format,
    )
    expect(centered).toBe('<w:pPr><w:jc w:val="center"/></w:pPr>')
    // pPr-less paragraph (rawPPr ''): the rebuilt pPr carries the cancel too
    expect(mergePPrFormat('', { indentFirstLine: 720 }, styled.format)).toBe(
      '<w:pPr><w:ind w:firstLine="720" w:firstLineChars="0"/></w:pPr>',
    )

    expect(withLeft.format).toEqual({
      indentLeft: 720,
      indentFirstLine: 480,
      charIndents: { firstLine: 200 },
    })
    expect(
      mergePPrFormat(withLeft.rawPPr!, { indentLeft: 1440, indentFirstLine: 480 }, withLeft.format),
    ).toBe('<w:pPr><w:ind w:left="1440" w:firstLine="480" w:firstLineChars="0"/></w:pPr>')

    expect(hanging.format).toEqual({
      indentLeft: 1320,
      indentFirstLine: -480,
      charIndents: { left: 400, hanging: 200 },
    })
    expect(mergePPrFormat(hanging.rawPPr!, { indentLeft: 720 }, hanging.format)).toBe(
      '<w:pPr><w:ind w:left="720" w:leftChars="0" w:hangingChars="0"/></w:pPr>',
    )

    // round trip: the saved paragraphs re-parse to the edited values, not the style's
    const saved = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p>${changed}${run('body', 24)}</w:p>` +
          `<w:p>${removed}${run('body', 24)}</w:p>` +
          `<w:p>${mergePPrFormat(hanging.rawPPr!, { indentLeft: 720 }, hanging.format)}${run('body', 24)}</w:p>`,
        stylesXml: styles,
      }),
    )
    expect(saved.blocks[0].format).toEqual({ indentFirstLine: 720, align: 'justify' })
    expect(saved.blocks[1].format).toEqual({ align: 'justify' })
    expect(saved.blocks[2].format).toEqual({ indentLeft: 720 })
  })
})
