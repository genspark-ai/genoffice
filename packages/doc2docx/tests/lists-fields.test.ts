import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  convertLegacyDocToDocx,
  openLegacyDoc,
  parseFields,
  parseListTables,
  readLegacyDocBody,
} from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  encodeListParagraphPapx,
  encodePapxFkp,
  encodeSprm,
  encodePlfLst,
  type SyntheticLevelOverride,
  type SyntheticListDefinition,
  type SyntheticPapxEntry,
} from './helpers/synthetic-doc'

function fixtureBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(fileURLToPath(new URL('./fixtures/' + name, import.meta.url))))
}

async function zipText(docx: Uint8Array, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(docx)
  const file = zip.file(path)
  return file ? file.async('string') : undefined
}

/** build a .doc whose body text marks list paragraphs via their PAPX sprms */
function syntheticListDoc(options: {
  text: string
  paragraphs: Array<{ startCp: number; ilfo: number; ilvl?: number }>
  lists: SyntheticListDefinition[]
  listOverrides?: number[]
  listLevelOverrides?: Array<SyntheticLevelOverride[] | undefined>
  fonts?: string[]
}): Uint8Array {
  const entries: SyntheticPapxEntry[] = options.paragraphs.map((paragraph) => ({
    fcStart: SYNTHETIC_TEXT_OFFSET + paragraph.startCp * 2,
    fcEnd: 0,
    istd: 0,
    grpprl: encodeListParagraphPapx(paragraph.ilfo, paragraph.ilvl ?? 0),
  }))
  entries.sort((left, right) => left.fcStart - right.fcStart)
  entries.forEach((entry, index) => {
    entry.fcEnd =
      index + 1 < entries.length
        ? entries[index + 1]!.fcStart
        : SYNTHETIC_TEXT_OFFSET + options.text.length * 2
  })
  return buildSyntheticDoc({
    pieces: [{ text: options.text }],
    papxFkpPages: [encodePapxFkp(entries)],
    lists: options.lists,
    ...(options.listOverrides ? { listOverrides: options.listOverrides } : {}),
    ...(options.listLevelOverrides ? { listLevelOverrides: options.listLevelOverrides } : {}),
    ...(options.fonts ? { fonts: options.fonts } : {}),
  })
}

describe('lists-link.doc: numbering + hyperlink reconstruction', () => {
  const bytes = fixtureBytes('lists-link.doc')

  it('reads the list tables and tags the list paragraphs', () => {
    const doc = openLegacyDoc(bytes)
    const lists = parseListTables(doc.fib, doc.table)
    expect(lists.warnings).toEqual([])
    expect(lists.overrides.size).toBe(7)
    expect(lists.overrides.get(1)).toBe(1)
    expect(lists.overrides.get(4)).toBe(4)
    const bullet = lists.lists.get(1)!
    expect(bullet.levels).toHaveLength(9)
    expect(bullet.levels[0]).toMatchObject({
      numberFormat: 23,
      startAt: 1,
      indentLeftTwips: 360,
      indentFirstLineTwips: -360,
    })
    const numbered = lists.lists.get(4)!
    expect(numbered.levels[0]!.levelText).toBe('%1.')
    const body = readLegacyDocBody(bytes)
    const listBlocks = body.blocks.filter(
      (block) => block.type === 'paragraph' && block.list !== undefined,
    )
    expect(
      listBlocks.map((block) => {
        if (block.type !== 'paragraph') throw new Error('unreachable')
        return { text: block.paragraph.text, numId: block.list!.numId, ilvl: block.list!.ilvl }
      }),
    ).toEqual([
      { text: 'Bullet A', numId: '101', ilvl: 0 },
      { text: 'Bullet B', numId: '101', ilvl: 0 },
      { text: 'Number 1', numId: '104', ilvl: 0 },
    ])
    const linkRun = body.blocks
      .filter((block) => block.type === 'paragraph')
      .flatMap((block) => (block.type === 'paragraph' ? block.paragraph.runs : []))
      .find((run) => run.link !== undefined)
    expect(linkRun?.text).toBe('the site')
    expect(linkRun?.link?.href).toBe('https://example.org/')
  })

  it('writes numPr, numbering definitions and an external hyperlink', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings).toEqual([])
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect((xml.match(/<w:numPr>/g) ?? []).length).toBe(3)
    expect(xml).toContain('<w:numId w:val="101"/>')
    expect(xml).toContain('<w:numId w:val="104"/>')
    expect(xml).toContain('<w:hyperlink r:id=')
    expect(xml).toContain('the site')
    expect(xml).not.toContain('HYPERLINK')
    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).toMatch(
      /Type="[^"]+\/hyperlink" Target="https:\/\/example\.org\/" TargetMode="External"/,
    )
    const numbering = (await zipText(docx, 'word/numbering.xml'))!
    expect(numbering).toContain('<w:abstractNum w:abstractNumId="10">')
    expect(numbering).toContain('<w:abstractNum w:abstractNumId="11">')
    expect(numbering).toContain('<w:numFmt w:val="bullet"/>')
    // the Symbol-font bullet keeps its private-use code with the level's rFonts
    expect(numbering).toContain(
      '<w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>',
    )
    expect(numbering).toContain('<w:lvlText w:val="%1."/>')
    expect(numbering).toContain('<w:num w:numId="101">')
    expect(numbering).toContain('<w:num w:numId="104">')
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })
})

describe('rich.doc: lists + hyperlink alongside table and image', () => {
  it('keeps the stage 3c-2/3c-3 content and reconstructs stage 3c-4', async () => {
    const bytes = fixtureBytes('rich.doc')
    const body = readLegacyDocBody(bytes)
    expect(body.blocks.some((block) => block.type === 'table')).toBe(true)
    expect(body.blocks.some((block) => block.type === 'image')).toBe(true)
    const links = body.blocks
      .filter((block) => block.type === 'paragraph')
      .flatMap((block) => (block.type === 'paragraph' ? block.paragraph.runs : []))
      .filter((run) => run.link !== undefined)
    expect(links.map((run) => run.link?.href)).toEqual(['https://example.com/doc'])
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain('<w:tbl>')
    expect(xml).toContain('<w:drawing>')
    expect(xml).toContain('<w:hyperlink r:id=')
    expect(xml).toContain('<w:numId w:val="101"/>')
    expect(warnings.some((warning) => warning.includes('lists are not converted'))).toBe(false)
    expect(warnings.some((warning) => warning.includes('hyperlinks are not converted'))).toBe(false)
    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).toContain('Target="https://example.com/doc" TargetMode="External"')
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })
})

describe('synthetic list tables', () => {
  it('applies LFOLVL overrides: a restart value and a replacement bullet level', async () => {
    const bytes = syntheticListDoc({
      text: 'One\rTwo\r',
      paragraphs: [
        { startCp: 0, ilfo: 1 },
        { startCp: 4, ilfo: 2 },
      ],
      lists: [
        { lsid: 7, simple: true, levels: [{ numberFormat: 0, levelText: '\u0000.' }] },
        { lsid: 8, simple: true, levels: [{ numberFormat: 0, levelText: '*' }] },
      ],
      listLevelOverrides: [
        [{ ilvl: 0, startAt: 5 }],
        [{ ilvl: 0, level: { numberFormat: 23, levelText: '\uf0b7', indentLeft: 360 } }],
      ],
    })
    const doc = openLegacyDoc(bytes)
    const tables = parseListTables(doc.fib, doc.table)
    expect(tables.warnings).toEqual([])
    expect(tables.levelOverrides.get(1)).toEqual([{ ilvl: 0, startAt: 5 }])
    expect(tables.levelOverrides.get(2)?.[0]?.level?.numberFormat).toBe(23)
    const body = readLegacyDocBody(bytes)
    const kinds = body.blocks.map((block) => (block.type === 'paragraph' ? block.list?.kind : '?'))
    expect(kinds.slice(0, 2)).toEqual(['ordered', 'bullet'])
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings).toEqual([])
    const numbering = (await zipText(docx, 'word/numbering.xml'))!
    expect(numbering).toContain(
      '<w:num w:numId="101"><w:abstractNumId w:val="10"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>',
    )
    expect(numbering).toMatch(
      /<w:num w:numId="102"><w:abstractNumId w:val="11"\/><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:start w:val="1"\/><w:numFmt w:val="bullet"\/>.*<w:lvlText w:val="\u2022"\/>.*<\/w:lvlOverride><\/w:num>/,
    )
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })

  it('parses levels, placeholders and numFmt values', () => {
    const bytes = syntheticListDoc({
      text: 'One\rTwo\r',
      paragraphs: [
        { startCp: 0, ilfo: 1 },
        { startCp: 4, ilfo: 1, ilvl: 1 },
      ],
      lists: [
        {
          lsid: 42,
          levels: [
            { numberFormat: 23, levelText: '\uf0b7', indentLeft: 720, indentFirstLine: -360 },
            { numberFormat: 0, levelText: '\u0001.' },
            { numberFormat: 1, levelText: '\u0002)' },
            { numberFormat: 3, levelText: '\u0003.' },
          ],
        },
      ],
    })
    const doc = openLegacyDoc(bytes)
    const lists = parseListTables(doc.fib, doc.table)
    expect(lists.overrides.get(1)).toBe(42)
    const definition = lists.lists.get(42)!
    expect(definition.levels).toHaveLength(9)
    expect(definition.levels.slice(0, 4).map((level) => level.numberFormat)).toEqual([23, 0, 1, 3])
    expect(definition.levels.slice(0, 4).map((level) => level.levelText)).toEqual([
      '\uf0b7',
      '%2.',
      '%3)',
      '%4.',
    ])
    expect(definition.levels[0]!.indentLeftTwips).toBe(720)
    expect(definition.levels[0]!.indentFirstLineTwips).toBe(-360)
    const body = readLegacyDocBody(bytes)
    const listBlocks = body.blocks.filter(
      (block) => block.type === 'paragraph' && block.list !== undefined,
    )
    expect(listBlocks).toHaveLength(2)
    if (listBlocks[0]!.type !== 'paragraph' || listBlocks[1]!.type !== 'paragraph')
      throw new Error('expected list paragraphs')
    expect(listBlocks[0]!.list).toEqual({ kind: 'bullet', numId: '101', ilvl: 0 })
    expect(listBlocks[1]!.list).toEqual({ kind: 'ordered', numId: '101', ilvl: 1 })
  })

  it('writes the reconstructed levels (numFmt, start, indents) into numbering.xml', async () => {
    const bytes = syntheticListDoc({
      text: 'One\r',
      paragraphs: [{ startCp: 0, ilfo: 1 }],
      lists: [
        {
          lsid: 7,
          levels: [
            { numberFormat: 23, levelText: '\uf0b7', indentLeft: 720, indentFirstLine: -360 },
            { numberFormat: 1, levelText: '\u0001)' },
            { numberFormat: 23, levelText: '\uf0a7' },
          ],
        },
      ],
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const numbering = (await zipText(docx, 'word/numbering.xml'))!
    expect(numbering).toContain('<w:numFmt w:val="bullet"/>')
    expect(numbering).toContain('<w:numFmt w:val="decimal"/>')
    expect(numbering).toContain('<w:numFmt w:val="upperRoman"/>')
    expect(numbering).toContain('<w:lvlText w:val="%2)"/>')
    expect(numbering).toContain('<w:lvlText w:val="\u25aa"/>')
    expect(numbering).toContain('<w:ind w:left="720" w:hanging="360"/>')
    expect(numbering).toContain('<w:num w:numId="101">')
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="101"/></w:numPr>')
  })

  it('reads grpprlPapx and grpprlChpx of unequal lengths in LVLF order (cbChpx first)', () => {
    const bytes = syntheticListDoc({
      text: 'One\rTwo\r',
      paragraphs: [
        { startCp: 0, ilfo: 1 },
        { startCp: 4, ilfo: 1, ilvl: 1 },
      ],
      lists: [
        {
          lsid: 7,
          levels: [
            {
              numberFormat: 0,
              levelText: '\u0000.',
              indentLeft: 720,
              indentFirstLine: -360,
              characterSprms: encodeSprm(0x0835, [1]),
            },
            {
              numberFormat: 4,
              levelText: '\u0001)',
              indentLeft: 1440,
              characterSprms: new Uint8Array([
                ...encodeSprm(0x0835, [1]),
                ...encodeSprm(0x0836, [1]),
                ...encodeSprm(0x4a43, [24, 0]),
              ]),
            },
          ],
        },
      ],
    })
    const doc = openLegacyDoc(bytes)
    const lists = parseListTables(doc.fib, doc.table)
    const levels = lists.lists.get(7)!.levels
    expect(levels[0]).toMatchObject({
      levelText: '%1.',
      indentLeftTwips: 720,
      indentFirstLineTwips: -360,
    })
    expect(levels[0]!.characterSprms).toHaveLength(3)
    expect(levels[1]).toMatchObject({ levelText: '%2)', indentLeftTwips: 1440, numberFormat: 4 })
    expect(levels[1]!.characterSprms).toHaveLength(10)
    expect(levels[1]!.paragraphSprms).toHaveLength(4)
    expect(levels[2]!.levelText).toBe('%3.')
  })

  it('writes the level number East Asian face next to its Latin face', async () => {
    const bytes = syntheticListDoc({
      text: 'One\r',
      paragraphs: [{ startCp: 0, ilfo: 1 }],
      fonts: ['Times New Roman', 'Arial', 'SimSun'],
      lists: [
        {
          lsid: 7,
          simple: true,
          levels: [
            {
              numberFormat: 0,
              levelText: '\u0000.',
              characterSprms: new Uint8Array([
                ...encodeSprm(0x4a4f, [1, 0]),
                ...encodeSprm(0x4a50, [2, 0]),
              ]),
            },
          ],
        },
      ],
    })
    const numbering = (await zipText(
      (await convertLegacyDocToDocx(bytes)).docx,
      'word/numbering.xml',
    ))!
    expect(numbering).toContain(
      '<w:rFonts w:ascii="Arial" w:eastAsia="SimSun" w:hAnsi="Arial" w:hint="default"/>',
    )
  })

  it('supports a simple (single-level) list', () => {
    const bytes = syntheticListDoc({
      text: 'Only\r',
      paragraphs: [{ startCp: 0, ilfo: 1 }],
      lists: [{ lsid: 3, simple: true, levels: [{ numberFormat: 0, levelText: '\u0000.' }] }],
    })
    const doc = openLegacyDoc(bytes)
    const lists = parseListTables(doc.fib, doc.table)
    const definition = lists.lists.get(3)!
    expect(definition.simple).toBe(true)
    expect(definition.levels).toHaveLength(1)
  })

  it('warns when a paragraph references a missing list', () => {
    const bytes = syntheticListDoc({
      text: 'Orphan\r',
      paragraphs: [{ startCp: 0, ilfo: 9 }],
      lists: [{ lsid: 42, levels: [{ numberFormat: 0, levelText: '\u0000.' }] }],
    })
    const body = readLegacyDocBody(bytes)
    expect(body.warnings.some((warning) => warning.includes('list (ilfo 9)'))).toBe(true)
    expect(
      body.blocks.some((block) => block.type === 'paragraph' && block.list === undefined),
    ).toBe(true)
  })

  it('drops malformed PlfLst tables with a warning instead of failing', () => {
    const header = encodePlfLst([
      { lsid: 1, levels: [{ numberFormat: 0, levelText: '\u0000.' }] },
    ]).header
    const corruptedCount = new Uint8Array([0xff, 0xff])
    const badCount = buildSyntheticDoc({
      pieces: [{ text: 'x\r' }],
      plfLst: corruptedCount,
    })
    const doc = openLegacyDoc(badCount)
    const lists = parseListTables(doc.fib, doc.table)
    expect(lists.lists.size).toBe(0)
    expect(lists.warnings.some((warning) => warning.includes('declares 65535 lists'))).toBe(true)
    const truncatedLevels = buildSyntheticDoc({ pieces: [{ text: 'x\r' }], plfLst: header })
    const truncatedDoc = openLegacyDoc(truncatedLevels)
    const truncated = parseListTables(truncatedDoc.fib, truncatedDoc.table)
    expect(truncated.lists.size).toBe(0)
    expect(truncated.warnings.some((warning) => warning.includes('truncated (no LVLF)'))).toBe(true)
    expect(() => convertLegacyDocToDocx(truncatedLevels)).not.toThrow()
  })

  it('keeps the lists read before a truncated level', () => {
    const encoded = encodePlfLst([
      { lsid: 1, simple: true, levels: [{ numberFormat: 0, levelText: '\u0000.' }] },
      { lsid: 2, simple: true, levels: [{ numberFormat: 23, levelText: '\uf0b7' }] },
    ])
    const table = new Uint8Array(encoded.header.length + encoded.levels.length - 20)
    table.set(encoded.header, 0)
    table.set(encoded.levels.subarray(0, encoded.levels.length - 20), encoded.header.length)
    const doc = openLegacyDoc(
      buildSyntheticDoc({
        pieces: [{ text: 'x\r' }],
        plfLst: table,
        lcbPlfLst: encoded.header.length,
      }),
    )
    const lists = parseListTables(doc.fib, doc.table)
    expect([...lists.lists.keys()]).toEqual([1])
    expect(lists.warnings.some((warning) => warning.includes('level 0 of list 1'))).toBe(true)
  })

  it('keeps a level whose grpprlPapx holds a truncated sprm', () => {
    const bytes = syntheticListDoc({
      text: 'One\r',
      paragraphs: [{ startCp: 0, ilfo: 1 }],
      lists: [
        {
          lsid: 5,
          simple: true,
          levels: [
            {
              numberFormat: 0,
              levelText: '\u0000.',
              paragraphSprms: new Uint8Array([
                ...encodeSprm(0x845e, [0xd0, 0x02]),
                0x5e,
                0x84,
                0x01,
              ]),
            },
          ],
        },
      ],
    })
    const doc = openLegacyDoc(bytes)
    const lists = parseListTables(doc.fib, doc.table)
    const level = lists.lists.get(5)!.levels[0]!
    expect(level.indentLeftTwips).toBe(720)
    expect(level.levelText).toBe('%1.')
    expect(lists.warnings.some((warning) => warning.includes('level 0 of list 0'))).toBe(true)
    const body = readLegacyDocBody(bytes)
    expect(body.blocks[0]).toMatchObject({ type: 'paragraph', list: { numId: '101' } })
  })

  it('drops an out-of-range PlfLfo table with a warning', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'x\r' }],
      plfLst: new Uint8Array([0, 0]),
    })
    const doc = openLegacyDoc(bytes)
    const fib = { ...doc.fib, fcPlfLfo: doc.table.length - 2, lcbPlfLfo: 40 }
    const lists = parseListTables(fib, doc.table)
    expect(lists.overrides.size).toBe(0)
    expect(lists.warnings.some((warning) => warning.includes('list override table'))).toBe(true)
  })
})

describe('HYPERLINK field parser', () => {
  it('parses the fixture field into a result range', () => {
    const rawText = 'Visit \u0013 HYPERLINK "https://example.org/"\u0001\u0014the site\u0015\r'
    const parsed = parseFields(rawText)
    expect(parsed.warnings).toEqual([])
    const link = parsed.hyperlinks[0]!
    expect(link.target).toBe('https://example.org/')
    expect(rawText.slice(link.startCp, link.endCp)).toBe('the site')
  })

  it('handles quoted URLs, bare URLs and \\l / \\o switches', () => {
    const quoted = parseFields('\u0013 HYPERLINK "https://example.test/a b" \u0014result\u0015')
    expect(quoted.hyperlinks[0]!.target).toBe('https://example.test/a b')
    const bare = parseFields('\u0013 HYPERLINK https://example.test/bare \u0014x\u0015')
    expect(bare.hyperlinks[0]!.target).toBe('https://example.test/bare')
    const anchor = parseFields('\u0013 HYPERLINK \\l "bookmark" \\o "jump here" \u0014jump\u0015')
    expect(anchor.hyperlinks[0]).toMatchObject({ anchor: 'bookmark', tooltip: 'jump here' })
    const both = parseFields('\u0013 HYPERLINK "https://example.test/" \\o "tip" \u0014site\u0015')
    expect(both.hyperlinks[0]).toMatchObject({
      target: 'https://example.test/',
      tooltip: 'tip',
    })
  })

  it('warns about unsupported and malformed fields instead of failing', () => {
    const page = parseFields('\u0013 PAGE \u0014 1 \u0015')
    expect(page.hyperlinks).toEqual([])
    expect(page.simpleFields).toEqual([{ startCp: 8, endCp: 11, instruction: 'PAGE' }])
    expect(page.warnings).toEqual([])
    const seq = parseFields('\u0013 SEQ Figure \u0014 1 \u0015')
    expect(seq.warnings.some((warning) => warning.includes('SEQ field'))).toBe(true)
    const noSeparator = parseFields('\u0013 HYPERLINK "https://example.test/" \u0015')
    expect(noSeparator.hyperlinks).toEqual([])
    expect(noSeparator.warnings.some((warning) => warning.includes('result separator'))).toBe(true)
    const noTarget = parseFields('\u0013 HYPERLINK \u0014result\u0015')
    expect(noTarget.hyperlinks).toEqual([])
    expect(noTarget.warnings.some((warning) => warning.includes('malformed'))).toBe(true)
    const unclosed = parseFields('\u0013 HYPERLINK "https://example.test/" \u0014site')
    expect(unclosed.hyperlinks).toEqual([])
    expect(unclosed.warnings.some((warning) => warning.includes('ends inside a field'))).toBe(true)
  })

  it('ignores empty field results and keeps nested fields sane', () => {
    const empty = parseFields('\u0013 HYPERLINK "https://example.test/" \u0014\u0015')
    expect(empty.hyperlinks).toEqual([])
    const nested = parseFields(
      '\u0013 IF 1 = 1 \u0013 HYPERLINK "https://example.test/" \u0014inner\u0015 \u0014outer\u0015',
    )
    expect(nested.hyperlinks).toHaveLength(1)
    expect(nested.warnings.some((warning) => warning.includes('IF field'))).toBe(true)
  })
})

describe('synthetic hyperlink conversion', () => {
  it('wraps the result runs in w:hyperlink and adds an external relationship', async () => {
    const bytes = buildSyntheticDoc({
      pieces: [
        { text: 'Visit \u0013 HYPERLINK "https://example.test/link"\u0001\u0014the site\u0015\r' },
      ],
    })
    const body = readLegacyDocBody(bytes)
    const links = body.blocks
      .filter((block) => block.type === 'paragraph')
      .flatMap((block) => (block.type === 'paragraph' ? block.paragraph.runs : []))
      .filter((run) => run.link !== undefined)
    expect(links.map((run) => run.text)).toEqual(['the site'])
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain('<w:hyperlink r:id=')
    expect(xml).not.toContain('HYPERLINK')
    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).toContain('Target="https://example.test/link" TargetMode="External"')
  })

  it('writes an internal bookmark link as w:anchor without a relationship', async () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'Go \u0013 HYPERLINK \\l "chapter1" \u0014there\u0015\r' }],
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain('<w:hyperlink w:anchor="chapter1">')
    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).not.toContain('/hyperlink')
  })

  it('keeps the file target of a bookmark link into another document', async () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'Go \u0013 HYPERLINK "other.docx" \\l "chapter1" \u0014there\u0015\r' }],
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).not.toContain('w:anchor=')
    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).toContain('Target="other.docx#chapter1"')
  })

  it('drops the Hyperlink character style inside a TOC entry but keeps it on other links', async () => {
    // Word ignores the Hyperlink style in TOC field results (probed: black text, no underline)
    const text =
      '\u0013 TOC \\o "1-3" \\h \u0014\u0013 HYPERLINK \\l "_Toc1" \u0014Entry\u0015\u0015\r' +
      'See \u0013 HYPERLINK \\l "_Toc1" \u0014Link\u0015\r'
    const fields = parseFields(text)
    expect(fields.hyperlinks.map((link) => link.inToc)).toEqual([true, undefined])
    const hyperlinkStyle = new Uint8Array([
      ...encodeSprm(0x2a3e, [0x01]),
      ...encodeSprm(0x6870, [0x00, 0x00, 0xff, 0x00]),
    ])
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      stshStyles: [
        { name: 'Normal', kind: 'paragraph', sti: 0 },
        { name: 'Hyperlink', kind: 'character', sti: 85, chpxSprms: hyperlinkStyle },
      ],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET,
            fcEnd: SYNTHETIC_TEXT_OFFSET + text.length * 2,
            grpprl: encodeSprm(0x4a30, [0x01, 0x00]),
          },
        ]),
      ],
    })
    const runs = readLegacyDocBody(bytes)
      .blocks.flatMap((block) => (block.type === 'paragraph' ? block.paragraph.runs : []))
      .filter((run) => run.link !== undefined)
    expect(runs.map((run) => [run.text, run.link!.plain, run.formatting.colorRgb])).toEqual([
      ['Entry', true, undefined],
      ['Link', undefined, '0000FF'],
    ])
    const xml = (await zipText((await convertLegacyDocToDocx(bytes)).docx, 'word/document.xml'))!
    const entry = xml.slice(xml.indexOf('<w:hyperlink'), xml.indexOf('Entry</w:t>'))
    expect(entry).not.toMatch(/Hyperlink|0000FF|<w:u /)
    const link = xml.slice(xml.lastIndexOf('<w:hyperlink'), xml.indexOf('Link</w:t>'))
    expect(link).toContain('<w:rStyle w:val="Hyperlink"/>')
    expect(link).toContain('<w:color w:val="0000FF"/>')
    expect(link).toContain('<w:u w:val="single"/>')
  })

  it('passes an Office Open XML file saved under a .doc name through unchanged', async () => {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', '<Types/>')
    const bytes = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(docx).toBe(bytes)
    expect(warnings).toEqual(['the file is already Office Open XML; it was kept as is'])
  })

  it('writes PAGE / NUMPAGES fields as live complex fields with their cached result', async () => {
    const text = 'Page \u0013 PAGE \u0014 3\u0015 of \u0013NUMPAGES\u0014 9\u0015\r'
    const bytes = buildSyntheticDoc({ pieces: [{ text }] })
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings).toEqual([])
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve"> 3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    )
    expect(xml).toContain('<w:instrText xml:space="preserve"> NUMPAGES </w:instrText>')
    expect(xml).not.toMatch(/[\ue000-\ue002]/)
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })

  it('keeps a PAGE field whose cached result is empty', async () => {
    const text = 'Page \u0013 PAGE \u0014\u0015 end\r'
    const bytes = buildSyntheticDoc({ pieces: [{ text }] })
    expect(parseFields(text).simpleFields).toEqual([
      { startCp: 13, endCp: 13, instruction: 'PAGE' },
    ])
    const xml = (await zipText((await convertLegacyDocToDocx(bytes)).docx, 'word/document.xml'))!
    expect(xml).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>')
    expect(xml).toMatch(
      /Page <\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="begin"\/>.*?<w:fldChar w:fldCharType="end"\/><\/w:r><w:r><w:t xml:space="preserve"> end<\/w:t>/,
    )
  })
})

describe('NUMBER_FORMATS', () => {
  it('maps every nfc to a value Word accepts in w:numFmt', async () => {
    const { NUMBER_FORMATS } = await import('../src/lists')
    const valid = new Set([
      'decimal',
      'upperRoman',
      'lowerRoman',
      'upperLetter',
      'lowerLetter',
      'ordinal',
      'cardinalText',
      'ordinalText',
      'hex',
      'chicago',
      'ideographDigital',
      'japaneseCounting',
      'aiueo',
      'iroha',
      'decimalFullWidth',
      'decimalHalfWidth',
      'japaneseLegal',
      'japaneseDigitalTenThousand',
      'decimalEnclosedCircle',
      'decimalFullWidth2',
      'aiueoFullWidth',
      'irohaFullWidth',
      'decimalZero',
      'bullet',
      'ganada',
      'chosung',
      'decimalEnclosedFullstop',
      'decimalEnclosedParen',
      'decimalEnclosedCircleChinese',
      'ideographEnclosedCircle',
      'ideographTraditional',
      'ideographZodiac',
      'ideographZodiacTraditional',
      'taiwaneseCounting',
      'ideographLegalTraditional',
      'taiwaneseCountingThousand',
      'taiwaneseDigital',
      'chineseCounting',
      'chineseLegalSimplified',
      'chineseCountingThousand',
      'koreanDigital',
      'koreanCounting',
      'koreanLegal',
      'koreanDigital2',
      'vietnameseCounting',
      'russianLower',
      'russianUpper',
      'none',
      'numberInDash',
      'hebrew1',
      'hebrew2',
      'arabicAlpha',
      'arabicAbjad',
      'hindiVowels',
      'hindiConsonants',
      'hindiNumbers',
      'hindiCounting',
      'thaiLetters',
      'thaiNumbers',
      'thaiCounting',
      'bahtText',
      'dollarText',
    ])
    for (const [nfc, fmt] of Object.entries(NUMBER_FORMATS))
      expect(valid.has(fmt), `nfc ${nfc} -> ${fmt}`).toBe(true)
    expect(NUMBER_FORMATS[20]).toBe('aiueoFullWidth')
    expect(NUMBER_FORMATS[22]).toBe('decimalZero')
    expect(NUMBER_FORMATS[0x28]).toBe('decimal')
    expect(NUMBER_FORMATS[0x2d]).toBe('hebrew1')
    expect(NUMBER_FORMATS[0x38]).toBe('vietnameseCounting')
    expect(NUMBER_FORMATS[0x3a]).toBe('russianLower')
    expect(NUMBER_FORMATS[0xff]).toBe('none')
  })
})
