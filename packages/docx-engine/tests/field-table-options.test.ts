/** Field switches are read and rewritten in place, and a generated table's lines follow its layout. */
import { describe, expect, it } from 'vitest'
import {
  fieldInstructionOfXml,
  fieldSwitchArg,
  formatFieldInstruction,
  generateFieldTableXml,
  parseFieldInstruction,
  patchFieldTablePageXml,
  parseDocx,
  withFieldSwitch,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

describe('field instruction switches', () => {
  it('reads keyword, bare switches and quoted or bare arguments', () => {
    const instr = parseFieldInstruction(' TOC \\o "1-3" \\h \\z \\u \\n 2-3 \\p " " ')
    expect(instr.keyword).toBe('TOC')
    expect(fieldSwitchArg(instr, 'o')).toBe('1-3')
    expect(fieldSwitchArg(instr, 'h')).toBe('')
    expect(fieldSwitchArg(instr, 'n')).toBe('2-3')
    expect(fieldSwitchArg(instr, 'p')).toBe(' ')
    expect(fieldSwitchArg(instr, 'c')).toBeUndefined()
  })

  it('sets, replaces and removes one switch, the others keeping their order', () => {
    const instr = parseFieldInstruction(' TOC \\o "1-3" \\h \\z \\u ')
    expect(formatFieldInstruction(withFieldSwitch(instr, 'o', '1-2'))).toBe(
      ' TOC \\o "1-2" \\h \\z \\u ',
    )
    expect(formatFieldInstruction(withFieldSwitch(instr, 'h', false))).toBe(
      ' TOC \\o "1-3" \\z \\u ',
    )
    expect(formatFieldInstruction(withFieldSwitch(instr, 'n', true))).toBe(
      ' TOC \\o "1-3" \\h \\z \\u \\n ',
    )
    expect(formatFieldInstruction(parseFieldInstruction('INDEX \\e "\t" \\c "2"'))).toBe(
      ' INDEX \\e "\t" \\c "2" ',
    )
  })

  it('joins the instruction runs of a paragraph, up to the field separator', () => {
    const xml =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> TOC \\o </w:instrText></w:r>' +
      '<w:r><w:instrText>"1-2" \\p "&amp;"</w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:instrText> PAGEREF _Toc1 \\h </w:instrText></w:r></w:p>'
    expect(fieldInstructionOfXml(xml)).toBe(' TOC \\o "1-2" \\p "&"')
  })
})

describe('generateFieldTableXml', () => {
  const lines = [
    { text: 'One', pageNo: 1, style: 'TOC1' },
    { text: 'Two', pageNo: 4, style: 'TOC2' },
  ]

  it('puts the page behind a right tab with the chosen leader', () => {
    const xml = generateFieldTableXml(lines, { instr: ' TOC \\o "1-2" ', leader: 'hyphen' })
    expect(xml).toHaveLength(2)
    expect(xml[0]).toContain('<w:tab w:val="right" w:leader="hyphen" w:pos="9350"/>')
    expect(xml[0]).toContain(' TOC \\o "1-2" ')
    expect(xml[0]).toContain('<w:tab/>')
    expect(xml[1]).toContain('<w:t>4</w:t>')
    expect(xml[1]).toContain('fldCharType="end"')
  })

  it('writes no tab stop and no tab when the page follows a separator', () => {
    const xml = generateFieldTableXml(lines, { instr: ' TOC \\p " " ', pageSep: ' ' })
    expect(xml[0]).not.toContain('<w:tabs>')
    expect(xml[0]).not.toContain('<w:tab/>')
    expect(xml[0]).toContain('<w:t xml:space="preserve"> 1</w:t>')
    expect(patchFieldTablePageXml(xml[0], ' 12')).toContain('<w:t xml:space="preserve"> 12</w:t>')
    expect(patchFieldTablePageXml(xml[0], ' 12')).toContain('>One</w:t>')
  })

  it('leaves the page number out of a line that takes none', () => {
    const xml = generateFieldTableXml(
      lines.map((line) => ({ ...line, noPage: true })),
      { instr: ' TOC \\n ' },
    )
    expect(xml.join('')).not.toContain('w:tab')
    expect(xml.join('')).not.toContain('<w:t>1</w:t>')
  })
})

describe('a saved table read back', () => {
  const linesOf = async (bodyXml: string) =>
    (await parseDocx(await buildDocx({ bodyXml }))).blocks
      .map((b) => b.fieldDisplay)
      .filter((d) => d?.kind === 'tocLine')
      .map((d) => [d?.left, d?.noLink ?? false])
  const toc = (instr: string) =>
    generateFieldTableXml(
      [
        { text: 'One', pageNo: 1, style: 'TOC1' },
        { text: 'Two', pageNo: 4, style: 'TOC1' },
      ],
      { instr },
    ).join('')

  it('keeps the lines of a table without \\h, or of a table of authorities, out of the links', async () => {
    expect(await linesOf(toc(' TOC \\o "1-3" \\z \\u '))).toEqual([
      ['One', true],
      ['Two', true],
    ])
    expect(await linesOf(toc(' TOC \\o "1-3" \\h \\z \\u '))).toEqual([
      ['One', false],
      ['Two', false],
    ])
    const authorities = generateFieldTableXml(
      [
        { text: 'Cases', style: 'TOAHeading' },
        { text: 'Baldwin v. Alberti', pageNo: 3, style: 'TableofAuthorities' },
      ],
      { instr: ' TOA \\h \\p ' },
    ).join('')
    // a later table with \\h is linked again
    expect(await linesOf(authorities + toc(' TOC \\h '))).toEqual([
      ['Cases', true],
      ['Baldwin v. Alberti', true],
      ['One', false],
      ['Two', false],
    ])
  })
})

describe('an index read back', () => {
  const shown = async (bodyXml: string) =>
    (await parseDocx(await buildDocx({ bodyXml }))).blocks
      .filter((b) => b.originalXml?.startsWith('<w:p>'))
      .map((b) => {
        const d = b.fieldDisplay
        return d?.kind === 'tocLine'
          ? ['line', d.level, d.left, d.right, d.leader, d.noLink ?? false]
          : [b.type, d?.left ?? b.runs?.map((r) => r.text).join('')]
      })

  it('keeps entry and page apart when the page sits behind a tab, in every paragraph of the field', async () => {
    const xml = generateFieldTableXml(
      [
        { text: 'Apple', pageNo: 1, style: 'Index1' },
        { text: 'Green', pageNo: '2, 5', style: 'Index2', indent: 240 },
        { text: 'Mango', pageNo: 3, style: 'Index1' },
        { text: 'Zebra', pageNo: 4, style: 'Index1' },
      ],
      { instr: ' INDEX \\e "\t" \\c "2" ', leader: 'hyphen', tabPos: 4300 },
    ).join('')
    expect(await shown(xml)).toEqual([
      ['line', 1, 'Apple', '1', 'hyphen', true],
      ['line', 2, 'Green', '2, 5', 'hyphen', true],
      ['line', 1, 'Mango', '3', 'hyphen', true],
      ['line', 1, 'Zebra', '4', 'hyphen', true],
    ])
  })

  it('reads the tab-leader index Word writes the same way, and leaves its other forms as text', async () => {
    const p = (style: string, inner: string, tabs = '') =>
      `<w:p><w:pPr><w:pStyle w:val="${style}"/>${tabs}</w:pPr>${inner}</w:p>`
    const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
    const begin = (instr: string) =>
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      `<w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    const stop = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="4310"/></w:tabs>'
    const tabbed =
      p(
        'Index1',
        begin(' INDEX \\e "\t" \\c "2" \\z "1033" ') + r('Apple') + '<w:r><w:tab/></w:r>' + r('1'),
        stop,
      ) +
      p('Index1', r('Mango') + '<w:r><w:tab/></w:r>' + r('3'), stop) +
      p('Index1', r('Zebra') + '<w:r><w:tab/></w:r>' + r('4') + end, stop)
    expect(await shown(tabbed)).toEqual([
      ['line', 1, 'Apple', '1', 'dot', true],
      ['line', 1, 'Mango', '3', 'dot', true],
      ['line', 1, 'Zebra', '4', 'dot', true],
    ])
    // page after a comma, and run-in: no tab, so nothing to keep apart
    const comma =
      p('Index1', begin(' INDEX \\c "2" ') + r('Apple, 1')) +
      p('Index1', r('Mango, 3')) +
      p('Index1', r('Zebra, 4') + end)
    expect(await shown(comma)).toEqual([
      ['passthrough', 'Apple, 1'],
      ['paragraph', 'Mango, 3'],
      ['passthrough', 'Zebra, 4'],
    ])
    const runIn =
      p('Index1', begin(' INDEX \\r \\c "2" ') + r('Fruit: Apple, 1; Pear, 3')) +
      p('Index1', r('Zebra, 4') + end)
    expect(await shown(runIn)).toEqual([
      ['passthrough', 'Fruit: Apple, 1; Pear, 3'],
      ['passthrough', 'Zebra, 4'],
    ])
    // an Index-styled paragraph with a tab outside any index stays a paragraph
    expect(
      (await shown(p('Index1', r('Loose') + '<w:r><w:tab/></w:r>' + r('9'), stop)))[0][0],
    ).toBe('paragraph')
  })
  it('gives a subentry one indent: the direct one, or the style\u2019s, never the TOC ladder on top', async () => {
    const indents = async (bodyXml: string, extraStylesXml = '') =>
      (await parseDocx(await buildDocx({ bodyXml, extraStylesXml }))).blocks
        .filter((b) => b.fieldDisplay?.kind === 'tocLine')
        .map((b) => [b.fieldDisplay?.left, b.fieldDisplay?.level, b.fieldDisplay?.indentLeftTwips])
    const lines = (indent?: number) =>
      generateFieldTableXml(
        [
          { text: 'Fruit', pageNo: 1, style: 'Index1' },
          { text: 'Apple', pageNo: 2, style: 'Index2', indent },
          { text: 'Zebra', pageNo: 4, style: 'Index1' },
        ],
        { instr: ' INDEX \\e "\t" ', tabPos: 4300 },
      ).join('')
    // as generated: no Index styles in the file, the subentry indents by its own w:ind
    expect(await indents(lines(240))).toEqual([
      ['Fruit', 1, 0],
      ['Apple', 2, 240],
      ['Zebra', 1, 0],
    ])
    // as Word writes it: the indent is the Index 2 style's, painted once through the style
    const index2 =
      '<w:style w:type="paragraph" w:styleId="Index2"><w:name w:val="index 2"/>' +
      '<w:pPr><w:ind w:left="480" w:hanging="240"/></w:pPr></w:style>'
    expect(await indents(lines(), index2)).toEqual([
      ['Fruit', 1, 0],
      ['Apple', 2, 0],
      ['Zebra', 1, 0],
    ])
    // a direct indent over the style's adds only the difference
    expect((await indents(lines(720), index2))[1]).toEqual(['Apple', 2, 240])
  })
})
