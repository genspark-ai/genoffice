/** Table of authorities: citations (TA fields) ride with their text, and the table is a TOA field. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  generateFieldTableXml,
  generateTableOfAuthoritiesFieldXml,
  parseDocx,
  saveDocx,
} from '../src/index'
import type { Run } from '../src/types'
import { buildDocx } from './helpers/build-docx'

const field = (instr: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

const CITE = 'TA \\l "Baldwin v. Alberti, 58 Wn. 2d 243 (1961)" \\s "Baldwin" \\c 1'

describe('table of authorities citations', () => {
  it('reads a citation as a marker after its text and keeps the paragraph editable', async () => {
    const body = `<w:p><w:r><w:t>See Baldwin</w:t></w:r>${field(CITE)}<w:r><w:t xml:space="preserve"> above.</w:t></w:r></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const block = parsed.blocks[0]
    expect(block.type).toBe('paragraph')
    const runs = (block as { runs: Run[] }).runs
    const marker = runs.find((r) => r.taInstr !== undefined)!
    expect(marker.taInstr).toBe(
      '\\l "Baldwin v. Alberti, 58 Wn. 2d 243 (1961)" \\s "Baldwin" \\c 1',
    )
    // it sits in the index entry's slot, so everything that keeps an index entry whole keeps it
    expect(marker.xeTerm).toBe('')
    expect(runs.map((r) => r.text).join('')).toBe('See Baldwin above.')
  })

  it('writes the citation back with its instruction when the paragraph is edited', async () => {
    const body = `<w:p><w:r><w:t>See Baldwin</w:t></w:r>${field(CITE)}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const runs = (parsed.blocks[0] as { runs: Run[] }).runs
    const saved = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'Edited. ' }, ...runs] },
        docxIndex: 0,
      },
    ])
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain(` ${CITE} `)
    expect(xml).not.toContain(' XE ')
    const reparsed = (await parseDocx(saved)).blocks[0] as { runs: Run[] }
    expect(reparsed.runs.filter((r) => r.taInstr !== undefined)).toHaveLength(1)
  })

  it('an index entry is still written as an index entry', async () => {
    const body = `<w:p><w:r><w:t>term</w:t></w:r>${field('XE "term"')}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const runs = (parsed.blocks[0] as { runs: Run[] }).runs
    expect(runs.some((r) => r.taInstr !== undefined)).toBe(false)
    const saved = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs }, docxIndex: 0 },
    ])
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain(' XE "term" ')
  })
})

describe('a citation keeps the formatting of its long text', () => {
  const instr = (text: string, rPr = '') =>
    `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:instrText xml:space="preserve">${text}</w:instrText></w:r>`
  const formatted =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    instr(' TA \\l "') +
    instr('Baldwin v. Alberti', '<w:i/>') +
    instr(', 58 Wn. 2d 243 ') +
    instr('(1961)', '<w:b/><w:u w:val="single"/>') +
    instr('" \\s "Baldwin" \\c 1 ') +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

  it('reads the formatted pieces of the field code, and none from a plain one', async () => {
    const body = `<w:p><w:r><w:t>See Baldwin</w:t></w:r>${formatted}</w:p><w:p><w:r><w:t>Plain</w:t></w:r>${field(CITE)}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const [first, second] = parsed.blocks
      .slice(0, 2)
      .map((b) => (b as { runs: Run[] }).runs.find((r) => r.taInstr !== undefined)!)
    expect(first.taInstr).toBe('\\l "Baldwin v. Alberti, 58 Wn. 2d 243 (1961)" \\s "Baldwin" \\c 1')
    expect(first.taRuns).toEqual([
      { text: 'Baldwin v. Alberti', italic: true },
      { text: ', 58 Wn. 2d 243 ' },
      { text: '(1961)', bold: true, underline: true },
    ])
    expect(second.taRuns).toBeUndefined()
  })

  it('writes them back as formatted field code runs, and reads the same again', async () => {
    const body = `<w:p><w:r><w:t>See Baldwin</w:t></w:r>${formatted}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const runs = (parsed.blocks[0] as { runs: Run[] }).runs
    const saved = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'Edited. ' }, ...runs] },
        docxIndex: 0,
      },
    ])
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain(
      '<w:r><w:rPr><w:i/></w:rPr><w:instrText xml:space="preserve">Baldwin v. Alberti</w:instrText></w:r>',
    )
    expect(xml).toContain(
      '<w:r><w:rPr><w:b/><w:u w:val="single"/></w:rPr><w:instrText xml:space="preserve">(1961)</w:instrText></w:r>',
    )
    const reparsed = (await parseDocx(saved)).blocks[0] as { runs: Run[] }
    const marker = reparsed.runs.find((r) => r.taInstr !== undefined)!
    expect(marker.taInstr).toBe(runs.find((r) => r.taInstr !== undefined)!.taInstr)
    expect(marker.taRuns).toEqual(runs.find((r) => r.taInstr !== undefined)!.taRuns)
  })

  it('a table line carries the formatted pieces, and they are read back as its title', async () => {
    const [heading, line] = generateFieldTableXml(
      [
        { text: 'Cases', style: 'TOAHeading' },
        {
          text: 'Baldwin v. Alberti (1961)',
          textRuns: [{ text: 'Baldwin v. Alberti', italic: true }, { text: ' (1961)' }],
          pageNo: 3,
          style: 'TableofAuthorities',
        },
      ],
      { instr: ' TOA \\h \\p ' },
    )
    expect(heading).toContain('>Cases</w:t>')
    expect(line).toContain(
      '<w:r><w:rPr><w:i/><w:noProof/></w:rPr><w:t xml:space="preserve">Baldwin v. Alberti</w:t></w:r>',
    )
    const parsed = await parseDocx(await buildDocx({ bodyXml: heading + line }))
    const shown = parsed.blocks[1].fieldDisplay!
    expect(shown.left).toBe('Baldwin v. Alberti (1961)')
    expect(shown.leftRuns).toEqual([
      { text: 'Baldwin v. Alberti', italic: true },
      { text: ' (1961)' },
    ])
    expect(shown.right).toBe('3')
    expect(parsed.blocks[0].fieldDisplay?.leftRuns).toBeUndefined()
  })
})

describe('generateTableOfAuthoritiesFieldXml', () => {
  it('emits a TOA field for the category: its heading, then the citations with their pages', () => {
    const frags = generateTableOfAuthoritiesFieldXml(
      'Cases',
      [{ text: 'Baldwin v. Alberti', pageNo: 3 }, { text: 'Dravo <Corp.> & Co.' }],
      1,
    )
    expect(frags).toHaveLength(3)
    expect(frags[0]).toContain(' TOA \\h \\c "1" \\p ')
    expect(frags[0]).toContain('<w:pStyle w:val="TOAHeading"/>')
    expect(frags[0]).toContain('>Cases<')
    expect(frags[1]).toContain('<w:pStyle w:val="TableofAuthorities"/>')
    expect(frags[1]).toContain('<w:t>3</w:t>')
    expect(frags[2]).toContain('Dravo &lt;Corp.&gt; &amp; Co.')
    expect(frags[2]).toContain('w:fldCharType="end"')
    expect(generateTableOfAuthoritiesFieldXml('Cases', [], 1)).toEqual([])
  })

  it('reads its own output back as a listing: the heading, then each citation and its page', async () => {
    const frags = generateTableOfAuthoritiesFieldXml(
      'Cases',
      [{ text: 'Baldwin v. Alberti', pageNo: 3 }],
      1,
    )
    const parsed = await parseDocx(await buildDocx({ bodyXml: frags.join('') }))
    const shown = parsed.blocks.map(
      (b) =>
        (b as { fieldDisplay?: { kind: string; left: string; right?: string; level?: number } })
          .fieldDisplay,
    )
    // the entries one level in, as a newly inserted table shows them
    expect(shown.slice(0, 2).map((d) => [d?.kind, d?.left, d?.right, d?.level])).toEqual([
      ['tocLine', 'Cases', '', 1],
      ['tocLine', 'Baldwin v. Alberti', '3', 2],
    ])
  })
})
