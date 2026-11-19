/**
 * The textbox numPr write-back: a list toggled inside the sub-editor reaches
 * the paragraph's w:numPr (set, replace, clear), placed per CT_PPr order after
 * w:pStyle, and the surrounding pPr content survives byte-exactly.
 */
import { describe, expect, it } from 'vitest'
import { patchTextboxParas } from '../src/generate'

const CONTENT = (pPr: string): string =>
  `<w:txbxContent><w:p>${pPr}<w:r><w:t>Item</w:t></w:r></w:p></w:txbxContent>`

const NUM = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>'

describe('patchTextboxParas numPr', () => {
  it('inserts numPr after w:pStyle, before spacing and jc', () => {
    const xml = CONTENT(
      '<w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:after="0"/><w:jc w:val="left"/></w:pPr>',
    )
    const out = patchTextboxParas(xml, {
      byIndex: [[{ runs: [], align: 'left', numPr: { numId: '7', ilvl: 0 } }]],
    })
    const pPr = /<w:pPr>[\s\S]*<\/w:pPr>/.exec(out)![0]!
    expect(pPr.indexOf('<w:pStyle')).toBeLessThan(pPr.indexOf(NUM))
    expect(pPr.indexOf(NUM)).toBeLessThan(pPr.indexOf('<w:spacing'))
    expect(pPr.indexOf(NUM)).toBeLessThan(pPr.indexOf('<w:jc'))
  })

  it('inserts numPr right after the pPr opening when no pStyle exists', () => {
    const xml = CONTENT('<w:pPr><w:jc w:val="center"/></w:pPr>')
    const out = patchTextboxParas(xml, {
      byIndex: [
        [{ runs: [], align: 'center', numPr: { numId: '3', ilvl: '0' as unknown as number } }],
      ],
    })
    expect(out).toContain(
      '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr><w:jc w:val="center"/></w:pPr>',
    )
  })

  it('creates a pPr for a paragraph that had none', () => {
    const xml = CONTENT('')
    const out = patchTextboxParas(xml, {
      byIndex: [
        [
          {
            runs: [{ text: 'Item' }],
            align: null,
            numPr: { numId: '7', ilvl: 0 },
          },
        ],
      ],
    })
    expect(out).toContain(`${NUM}</w:pPr><w:r>`)
  })

  it('replaces an existing numPr and clears it on null', () => {
    const old = '<w:numPr><w:ilvl w:val="2"/><w:numId w:val="4"/></w:numPr>'
    const xml = CONTENT(
      `<w:pPr><w:pStyle w:val="P"/><w:numPr><w:ilvl w:val="2"/><w:numId w:val="4"/></w:numPr><w:spacing w:before="10"/></w:pPr>`,
    )
    const out = patchTextboxParas(xml, {
      byIndex: [[{ runs: [], align: null, numPr: { numId: '7', ilvl: 0 } }]],
    })
    expect(out).not.toContain(old)
    expect(out).toContain('<w:pStyle w:val="P"/>')
    expect(out).toContain('<w:spacing w:before="10"/>')

    const cleared = patchTextboxParas(out, {
      byIndex: [[{ runs: [], align: null, numPr: null }]],
    })
    expect(cleared).not.toContain('<w:numPr>')
  })
})

describe('patchTextboxParas indentLeft', () => {
  it('adds w:ind before w:jc and keeps the paragraph-mark rPr last', () => {
    const xml = CONTENT('<w:pPr><w:jc w:val="left"/><w:rPr><w:b/></w:rPr></w:pPr>')
    const out = patchTextboxParas(xml, {
      byIndex: [[{ runs: [], align: 'left', indentLeft: 720 }]],
    })
    expect(out).toContain(
      '<w:pPr><w:ind w:left="720" w:leftChars="0"/><w:jc w:val="left"/><w:rPr><w:b/></w:rPr></w:pPr>',
    )
  })

  it('replaces w:left on an existing w:ind and keeps its hanging indent', () => {
    const xml = CONTENT('<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>')
    const out = patchTextboxParas(xml, { byIndex: [[{ runs: [], indentLeft: 1440 }]] })
    expect(out).toContain('<w:ind w:hanging="360" w:left="1440" w:leftChars="0"/>')
  })

  it('clears w:left on null and drops a w:ind left empty', () => {
    const xml = CONTENT('<w:pPr><w:ind w:left="720"/><w:jc w:val="center"/></w:pPr>')
    const out = patchTextboxParas(xml, {
      byIndex: [[{ runs: [], align: 'center', indentLeft: null }]],
    })
    expect(out).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>')
    expect(out).not.toContain('<w:ind')
  })

  it('replaces an own character-unit left indent instead of leaving it to win', () => {
    const xml = CONTENT('<w:pPr><w:ind w:leftChars="200" w:left="420"/></w:pPr>')
    const out = patchTextboxParas(xml, { byIndex: [[{ runs: [], indentLeft: 720 }]] })
    expect(out).toContain('<w:ind w:left="720" w:leftChars="0"/>')
    const cleared = patchTextboxParas(xml, { byIndex: [[{ runs: [], indentLeft: null }]] })
    expect(cleared).not.toContain('<w:ind')
  })

  it('leaves the original w:ind alone when the patch does not mention indentLeft', () => {
    const xml = CONTENT('<w:pPr><w:ind w:left="720"/></w:pPr>')
    const out = patchTextboxParas(xml, { byIndex: [[{ runs: [], align: null }]] })
    expect(out).toContain('<w:ind w:left="720"/>')
  })
})
