import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  DEFAULT_MATH_PR,
  latexToOmml,
  mathParagraphXml,
  ommlToMathML,
  parseDocx,
  saveDocx,
} from '../src/index'
import type { MathPrEdit } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const M_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'

const WORD_MATH_PR =
  '<m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/>' +
  '<m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/>' +
  '<m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/>' +
  '<m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>'

const settingsPart = (inner: string, ns = `${W_NS} ${M_NS}`) =>
  `${XML_DECL}<w:settings ${ns}>${inner}</w:settings>`

async function docWith(settingsXml: string | null, bodyXml = BODY) {
  return buildDocx({
    bodyXml,
    extraParts: settingsXml
      ? [{ path: 'word/settings.xml', xml: settingsXml, contentType: SETTINGS_TYPE }]
      : [],
  })
}

async function saveMath(settingsXml: string | null, mathPr: MathPrEdit) {
  const parsed = await parseDocx(await docWith(settingsXml))
  const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], { mathPr })
  const zip = await JSZip.loadAsync(saved)
  return { saved, xml: await zip.file('word/settings.xml')!.async('string') }
}

const edit = (over: Partial<MathPrEdit> = {}): MathPrEdit => ({
  dispDef: true,
  defJc: 'centerGroup',
  lMarginTwips: 0,
  rMarginTwips: 0,
  intLim: 'subSup',
  naryLim: 'undOvr',
  ...over,
})

describe('reading m:mathPr', () => {
  it('reads every setting of the element Word writes', async () => {
    const doc = await parseDocx(
      await docWith(
        settingsPart(
          '<m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="after"/>' +
            '<m:brkBinSub m:val="-+"/><m:smallFrac/><m:dispDef m:val="0"/>' +
            '<m:lMargin m:val="720"/><m:rMargin m:val="360"/><m:defJc m:val="right"/>' +
            '<m:wrapRight/><m:intLim m:val="undOvr"/><m:naryLim m:val="subSup"/></m:mathPr>',
        ),
      ),
    )
    expect(doc.mathPr).toEqual({
      dispDef: false,
      defJc: 'right',
      lMarginTwips: 720,
      rMarginTwips: 360,
      intLim: 'undOvr',
      naryLim: 'subSup',
      mathFont: 'Cambria Math',
      brkBin: 'after',
      brkBinSub: '-+',
      smallFrac: true,
      wrapRight: true,
    })
  })

  it('fills Word defaults for the children an m:mathPr leaves out', async () => {
    const doc = await parseDocx(await docWith(settingsPart('<m:mathPr/>')))
    expect(doc.mathPr).toEqual(DEFAULT_MATH_PR)
  })

  it('reports nothing for a part without m:mathPr, and for no settings part', async () => {
    expect((await parseDocx(await docWith(settingsPart('<w:zoom w:percent="100"/>')))).mathPr).toBe(
      undefined,
    )
    expect((await parseDocx(await docWith(null))).mathPr).toBe(undefined)
  })

  it('follows the prefix the part binds to the math namespace', async () => {
    const xml = settingsPart(
      '<mm:mathPr><mm:defJc mm:val="left"/></mm:mathPr>',
      `${W_NS} ${M_NS.replace('xmlns:m', 'xmlns:mm')}`,
    )
    expect((await parseDocx(await docWith(xml))).mathPr?.defJc).toBe('left')
  })
})

describe('writing m:mathPr', () => {
  it('rewrites only the changed children and keeps the rest byte for byte', async () => {
    const { xml, saved } = await saveMath(
      settingsPart(`<w:zoom w:percent="100"/>${WORD_MATH_PR}<w:themeFontLang w:val="en-US"/>`),
      edit({ defJc: 'left', lMarginTwips: 720 }),
    )
    expect(xml).toBe(
      settingsPart(
        '<w:zoom w:percent="100"/>' +
          WORD_MATH_PR.replace('<m:lMargin m:val="0"/>', '<m:lMargin m:val="720"/>').replace(
            '<m:defJc m:val="centerGroup"/>',
            '<m:defJc m:val="left"/>',
          ) +
          '<w:themeFontLang w:val="en-US"/>',
      ),
    )
    expect((await parseDocx(saved)).mathPr).toMatchObject({ defJc: 'left', lMarginTwips: 720 })
  })

  it('writes the math font and the nested-fraction size in place', async () => {
    const { xml } = await saveMath(settingsPart(WORD_MATH_PR), {
      ...DEFAULT_MATH_PR,
      mathFont: 'STIX Two Math',
      smallFrac: true,
    })
    expect(xml).toContain('<m:mathFont m:val="STIX Two Math"/><m:brkBin m:val="before"/>')
    expect(xml).toContain('<m:brkBinSub m:val="--"/><m:smallFrac/><m:dispDef/>')
    const back = await saveMath(xml, { ...DEFAULT_MATH_PR, mathFont: 'STIX Two Math' })
    expect(back.xml).toContain('<m:smallFrac m:val="0"/>')
  })

  it('writes the operator break rules in place', async () => {
    const { xml, saved } = await saveMath(
      settingsPart(WORD_MATH_PR),
      edit({ brkBin: 'repeat', brkBinSub: '+-' }),
    )
    expect(xml).toBe(
      settingsPart(
        WORD_MATH_PR.replace('<m:brkBin m:val="before"/>', '<m:brkBin m:val="repeat"/>').replace(
          '<m:brkBinSub m:val="--"/>',
          '<m:brkBinSub m:val="+-"/>',
        ),
      ),
    )
    expect((await parseDocx(saved)).mathPr).toMatchObject({ brkBin: 'repeat', brkBinSub: '+-' })
  })

  it('swaps m:wrapIndent and m:wrapRight, which are one choice', async () => {
    const right = await saveMath(settingsPart(WORD_MATH_PR), edit({ wrapRight: true }))
    expect(right.xml).toBe(
      settingsPart(WORD_MATH_PR.replace('<m:wrapIndent m:val="1440"/>', '<m:wrapRight/>')),
    )
    expect((await parseDocx(right.saved)).mathPr?.wrapRight).toBe(true)
    const indent = await saveMath(right.xml, edit({ wrapRight: false, wrapIndentTwips: 720 }))
    expect(indent.xml).toBe(
      settingsPart(
        WORD_MATH_PR.replace('<m:wrapIndent m:val="1440"/>', '<m:wrapIndent m:val="720"/>'),
      ),
    )
  })

  it('writes no break or wrap child for the values Word assumes', async () => {
    const source = settingsPart('<m:mathPr><m:defJc m:val="left"/></m:mathPr>')
    const same = await saveMath(
      source,
      edit({ defJc: 'left', brkBin: 'before', brkBinSub: '--', wrapIndentTwips: 1440 }),
    )
    expect(same.xml).toBe(source)
    const after = await saveMath(source, edit({ defJc: 'left', brkBin: 'after' }))
    expect(after.xml).toBe(
      settingsPart('<m:mathPr><m:brkBin m:val="after"/><m:defJc m:val="left"/></m:mathPr>'),
    )
    const indent = await saveMath(source, edit({ defJc: 'left', wrapIndentTwips: 360 }))
    expect(indent.xml).toBe(
      settingsPart('<m:mathPr><m:defJc m:val="left"/><m:wrapIndent m:val="360"/></m:mathPr>'),
    )
  })

  it('leaves the part untouched when nothing differs from what it says', async () => {
    const source = settingsPart(`<w:zoom w:percent="100"/>${WORD_MATH_PR}`)
    expect((await saveMath(source, edit())).xml).toBe(source)
  })

  it('adds missing children at their CT_MathPr position', async () => {
    const { xml } = await saveMath(
      settingsPart(
        '<m:mathPr><m:mathFont m:val="Cambria Math"/><m:interSp m:val="20"/>' +
          '<m:wrapIndent m:val="1440"/></m:mathPr>',
      ),
      edit({
        dispDef: false,
        defJc: 'right',
        lMarginTwips: 100,
        rMarginTwips: 200,
        intLim: 'undOvr',
        naryLim: 'subSup',
      }),
    )
    expect(xml).toContain(
      '<m:mathPr><m:mathFont m:val="Cambria Math"/><m:dispDef m:val="0"/>' +
        '<m:lMargin m:val="100"/><m:rMargin m:val="200"/><m:defJc m:val="right"/>' +
        '<m:interSp m:val="20"/><m:wrapIndent m:val="1440"/><m:intLim m:val="undOvr"/>' +
        '<m:naryLim m:val="subSup"/></m:mathPr>',
    )
  })

  it('turns an explicit m:dispDef off back on in place', async () => {
    const { xml } = await saveMath(
      settingsPart('<m:mathPr><m:dispDef m:val="off"/><m:defJc m:val="left"/></m:mathPr>'),
      edit({ defJc: 'left' }),
    )
    expect(xml).toContain('<m:mathPr><m:dispDef/><m:defJc m:val="left"/></m:mathPr>')
  })

  it('expands a self-closed m:mathPr', async () => {
    const { xml } = await saveMath(settingsPart('<m:mathPr/>'), edit({ naryLim: 'subSup' }))
    expect(xml).toContain('<m:mathPr><m:naryLim m:val="subSup"/></m:mathPr>')
  })

  it('creates the element at its CT_Settings position and declares the namespace on the root', async () => {
    const { xml, saved } = await saveMath(
      settingsPart(
        '<w:zoom w:percent="100"/><w:compat/><w:rsids/><w:themeFontLang w:val="en-US"/>',
        W_NS,
      ),
      edit({ defJc: 'right' }),
    )
    expect(xml).toBe(
      `${XML_DECL}<w:settings ${M_NS} ${W_NS}><w:zoom w:percent="100"/><w:compat/><w:rsids/>` +
        '<m:mathPr><m:defJc m:val="right"/></m:mathPr><w:themeFontLang w:val="en-US"/></w:settings>',
    )
    expect((await parseDocx(saved)).mathPr?.defJc).toBe('right')
  })

  it('declares the namespace on the element when the root binds m: to something else', async () => {
    const { xml, saved } = await saveMath(
      settingsPart('<w:zoom w:percent="100"/>', `${W_NS} xmlns:m="urn:other"`),
      edit({ lMarginTwips: 240 }),
    )
    expect(xml).toContain(
      '<m:mathPr xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
        '<m:lMargin m:val="240"/></m:mathPr></w:settings>',
    )
    expect((await parseDocx(saved)).mathPr?.lMarginTwips).toBe(240)
  })

  it('writes with the prefix the part already uses', async () => {
    const { xml } = await saveMath(
      settingsPart(
        '<mm:mathPr><mm:defJc mm:val="left"/></mm:mathPr>',
        `${W_NS} ${M_NS.replace('xmlns:m', 'xmlns:mm')}`,
      ),
      edit({ defJc: 'center' }),
    )
    expect(xml).toContain('<mm:mathPr><mm:defJc mm:val="center"/></mm:mathPr>')
  })

  it('creates the settings part of a document that has none', async () => {
    const { xml, saved } = await saveMath(null, edit({ defJc: 'left' }))
    expect(xml).toContain(
      `<w:settings ${M_NS} ${W_NS}><m:mathPr><m:defJc m:val="left"/></m:mathPr>`,
    )
    expect((await parseDocx(saved)).mathPr?.defJc).toBe('left')
  })
})

describe('display equations and the document defaults', () => {
  const equation = (pr: string) =>
    `<w:p><m:oMathPara>${pr}<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></w:p>`

  it('marks an m:oMathPara without m:jc as following the default', async () => {
    const doc = await parseDocx(
      await docWith(
        null,
        equation('') +
          equation('<m:oMathParaPr><m:jc m:val="right"/></m:oMathParaPr>') +
          '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></w:p>',
      ),
    )
    const [byDefault, own, inline] = doc.blocks.map((b) => b.formulaDisplay)
    expect(byDefault).toMatchObject({ mathPara: true, jcDefault: true, paraAlign: 'left' })
    expect(own).toMatchObject({ mathPara: true, align: 'right' })
    expect(own?.jcDefault).toBe(undefined)
    expect(inline?.mathPara).toBe(undefined)
    expect(inline?.align).toBe('center')
  })

  it('an equation built without an alignment carries no m:jc', () => {
    expect(mathParagraphXml('<m:r><m:t>x</m:t></m:r>', 'default')).toBe(
      '<w:p><m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></w:p>',
    )
  })
})

describe('n-ary limits and the document defaults', () => {
  const nary = (chr: string, limLoc = '') =>
    `<m:oMath><m:nary><m:naryPr><m:chr m:val="${chr}"/>${limLoc}</m:naryPr>` +
    '<m:sub><m:r><m:t>a</m:t></m:r></m:sub><m:sup><m:r><m:t>b</m:t></m:r></m:sup>' +
    '<m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary></m:oMath>'

  it('places limits by m:intLim / m:naryLim when the operator has no m:limLoc', () => {
    expect(ommlToMathML(nary('∫'))).toContain('<msubsup>')
    expect(ommlToMathML(nary('∫'), { intLim: 'undOvr' })).toContain('<munderover>')
    expect(ommlToMathML(nary('∑'))).toContain('<munderover>')
    expect(ommlToMathML(nary('∑'), { naryLim: 'subSup' })).toContain('<msubsup>')
    expect(ommlToMathML(nary('∑'), { intLim: 'undOvr' })).toContain('<munderover>')
    expect(ommlToMathML(nary('∬'), { naryLim: 'undOvr' })).toContain('<msubsup>')
  })

  it("keeps an operator's own m:limLoc over the default", () => {
    expect(ommlToMathML(nary('∑', '<m:limLoc m:val="undOvr"/>'), { naryLim: 'subSup' })).toContain(
      '<munderover>',
    )
    expect(ommlToMathML(nary('∫', '<m:limLoc m:val="subSup"/>'), { intLim: 'undOvr' })).toContain(
      '<msubsup>',
    )
  })

  it('does not carry one conversion’s defaults into the next', () => {
    ommlToMathML(nary('∑'), { naryLim: 'subSup' })
    expect(ommlToMathML(nary('∑'))).toContain('<munderover>')
  })

  it('applies the defaults at parse time', async () => {
    const doc = await parseDocx(
      await docWith(
        settingsPart('<m:mathPr><m:naryLim m:val="subSup"/></m:mathPr>'),
        `<w:p><m:oMathPara>${nary('∑')}</m:oMathPara></w:p>`,
      ),
    )
    expect(doc.blocks[0].formulaDisplay?.mathml).toContain('<msubsup>')
  })

  it('a new equation writes the placement the document asks for', () => {
    expect(latexToOmml('\\sum_{i}^{n}{x}')).toContain('<m:limLoc m:val="undOvr"/>')
    expect(latexToOmml('\\sum_{i}^{n}{x}', { naryLim: 'subSup' })).toContain(
      '<m:limLoc m:val="subSup"/>',
    )
    expect(latexToOmml('\\int_{a}^{b}{x}', { intLim: 'undOvr' })).toContain(
      '<m:limLoc m:val="undOvr"/>',
    )
    expect(latexToOmml('\\sqrt[\\sum_{i}{x}]{y}', { naryLim: 'subSup' })).toContain(
      '<m:limLoc m:val="subSup"/>',
    )
  })
})
