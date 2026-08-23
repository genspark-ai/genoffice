import { describe, expect, it } from 'vitest'
import {
  latexToOmml,
  mathParagraphXml,
  mathTokensOf,
  ommlFragmentsOf,
  ommlToLatex,
  ommlToMathML,
  parseDocx,
  patchMathTokens,
  saveDocx,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const FRACTION_OMATH =
  '<m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>' +
  '<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>'

describe('ommlToMathML', () => {
  it('converts fractions', () => {
    const mathml = ommlToMathML(FRACTION_OMATH)
    expect(mathml).toContain('<mfrac>')
    expect(mathml).toContain('<mi>a</mi>')
    expect(mathml).toContain('<mi>b</mi>')
  })

  it('converts scripts, radicals and delimiters', () => {
    const omml =
      '<m:oMath>' +
      '<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>' +
      '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>y</m:t></m:r></m:e></m:rad>' +
      '<m:d><m:e><m:r><m:t>z</m:t></m:r></m:e></m:d>' +
      '</m:oMath>'
    const mathml = ommlToMathML(omml)
    expect(mathml).toContain('<msup>')
    expect(mathml).toContain('<mn>2</mn>')
    expect(mathml).toContain('<msqrt>')
    expect(mathml).toContain('<mo stretchy="true">(</mo>')
  })

  it('converts n-ary operators with under/over limits', () => {
    const omml =
      '<m:oMath><m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr>' +
      '<m:sub><m:r><m:t>k=0</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup>' +
      '<m:e><m:r><m:t>k</m:t></m:r></m:e></m:nary></m:oMath>'
    const mathml = ommlToMathML(omml)
    expect(mathml).toContain('<munderover>')
    expect(mathml).toContain('∑')
  })

  it('converts matrices', () => {
    const omml =
      '<m:oMath><m:m>' +
      '<m:mr><m:e><m:r><m:t>1</m:t></m:r></m:e><m:e><m:r><m:t>0</m:t></m:r></m:e></m:mr>' +
      '<m:mr><m:e><m:r><m:t>0</m:t></m:r></m:e><m:e><m:r><m:t>1</m:t></m:r></m:e></m:mr>' +
      '</m:m></m:oMath>'
    const mathml = ommlToMathML(omml)
    expect(mathml.match(/<mtr>/g)).toHaveLength(2)
    expect(mathml.match(/<mtd>/g)).toHaveLength(4)
  })

  it('classifies run text into mn / mi / mo tokens', () => {
    const omml = '<m:oMath><m:r><m:t>2x+1</m:t></m:r></m:oMath>'
    const mathml = ommlToMathML(omml)
    expect(mathml).toContain('<mn>2</mn>')
    expect(mathml).toContain('<mi>x</mi>')
    expect(mathml).toContain('<mo>+</mo>')
    expect(mathml).toContain('<mn>1</mn>')
  })

  it('handles multiple oMath fragments in one paragraph', () => {
    const xml = `<w:p>${FRACTION_OMATH}${FRACTION_OMATH}</w:p>`
    expect(ommlFragmentsOf(xml)).toHaveLength(2)
    expect(ommlToMathML(xml).match(/<math /g)).toHaveLength(2)
  })

  it('unwraps oMathPara', () => {
    const xml = `<w:p><m:oMathPara>${FRACTION_OMATH}</m:oMathPara></w:p>`
    expect(ommlFragmentsOf(xml)).toHaveLength(1)
    expect(ommlToMathML(xml)).toContain('<mfrac>')
  })
})

describe('latexToOmml', () => {
  it('builds fractions, scripts and radicals', () => {
    const omml = latexToOmml('x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}')
    expect(omml).toContain('<m:f><m:num>')
    expect(omml).toContain('<m:rad>')
    expect(omml).toContain('±')
    const mathml = ommlToMathML(`<m:oMath>${omml}</m:oMath>`)
    expect(mathml).toContain('<mfrac>')
    expect(mathml).toContain('<msqrt>')
    expect(mathml).toContain('<msup>')
  })

  it('builds n-ary sums with limits and binomials', () => {
    const omml = latexToOmml('\\sum_{k=0}^{n} \\binom{n}{k} x^k')
    expect(omml).toContain('<m:nary>')
    expect(omml).toContain('m:val="∑"')
    expect(omml).toContain('m:val="noBar"')
  })

  it('builds matrices with delimiters', () => {
    const omml = latexToOmml('\\begin{pmatrix} 1 & 0 \\\\ 0 & 1 \\end{pmatrix}')
    expect(omml).toContain('<m:m>')
    expect(omml.match(/<m:mr>/g)).toHaveLength(2)
    expect(omml).toContain('m:begChr m:val="("')
  })

  it('maps greek letters and functions', () => {
    const omml = latexToOmml('A = \\pi r^2, \\sin \\alpha')
    expect(omml).toContain('π')
    expect(omml).toContain('α')
    expect(omml).toContain('<m:sty m:val="p"/>')
    expect(omml).toContain('sin')
  })

  it('binds a script to the last character of a text run', () => {
    const omml = latexToOmml('ab^2')
    // the base of the superscript must be "b", not "ab"
    expect(omml).toContain('<m:t xml:space="preserve">a</m:t>')
    expect(omml).toMatch(/<m:sSup><m:e><m:r><m:t[^>]*>b<\/m:t>/)
  })

  it('supports \\left \\right stretchy delimiters', () => {
    const omml = latexToOmml('\\left( \\frac{a}{b} \\right)')
    expect(omml).toContain('<m:d>')
    expect(omml).toContain('<m:f>')
  })

  it('rejects unsupported commands with a helpful error', () => {
    expect(() => latexToOmml('\\unknowncmd{x}')).toThrow(/Unsupported command/)
    expect(() => latexToOmml('{unclosed')).toThrow(/Missing matching/)
  })

  it('token edits round-trip through patchMathTokens', () => {
    const omml = `<m:oMath>${latexToOmml('\\frac{a}{b}')}</m:oMath>`
    const tokens = mathTokensOf(omml)
    expect(tokens).toEqual(['a', 'b'])
    const patched = patchMathTokens(omml, ['x', 'y'])
    expect(mathTokensOf(patched)).toEqual(['x', 'y'])
    expect(ommlToMathML(patched)).toContain('<mi>x</mi>')
  })
})

describe('ommlToLatex', () => {
  const CASES = [
    'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}',
    '(x+a)^n = \\sum_{k=0}^{n} \\binom{n}{k} x^k a^{n-k}',
    'f(x) = a_0 + \\sum_{n=1}^{\\infty} \\left( a_n \\cos \\frac{n\\pi x}{L} + b_n \\sin \\frac{n\\pi x}{L} \\right)',
    'a^2 + b^2 = c^2',
    'e^x = 1 + \\frac{x}{1!} + \\frac{x^2}{2!} + \\cdots',
    '\\begin{pmatrix} 1 & 0 \\\\ 0 & 1 \\end{pmatrix}',
    '\\begin{cases} x & x > 0 \\\\ -x & x \\le 0 \\end{cases}',
    '\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1',
    '\\sqrt[3]{x} + \\hat{y} + \\overline{z}',
    '\\int_{0}^{1} x^2 dx',
    '\\underbrace{a + b}',
    '\\text{速度} = \\frac{s}{t}',
  ]

  it.each(CASES)('round-trips through latexToOmml: %s', (latex) => {
    const omml = `<m:oMath>${latexToOmml(latex)}</m:oMath>`
    const decompiled = ommlToLatex(omml)
    expect(decompiled).toBeTruthy()
    // semantic equivalence: recompiling the decompiled LaTeX renders the same
    const omml2 = `<m:oMath>${latexToOmml(decompiled!)}</m:oMath>`
    expect(ommlToMathML(omml2)).toBe(ommlToMathML(omml))
    // and the visible tokens survive
    expect(mathTokensOf(omml2).join('')).toBe(mathTokensOf(omml).join(''))
  })

  it('decompiles Word-authored OMML (attrs and property bags present)', () => {
    const omml =
      '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
      '<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr>' +
      '<m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>x</m:t></m:r></m:e>' +
      '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>' +
      '<m:r><m:t>+1=0</m:t></m:r></m:oMath>'
    expect(ommlToLatex(omml)).toBe('{x}^{2}+1=0')
  })

  it('returns null for structures outside the subset', () => {
    expect(
      ommlToLatex(
        '<m:oMath><m:sPre><m:sub><m:r><m:t>a</m:t></m:r></m:sub><m:e><m:r><m:t>X</m:t></m:r></m:e></m:sPre></m:oMath>',
      ),
    ).toBeNull()
    expect(
      ommlToLatex(`<m:oMath>${latexToOmml('a')}</m:oMath><m:oMath>${latexToOmml('b')}</m:oMath>`),
    ).toBeNull()
  })

  it('escapes parser-special characters', () => {
    const omml = '<m:oMath><m:r><m:t>100%_x</m:t></m:r></m:oMath>'
    const latex = ommlToLatex(omml)
    expect(latex).toBeTruthy()
    const round = `<m:oMath>${latexToOmml(latex!)}</m:oMath>`
    expect(mathTokensOf(round).join('')).toBe('100%_x')
  })
})

describe('formula parse + save integration', () => {
  it('parseDocx fills mathml and omml for display equations', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: `<w:p>${FRACTION_OMATH}</w:p>` }))
    const formula = parsed.blocks[0].formulaDisplay
    expect(formula?.tokens).toEqual(['a', 'b'])
    expect(formula?.mathml).toContain('<mfrac>')
    expect(formula?.omml).toContain('<m:oMath>')
  })

  it('parses paragraphs mixing math and text into editable runs with math runs', async () => {
    const mixed = `<w:p><w:r><w:t xml:space="preserve">see </w:t></w:r>${FRACTION_OMATH}<w:r><w:t xml:space="preserve"> here</w:t></w:r></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: mixed }))
    const block = parsed.blocks[0]
    expect(block.type).toBe('paragraph')
    expect(block.runs?.map((r) => r.text)).toEqual(['see ', 'ab', ' here'])
    expect(block.runs?.[1].math?.omml).toBe(FRACTION_OMATH)
  })

  it('a mixed paragraph regenerates with the math run emitted verbatim', async () => {
    const mixed = `<w:p><w:r><w:t xml:space="preserve">see </w:t></w:r>${FRACTION_OMATH}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: mixed }))
    const block = parsed.blocks[0]
    const bytes = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: 'edited ' }, ...block.runs!.filter((r) => r.math)],
        },
      },
    ])
    const reparsed = await parseDocx(bytes)
    const rb = reparsed.blocks[0]
    expect(rb.type).toBe('paragraph')
    expect(rb.runs?.map((r) => r.text)).toEqual(['edited ', 'ab'])
    expect(rb.runs?.[1].math?.omml).toBe(FRACTION_OMATH)
  })

  it('display equations (oMathPara) still parse as protected formula blocks', async () => {
    const display = `<w:p><m:oMathPara>${FRACTION_OMATH}</m:oMathPara></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: display }))
    expect(parsed.blocks[0].type).toBe('passthrough')
    expect(parsed.blocks[0].formulaDisplay?.mathml).toContain('<mfrac>')
  })

  it('an inserted equation paragraph survives save and reparse', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>hi</w:t></w:r></w:p>' }),
    )
    const omml = latexToOmml('a^2 + b^2 = c^2')
    const bytes = await saveDocx(parsed, [
      { kind: 'original', docxIndex: parsed.blocks[0].docxIndex! },
      { kind: 'xml', xml: mathParagraphXml(omml) },
    ])
    const reparsed = await parseDocx(bytes)
    const formula = reparsed.blocks[1].formulaDisplay
    expect(formula?.mathml).toContain('<msup>')
    expect(formula?.tokens.join('')).toContain('=')
  })

  it('declares xmlns:m when the original document root lacks it', async () => {
    const JSZip = (await import('jszip')).default
    // rebuild a docx whose document.xml root omits xmlns:m (non-Word generators)
    const base = await buildDocx({ bodyXml: '<w:p><w:r><w:t>hi</w:t></w:r></w:p>' })
    const baseZip = await JSZip.loadAsync(base)
    const baseDocXml = await baseZip.file('word/document.xml')!.async('string')
    expect(baseDocXml).toContain('xmlns:m=')
    baseZip.file(
      'word/document.xml',
      baseDocXml.replace(
        ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
        '',
      ),
    )
    const stripped = await parseDocx(await baseZip.generateAsync({ type: 'uint8array' }))

    const omml = latexToOmml('E = mc^2')
    const bytes = await saveDocx(stripped, [
      { kind: 'original', docxIndex: stripped.blocks[0].docxIndex! },
      { kind: 'xml', xml: mathParagraphXml(omml) },
    ])
    const outZip = await JSZip.loadAsync(bytes)
    const docXml = await outZip.file('word/document.xml')!.async('string')
    expect(/<w:document[^>]*xmlns:m=/.test(docXml)).toBe(true)
    // reparse must still see the formula
    const reparsed = await parseDocx(bytes)
    expect(reparsed.blocks[1].formulaDisplay?.tokens.join('')).toContain('=')
  })
})
