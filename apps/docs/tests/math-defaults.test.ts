import { describe, expect, it } from 'vitest'
import { DEFAULT_MATH_PR } from '@genoffice/docx-engine'
import {
  mathDefaultsCss,
  mathLimits,
  ommlFollowsLimitDefaults,
  setMathLimits,
} from '../src/renderer/editor/math-defaults'
import { equationBlockJson } from '../src/renderer/editor/equation'

describe('document equation defaults', () => {
  it('centres as a group with no margins for a document without m:mathPr', () => {
    expect(mathDefaultsCss().split('\n')[0]).toBe(
      '.doc-page { --doc-math-jc:center;--doc-math-line-l:0;--doc-math-line-r:0;' +
        '--doc-math-lm:0pt;--doc-math-rm:0pt }',
    )
  })

  it('keeps a nested fraction full size in a display equation unless m:smallFrac is on', () => {
    const hold =
      ".doc-page :is(math[display='block'], math[displaystyle='true']) mfrac > * { math-depth:add(0) }"
    expect(mathDefaultsCss()).toContain(hold)
    expect(mathDefaultsCss({ ...DEFAULT_MATH_PR, smallFrac: false })).toContain(hold)
    expect(mathDefaultsCss({ ...DEFAULT_MATH_PR, smallFrac: true })).not.toContain('math-depth')
    expect(mathDefaultsCss({ ...DEFAULT_MATH_PR, dispDef: false })).toContain(hold)
  })

  it("leads the math font stack with the document's m:mathFont when it is not Word's own", () => {
    expect(mathDefaultsCss()).not.toContain('--doc-math-font')
    expect(mathDefaultsCss({ ...DEFAULT_MATH_PR, mathFont: 'Cambria Math' })).not.toContain(
      '--doc-math-font',
    )
    expect(mathDefaultsCss({ ...DEFAULT_MATH_PR, mathFont: 'STIX Two Math' })).toContain(
      ".doc-page { --doc-math-font:'STIX Two Math' }",
    )
    expect(mathDefaultsCss({ ...DEFAULT_MATH_PR, mathFont: "X'; } body { color:red" })).toContain(
      "--doc-math-font:'X  body  color:red' }",
    )
  })

  it('writes the justification and the margins of the document', () => {
    const css = mathDefaultsCss({
      ...DEFAULT_MATH_PR,
      defJc: 'right',
      lMarginTwips: 720,
      rMarginTwips: 360,
    })
    expect(css).toContain('--doc-math-jc:right;--doc-math-line-l:auto;--doc-math-line-r:0')
    expect(css).toContain('--doc-math-lm:36pt;--doc-math-rm:18pt')
    expect(mathDefaultsCss({ ...DEFAULT_MATH_PR, defJc: 'center' })).toContain(
      '--doc-math-line-l:auto;--doc-math-line-r:auto',
    )
  })

  it('with m:dispDef off drops the margins and hands the equation to its paragraph', () => {
    const css = mathDefaultsCss({ ...DEFAULT_MATH_PR, dispDef: false, lMarginTwips: 720 })
    expect(css).not.toContain('--doc-math-lm')
    expect(css).toContain("[data-math-para='default']")
    expect(css).toContain('--doc-math-jc:var(--doc-math-para-jc, left)')
  })

  it('tells an n-ary operator that leaves its limits to the document', () => {
    const nary = (pr: string) => `<m:oMath><m:nary><m:naryPr>${pr}</m:naryPr></m:nary></m:oMath>`
    expect(ommlFollowsLimitDefaults(nary(''))).toBe(true)
    expect(ommlFollowsLimitDefaults(nary('<m:limLoc m:val="undOvr"/>'))).toBe(false)
    expect(ommlFollowsLimitDefaults('<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>')).toBe(false)
  })

  it('a new display equation follows the document: no m:jc, limits as the document places them', () => {
    setMathLimits({ ...DEFAULT_MATH_PR, naryLim: 'subSup' })
    expect(mathLimits()).toEqual({ intLim: 'subSup', naryLim: 'subSup' })
    const block = equationBlockJson('\\sum_{i}^{n}{x}')
    const attrs = block.attrs as { genXml: string; formulaDisplay: Record<string, unknown> }
    expect(attrs.genXml).not.toContain('<m:jc')
    expect(attrs.genXml).toContain('<m:limLoc m:val="subSup"/>')
    expect(attrs.formulaDisplay).toMatchObject({ mathPara: true, jcDefault: true })
    expect(String(attrs.formulaDisplay.mathml)).toContain('<msubsup>')
    setMathLimits(undefined)
    expect(mathLimits()).toEqual({})
  })
})
