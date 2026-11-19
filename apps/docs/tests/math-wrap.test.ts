import { describe, expect, it } from 'vitest'
import { DEFAULT_MATH_PR } from '@genoffice/docx-engine'
import {
  breakCandidates,
  mathWrapRulesOf,
  planWrapLines,
  wrapMathLine,
  type MathWrapRules,
} from '../src/renderer/editor/math-wrap'

const TOKEN = /[a-z]|\d+|./g

/** `<math>` for a flat formula: letters are identifiers, digits numbers, the rest operators */
function mathOf(formula: string): Element {
  const body = (formula.match(TOKEN) ?? [])
    .map((tok) =>
      /[a-z]/.test(tok)
        ? `<mi>${tok}</mi>`
        : /\d/.test(tok)
          ? `<mn>${tok}</mn>`
          : `<mo>${tok}</mo>`,
    )
    .join('')
  const host = document.createElement('div')
  host.innerHTML = `<math display="block"><mrow>${body}</mrow></math>`
  return host.firstElementChild!
}

/** every child 10 px wide, side by side */
const measure = (el: Element) => {
  const i = Array.from(el.parentElement!.children).indexOf(el)
  return { left: i * 10, right: i * 10 + 10 }
}

const rules = (over: Partial<MathWrapRules> = {}): MathWrapRules => ({
  ...mathWrapRulesOf(DEFAULT_MATH_PR),
  ...over,
})

const wrap = (formula: string, avail: number, over: Partial<MathWrapRules> = {}, rest = avail) =>
  wrapMathLine(mathOf(formula), avail, rest, rules(over), true, measure)

const texts = (lines: Element[] | null) => lines?.map((l) => l.textContent)

describe('where an equation line may break', () => {
  const at = (formula: string) => {
    const children = Array.from(mathOf(formula).firstElementChild!.children)
    return breakCandidates(children).map((i) => children[i].textContent)
  }

  it('breaks at binary and relational operators between operands', () => {
    expect(at('a+b=c−d×e')).toEqual(['+', '=', '−', '×'])
  })

  it('keeps a sign with its operand and a bracketed group whole', () => {
    expect(at('−a=−b')).toEqual(['='])
    expect(at('f(a+b)=c,d')).toEqual(['='])
    expect(at('a+')).toEqual([])
  })
})

describe('wrapping an over-wide equation line', () => {
  it('reads Word defaults: break before the operator, indent one inch', () => {
    expect(mathWrapRulesOf(undefined)).toEqual({
      brkBin: 'before',
      brkBinSub: '--',
      wrapRight: false,
      wrapIndentPx: 96,
    })
    expect(
      mathWrapRulesOf({ ...DEFAULT_MATH_PR, brkBin: 'repeat', wrapIndentTwips: 720 }),
    ).toMatchObject({ brkBin: 'repeat', wrapIndentPx: 48 })
  })

  it('starts the next line with the operator, ends the line with it, or shows it on both', () => {
    // 9 children, 90 px: 50 px holds five
    expect(texts(wrap('a+b+c+d+e', 50))).toEqual(['a+b+c', '+d+e'])
    expect(texts(wrap('a+b+c+d+e', 50, { brkBin: 'after' }))).toEqual(['a+b+', 'c+d+e'])
    expect(texts(wrap('a+b+c+d+e', 50, { brkBin: 'repeat' }))).toEqual(['a+b+', '+c+d+', '+e'])
  })

  it('repeats a subtraction as the document asks', () => {
    const minus = '−'
    const lines = (brkBinSub: MathWrapRules['brkBinSub']) =>
      texts(wrap(`a+b${minus}c`, 40, { brkBin: 'repeat', brkBinSub }))
    expect(lines('--')).toEqual([`a+b${minus}`, `${minus}c`])
    expect(lines('-+')).toEqual([`a+b${minus}`, '+c'])
    expect(lines('+-')).toEqual(['a+b+', `${minus}c`])
    // a plus is repeated as it is
    expect(texts(wrap(`a${minus}b+c`, 40, { brkBin: 'repeat', brkBinSub: '+-' }))).toEqual([
      `a${minus}b+`,
      '+c',
    ])
  })

  it('gives the wrapped lines the room left beside their indent', () => {
    expect(texts(wrap('a+b+c+d+e', 50, {}, 20))).toEqual(['a+b+c', '+d', '+e'])
  })

  it('marks and places the wrapped lines, keeping operator spacing at the line edges', () => {
    const indented = wrap('a+b+c+d+e', 50, { wrapIndentPx: 48 })!
    expect(indented.map((l) => l.getAttribute('class'))).toEqual([
      'doc-math-wrap-line',
      'doc-math-wrap-line doc-math-wrap-rest',
    ])
    expect(indented.map((l) => l.getAttribute('style'))).toEqual([null, 'margin-left:48px'])
    expect(indented[1].querySelector('mo')?.getAttribute('form')).toBe('infix')
    expect(indented[0].querySelector('mo')?.hasAttribute('form')).toBe(false)
    expect(indented.every((l) => l.getAttribute('display') === 'block')).toBe(true)

    const right = wrap('a+b+c+d+e', 50, { wrapRight: true })!
    expect(right[1].getAttribute('style')).toBe('margin-left:auto;margin-right:0')

    const centred = wrapMathLine(mathOf('a+b+c+d+e'), 50, 50, rules(), false, measure)!
    expect(centred[1].hasAttribute('style')).toBe(false)
  })

  it('leaves the source line untouched', () => {
    const math = mathOf('a+b+c+d+e')
    const before = math.outerHTML
    wrapMathLine(math, 50, 50, rules({ brkBin: 'repeat' }), true, measure)
    expect(math.outerHTML).toBe(before)
  })

  it('does not split what has no operator to break at, and overflows an operand too wide', () => {
    expect(wrap('abcdefgh', 50)).toBeNull()
    expect(wrap('f(a+b+c+d)', 50)).toBeNull()
    expect(texts(wrap('abcdefgh+i', 50))).toEqual(['abcdefgh', '+i'])
  })

  it('plans one line when everything fits', () => {
    expect(planWrapLines(5, [1, 3], (a, b) => (b - a + 1) * 10, 50, 50, 'before')).toEqual([
      { from: 0, to: 4 },
    ])
  })
})
