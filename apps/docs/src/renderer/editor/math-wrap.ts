import { DEFAULT_MATH_BREAKS, type MathPr } from '@genoffice/docx-engine'

/**
 * Chromium never breaks a formula, so a display equation wider than its column
 * ran off the page. This pass splits such an equation at its top-level binary
 * and relational operators into the same one-`<math>`-per-line shape a manual
 * break already has. Display only: the node's OMML and MathML stay whole.
 */

export type MathWrapRules = Required<Pick<MathPr, 'brkBin' | 'brkBinSub' | 'wrapRight'>> & {
  wrapIndentPx: number
}

export function mathWrapRulesOf(mathPr: MathPr | undefined): MathWrapRules {
  return {
    brkBin: mathPr?.brkBin ?? DEFAULT_MATH_BREAKS.brkBin,
    brkBinSub: mathPr?.brkBinSub ?? DEFAULT_MATH_BREAKS.brkBinSub,
    wrapRight: mathPr?.wrapRight === true,
    wrapIndentPx: (mathPr?.wrapIndentTwips ?? DEFAULT_MATH_BREAKS.wrapIndentTwips) / 15,
  }
}

const MINUS = '−'
const BREAK_OPERATORS = new Set('+−±∓×÷·⋅∙*=<>≤≥≠≈≡∼∝' + '→←↔⇒⇐⇔∈∉⊂⊃∪∩∧∨')
const OPENERS = new Set('([{')
const CLOSERS = new Set(')]}')

/** operators an operand can end with */
const OPERAND_ENDS = new Set(')]}!%\u2032\u2033\u00B0\u221E')

const isMo = (el: Element) => el.localName === 'mo'
const endsOperand = (el: Element) => !isMo(el) || OPERAND_ENDS.has(el.textContent ?? '')

/**
 * Indices of the row's children a line may break at: operators between two
 * operands, outside any bracket typed as plain characters. A sign after
 * another operator or an opening bracket is unary and stays with its operand.
 */
export function breakCandidates(children: readonly Element[]): number[] {
  const out: number[] = []
  let depth = 0
  for (let i = 0; i < children.length; i++) {
    const el = children[i]
    if (!isMo(el)) continue
    const ch = el.textContent ?? ''
    if (OPENERS.has(ch)) depth++
    else if (CLOSERS.has(ch)) depth = Math.max(0, depth - 1)
    else if (
      depth === 0 &&
      BREAK_OPERATORS.has(ch) &&
      i > 0 &&
      i < children.length - 1 &&
      endsOperand(children[i - 1])
    )
      out.push(i)
  }
  return out
}

export interface WrapLine {
  /** child indices [from, to] of the row */
  from: number
  to: number
}

/**
 * Greedy lines over the row's children. `span(from, to)` is the width those
 * children take side by side; a line that cannot fit even one operand keeps it
 * and overflows as the whole equation did.
 */
export function planWrapLines(
  count: number,
  candidates: readonly number[],
  span: (from: number, to: number) => number,
  availFirst: number,
  availRest: number,
  brkBin: MathWrapRules['brkBin'],
): WrapLine[] {
  const lines: WrapLine[] = []
  let start = 0
  while (start < count) {
    const avail = lines.length === 0 ? availFirst : availRest
    const open = candidates.filter((k) => k > start)
    if (open.length === 0 || span(start, count - 1) <= avail) {
      lines.push({ from: start, to: count - 1 })
      break
    }
    const lastOf = (k: number) => (brkBin === 'before' ? k - 1 : k)
    let pick = open[0]
    for (const k of open) {
      if (span(start, lastOf(k)) > avail) break
      pick = k
    }
    lines.push({ from: start, to: lastOf(pick) })
    start = brkBin === 'after' ? pick + 1 : pick
  }
  return lines
}

const MATH_NS = 'http://www.w3.org/1998/Math/MathML'
export const WRAP_LINE_CLASS = 'doc-math-wrap-line'
export const WRAP_REST_CLASS = 'doc-math-wrap-rest'

/** the sign a repeated subtraction shows on the line it ends and on the line it starts */
function repeatedSign(el: Element, side: 0 | 1, brkBinSub: MathWrapRules['brkBinSub']): string {
  const ch = el.textContent ?? ''
  if (ch !== MINUS && ch !== '-') return ch
  return brkBinSub[side] === '+' ? '+' : MINUS
}

/**
 * One `<math display="block">` per planned line, built from copies of the
 * row's children. An operator moved to a line's edge keeps its infix spacing.
 */
export function buildWrapLines(
  math: Element,
  children: readonly Element[],
  lines: readonly WrapLine[],
  rules: MathWrapRules,
  indent: boolean,
): Element[] {
  const doc = math.ownerDocument
  return lines.map((line, li) => {
    const out = math.cloneNode(false) as Element
    out.classList.add(WRAP_LINE_CLASS)
    if (li > 0) {
      out.classList.add(WRAP_REST_CLASS)
      const style = rules.wrapRight
        ? 'margin-left:auto;margin-right:0'
        : indent
          ? `margin-left:${Math.round(rules.wrapIndentPx * 100) / 100}px`
          : ''
      if (style) out.setAttribute('style', style)
    }
    const row = doc.createElementNS(MATH_NS, 'mrow')
    for (let i = line.from; i <= line.to; i++) {
      const copy = children[i].cloneNode(true) as Element
      const leads = i === line.from && li > 0
      const ends = i === line.to && li < lines.length - 1
      if (isMo(copy) && (leads || ends)) {
        copy.setAttribute('form', 'infix')
        if (rules.brkBin === 'repeat')
          copy.textContent = repeatedSign(copy, leads ? 1 : 0, rules.brkBinSub)
      }
      row.appendChild(copy)
    }
    out.appendChild(row)
    return out
  })
}

/** the single row ommlToMathML wraps an equation line in */
function rowOf(math: Element): Element | null {
  const row = math.firstElementChild
  return row && math.childElementCount === 1 && row.localName === 'mrow' ? row : null
}

interface Measured {
  left: number
  right: number
}

/**
 * The lines replacing one over-wide equation line, or null when it has no
 * operator to break at. `measure` reads layout px.
 */
export function wrapMathLine(
  math: Element,
  availFirst: number,
  availRest: number,
  rules: MathWrapRules,
  indent: boolean,
  measure: (el: Element) => Measured,
): Element[] | null {
  const row = rowOf(math)
  if (!row) return null
  const children = Array.from(row.children)
  const candidates = breakCandidates(children)
  if (candidates.length === 0) return null
  const boxes = children.map(measure)
  const span = (from: number, to: number) => boxes[to].right - boxes[from].left
  const lines = planWrapLines(
    children.length,
    candidates,
    span,
    availFirst,
    availRest,
    rules.brkBin,
  )
  if (lines.length < 2) return null
  return buildWrapLines(math, children, lines, rules, indent)
}

export const MATH_WRAP_CHANGED_EVENT = 'doc-math-wrap-changed'

let rules = mathWrapRulesOf(undefined)
const blocks = new Set<HTMLElement>()
/** the unwrapped markup of a host whose lines this pass replaced */
const pristine = new WeakMap<Element, string>()
const lastSig = new WeakMap<Element, string>()
let observer: ResizeObserver | null = null

/** a line built from measured operand positions can still come out a hair wide */
const REFIT_ROUNDS = 3
const FIT_TOLERANCE = 0.5

function inBodyFlow(block: HTMLElement): boolean {
  return (
    block.hasAttribute('data-math-para') &&
    block.parentElement?.matches('.editor-scroll .ProseMirror.doc-page') === true
  )
}

/**
 * Wrap (or unwrap) one display equation for its current column width. Returns
 * whether its lines changed.
 */
export function wrapDisplayEquation(block: HTMLElement, force = false): boolean {
  const host = block.querySelector('.doc-formula-math')
  if (!host || !inBodyFlow(block)) return false
  const blockRect = block.getBoundingClientRect()
  // hidden while its tokens are edited
  if (!block.offsetWidth || !blockRect.width || !host.getBoundingClientRect().width) return false
  const scale = blockRect.width / block.offsetWidth
  const cs = getComputedStyle(block)
  const avail =
    block.clientWidth - (parseFloat(cs.paddingLeft) || 0) - (parseFloat(cs.paddingRight) || 0)
  const first = host.querySelector('math')
  if (!first || avail <= 0) return false
  const indent =
    !rules.wrapRight &&
    getComputedStyle(first).getPropertyValue('--doc-math-line-l').trim() !== 'auto'
  const sig = [
    avail,
    rules.brkBin,
    rules.brkBinSub,
    rules.wrapRight,
    rules.wrapIndentPx,
    indent,
  ].join('|')
  if (!force && lastSig.get(host) === sig) return false
  lastSig.set(host, sig)

  const before = host.innerHTML
  const source = pristine.get(host) ?? before
  const measure = (el: Element): Measured => {
    const r = el.getBoundingClientRect()
    return { left: r.left / scale, right: r.right / scale }
  }
  const widthOf = (el: Element) => el.getBoundingClientRect().width / scale
  const availRest = indent ? Math.max(0, avail - rules.wrapIndentPx) : avail
  let wrapped = false
  let slack = 0
  for (let round = 0; round < REFIT_ROUNDS; round++) {
    host.innerHTML = source
    ;(host as HTMLElement).style.minWidth = ''
    wrapped = false
    let over = 0
    for (const math of Array.from(host.querySelectorAll(':scope > math'))) {
      if (widthOf(math) <= avail + FIT_TOLERANCE) continue
      const lines = wrapMathLine(math, avail - slack, availRest - slack, rules, indent, measure)
      if (!lines) continue
      math.replaceWith(...lines)
      wrapped = true
      lines.forEach((line, i) => {
        const breakable = breakCandidates(Array.from(rowOf(line)?.children ?? [])).length > 0
        const room = i === 0 ? avail : availRest
        if (breakable) over = Math.max(over, widthOf(line) - room)
      })
    }
    if (over <= FIT_TOLERANCE) break
    slack += over
  }
  if (wrapped) {
    pristine.set(host, source)
    // right-aligned lines end at the column edge, not at the widest line
    if (rules.wrapRight) (host as HTMLElement).style.minWidth = `${avail}px`
  } else {
    pristine.delete(host)
  }
  return host.innerHTML !== before
}

function run(block: HTMLElement, force: boolean): void {
  if (wrapDisplayEquation(block, force))
    document.dispatchEvent(new CustomEvent(MATH_WRAP_CHANGED_EVENT, { detail: block }))
}

function rewrapAll(): void {
  for (const block of blocks) if (block.isConnected) run(block, true)
}

/** the open document's break rules; every wrapped equation is laid out again */
export function setMathWrapRules(mathPr: MathPr | undefined): void {
  rules = mathWrapRulesOf(mathPr)
  rewrapAll()
}

/** keep one display equation wrapped to its column as the column changes */
export function observeMathWrap(block: HTMLElement): (() => void) | undefined {
  if (!block.hasAttribute('data-math-para') || typeof ResizeObserver === 'undefined') return
  if (!observer) {
    observer = new ResizeObserver((entries) => {
      for (const entry of entries) run(entry.target as HTMLElement, false)
    })
    document.fonts?.addEventListener('loadingdone', rewrapAll)
  }
  blocks.add(block)
  observer.observe(block)
  return () => {
    blocks.delete(block)
    observer?.unobserve(block)
  }
}
