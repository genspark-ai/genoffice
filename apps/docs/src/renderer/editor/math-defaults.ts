import { DEFAULT_MATH_PR, type MathLimits, type MathPr } from '@genoffice/docx-engine'

let limits: MathLimits = {}

/** m:intLim / m:naryLim of the open document, for every OMML conversion the editor runs */
export function mathLimits(): MathLimits {
  return limits
}

export function setMathLimits(mathPr: MathPr | undefined): void {
  limits = mathPr ? { intLim: mathPr.intLim, naryLim: mathPr.naryLim } : {}
}

/** an n-ary operator that leaves its limit placement to the document */
export function ommlFollowsLimitDefaults(omml: string): boolean {
  const count = (re: RegExp) => (omml.match(re) ?? []).length
  return count(/<m:nary[\s>]/g) > count(/<m:limLoc[\s/>]/g)
}

const LINE_MARGINS: Record<MathPr['defJc'], [string, string]> = {
  centerGroup: ['0', '0'],
  left: ['0', '0'],
  center: ['auto', 'auto'],
  right: ['auto', '0'],
}

/**
 * Custom properties placing the lines of a display equation for one m:jc value.
 * `para` names the set an equation carries for its own paragraph's alignment,
 * which mathDefaultsCss switches to while m:dispDef is off.
 */
export function mathJcVars(jc: MathPr['defJc'], para = false): string[] {
  const [left, right] = LINE_MARGINS[jc]
  const name = para ? '--doc-math-para' : '--doc-math-line'
  return [
    `--doc-math${para ? '-para' : ''}-jc:${jc === 'centerGroup' ? 'center' : jc}`,
    `${name}-l:${left}`,
    `${name}-r:${right}`,
  ]
}

/**
 * The document's display-equation defaults as CSS: equations without their own
 * m:jc read the alignment variables, every m:oMathPara reads the margins. With
 * m:dispDef off the margins drop and such an equation follows its paragraph.
 */
export function mathDefaultsCss(mathPr: MathPr = DEFAULT_MATH_PR): string {
  return [displayDefaultsCss(mathPr), ...mathLookCss(mathPr)].join('\n')
}

function displayDefaultsCss(mathPr: MathPr): string {
  if (!mathPr.dispDef) {
    return (
      `.doc-page [data-math-para='default'] { --doc-math-jc:var(--doc-math-para-jc, left); ` +
      `--doc-math-line-l:var(--doc-math-para-l, 0); --doc-math-line-r:var(--doc-math-para-r, 0) }`
    )
  }
  const decls = [
    ...mathJcVars(mathPr.defJc),
    `--doc-math-lm:${mathPr.lMarginTwips / 20}pt`,
    `--doc-math-rm:${mathPr.rMarginTwips / 20}pt`,
  ]
  return `.doc-page { ${decls.join(';')} }`
}

export const WORD_MATH_FONT = 'Cambria Math'

/** OpenType MATH faces a computer may carry besides Word's own, offered when installed */
export const MATH_FONT_CANDIDATES = [
  'STIX Two Math',
  'Latin Modern Math',
  'Asana Math',
  'Fira Math',
  'Libertinus Math',
  'Noto Sans Math',
  'TeX Gyre Pagella Math',
  'TeX Gyre Termes Math',
  'XITS Math',
]

/** MathML shrinks a nested fraction on its own; Word does so in a display equation only with m:smallFrac */
function mathLookCss(mathPr: MathPr): string[] {
  const rules: string[] = []
  const font = mathPr.mathFont?.replace(/["'\\;{}]/g, '').trim()
  if (font && font !== WORD_MATH_FONT) rules.push(`.doc-page { --doc-math-font:'${font}' }`)
  if (!mathPr.smallFrac)
    rules.push(
      `.doc-page :is(math[display='block'], math[displaystyle='true']) mfrac > * { math-depth:add(0) }`,
    )
  return rules
}
