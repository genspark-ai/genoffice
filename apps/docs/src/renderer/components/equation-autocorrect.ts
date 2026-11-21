import { latexToOmml } from '@genoffice/docx-engine'
import { autocorrectEnabled, autocorrectLists } from '../autocorrect-pref'

/** function names LaTeX has a command for; any other recognized name goes through \operatorname */
const LATEX_FUNCTIONS = new Set(
  'sin cos tan cot sec csc sinh cosh tanh coth arcsin arccos arctan ln log exp max min sup inf arg det gcd deg dim ker lim'.split(
    ' ',
  ),
)

/** inside \text{…} and the like the words are prose, not math */
function inTextGroup(latex: string): boolean {
  const open = /\\(?:text|mathrm|operatorname|textbf|textit)\{[^{}]*$/.test(latex)
  return open
}

/** a command the converter reads, bare or with the arguments it takes */
function latexKnows(command: string): boolean {
  return ['', '{x}', '{x}{y}'].some((args) => {
    try {
      return !latexToOmml(command + args).includes(command.slice(1))
    } catch {
      return false
    }
  })
}

/**
 * What the equation box holds after one more character: a recognized function
 * typed bare becomes its upright command at the space or bracket after it, and
 * a Math AutoCorrect entry is replaced at the space after it. Names LaTeX
 * already reads (\alpha) are left as typed.
 */
export function equationAutoCorrect(prev: string, next: string): string {
  if (next.length !== prev.length + 1 || !next.startsWith(prev)) return next
  if (!autocorrectEnabled('enabled') || inTextGroup(prev)) return next
  const typed = next.slice(-1)
  if (typed !== ' ' && typed !== '(') return next
  const lists = autocorrectLists()
  const fn = /(?:^|[^\\A-Za-z])([A-Za-z]{2,})$/.exec(prev)?.[1]
  if (fn && lists.functions.includes(fn)) {
    const command = LATEX_FUNCTIONS.has(fn) ? `\\${fn}` : `\\operatorname{${fn}}`
    return prev.slice(0, -fn.length) + command + typed
  }
  if (typed !== ' ' || !autocorrectEnabled('mathInside')) return next
  const token = /(?:^|\s)(\S+)$/.exec(prev)?.[1]
  const to = token && lists.math.find(([from]) => from === token)?.[1]
  if (!token || to === undefined || (token.startsWith('\\') && latexKnows(token))) return next
  return `${prev.slice(0, -token.length)}${to} `
}
