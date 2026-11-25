/** Persisted AutoCorrect settings (Tools → AutoCorrect Options…), read on every keystroke. */
export const AUTOCORRECT_KEY = 'aidocs.autocorrect'
const LISTS_KEY = 'aidocs.autocorrect.lists'

export const AUTOCORRECT_RULES = [
  'enabled',
  'smartButton',
  'twoInitialCaps',
  'capitalize',
  'capitalizeDays',
  'capitalizeCells',
  'replaceText',
  'spellSuggest',
  'mathOutside',
  'mathInside',
  'autoBullets',
  'autoNumbers',
  'borderLines',
  'tables',
  'headingStyles',
  'dateStyle',
  'indentSpaces',
  'asianSpaces',
  'boldItalic',
  'fractions',
  'hyperlinks',
  'parens',
  'ordinals',
  'smartQuotes',
  'dashes',
  'defineStyles',
  'tabIndents',
  'listStart',
  'closingPhrase',
  'autoComplete',
] as const

export type AutocorrectRule = (typeof AUTOCORRECT_RULES)[number]

// Word's defaults, except capitalizing sentences and table cells: they misfire
// on CJK-mixed and technical text often enough that we leave them opt-in
export const AUTOCORRECT_DEFAULTS: Record<AutocorrectRule, boolean> = {
  enabled: true,
  smartButton: true,
  twoInitialCaps: true,
  capitalize: false,
  capitalizeDays: true,
  capitalizeCells: false,
  replaceText: true,
  // on in Word; the dictionary here offers a single suggestion less carefully, so it is opt-in
  spellSuggest: false,
  mathOutside: false,
  mathInside: true,
  autoBullets: true,
  autoNumbers: true,
  borderLines: true,
  tables: true,
  headingStyles: false,
  dateStyle: false,
  indentSpaces: true,
  asianSpaces: false,
  boldItalic: false,
  fractions: true,
  hyperlinks: true,
  parens: true,
  ordinals: true,
  smartQuotes: true,
  dashes: true,
  defineStyles: false,
  tabIndents: true,
  listStart: true,
  closingPhrase: true,
  autoComplete: true,
}

// switches saved before the dialog had Word's tabs
const LEGACY_RULES: Record<string, AutocorrectRule[]> = {
  autoLists: ['autoBullets', 'autoNumbers'],
  symbols: ['replaceText', 'fractions'],
}

function readJson(key: string): Record<string, unknown> {
  try {
    const raw = globalThis.localStorage?.getItem(key)
    if (!raw) return {}
    const parsed: unknown = JSON.parse(raw)
    return parsed && typeof parsed === 'object' ? (parsed as Record<string, unknown>) : {}
  } catch {
    return {}
  }
}

function writeJson(key: string, value: unknown): void {
  try {
    globalThis.localStorage?.setItem(key, JSON.stringify(value))
  } catch {
    /* storage unavailable: the setting just doesn't persist */
  }
}

export function autocorrectPrefs(): Record<AutocorrectRule, boolean> {
  const stored = readJson(AUTOCORRECT_KEY)
  const out = { ...AUTOCORRECT_DEFAULTS }
  for (const [legacy, rules] of Object.entries(LEGACY_RULES)) {
    if (typeof stored[legacy] === 'boolean') for (const r of rules) out[r] = stored[legacy]
  }
  for (const rule of AUTOCORRECT_RULES) {
    const v = stored[rule]
    if (typeof v === 'boolean') out[rule] = v
  }
  return out
}

/** a rule runs only while the dialog's master switch is on */
export function autocorrectEnabled(rule: AutocorrectRule): boolean {
  const prefs = autocorrectPrefs()
  return prefs.enabled && prefs[rule]
}

export function setAutocorrectPref(rule: AutocorrectRule, on: boolean): void {
  writeJson(AUTOCORRECT_KEY, { ...readJson(AUTOCORRECT_KEY), [rule]: on })
}

export function resetAutocorrectPrefs(): void {
  try {
    globalThis.localStorage?.removeItem(AUTOCORRECT_KEY)
    globalThis.localStorage?.removeItem(LISTS_KEY)
  } catch {
    /* ignore */
  }
}

/** [typed, replacement, the replacement's formatted HTML when it was kept as formatted text] */
export type ReplaceEntry = [from: string, to: string, html?: string]

export const DEFAULT_REPLACE_ENTRIES: ReplaceEntry[] = [
  ['(c)', '\u00a9'],
  ['(r)', '\u00ae'],
  ['(tm)', '\u2122'],
  ['...', '\u2026'],
  ['-->', '\u2192'],
  ['<--', '\u2190'],
  ['==>', '\u21d2'],
  ['<==', '\u21d0'],
  ['<=>', '\u21d4'],
]

export const DEFAULT_MATH_ENTRIES: ReplaceEntry[] = [
  ['...', '\u2026'],
  ['->', '\u2192'],
  ['<=', '\u2264'],
  ['>=', '\u2265'],
  ['+-', '\u00b1'],
  ['-+', '\u2213'],
  ['~=', '\u2245'],
  ['<<', '\u226a'],
  ['>>', '\u226b'],
  ['!!', '\u203c'],
  ['\\aleph', '\u2135'],
  ['\\alpha', '\u03b1'],
  ['\\angle', '\u2220'],
  ['\\angstrom', '\u00c5'],
  ['\\approx', '\u2248'],
  ['\\ast', '\u2217'],
  ['\\because', '\u2235'],
  ['\\beta', '\u03b2'],
  ['\\beth', '\u2136'],
  ['\\bot', '\u22a5'],
  ['\\box', '\u25a1'],
  ['\\bullet', '\u2219'],
  ['\\cap', '\u2229'],
  ['\\cbrt', '\u221b'],
  ['\\cdot', '\u22c5'],
  ['\\cdots', '\u22ef'],
  ['\\chi', '\u03c7'],
  ['\\circ', '\u2218'],
  ['\\cong', '\u2245'],
  ['\\coprod', '\u2210'],
  ['\\cup', '\u222a'],
  ['\\dagger', '\u2020'],
  ['\\dashv', '\u22a3'],
  ['\\ddagger', '\u2021'],
  ['\\ddots', '\u22f1'],
  ['\\degree', '\u00b0'],
  ['\\delta', '\u03b4'],
  ['\\Delta', '\u0394'],
  ['\\diamond', '\u22c4'],
  ['\\div', '\u00f7'],
  ['\\dots', '\u2026'],
  ['\\doubleC', '\u2102'],
  ['\\doubleN', '\u2115'],
  ['\\doubleQ', '\u211a'],
  ['\\doubleR', '\u211d'],
  ['\\doubleZ', '\u2124'],
  ['\\downarrow', '\u2193'],
  ['\\Downarrow', '\u21d3'],
  ['\\ell', '\u2113'],
  ['\\emptyset', '\u2205'],
  ['\\epsilon', '\u03f5'],
  ['\\equiv', '\u2261'],
  ['\\eta', '\u03b7'],
  ['\\exists', '\u2203'],
  ['\\forall', '\u2200'],
  ['\\gamma', '\u03b3'],
  ['\\Gamma', '\u0393'],
  ['\\ge', '\u2265'],
  ['\\geq', '\u2265'],
  ['\\gets', '\u2190'],
  ['\\gg', '\u226b'],
  ['\\hbar', '\u210f'],
  ['\\hookleftarrow', '\u21a9'],
  ['\\hookrightarrow', '\u21aa'],
  ['\\iiint', '\u222d'],
  ['\\iint', '\u222c'],
  ['\\Im', '\u2111'],
  ['\\imath', '\u0131'],
  ['\\in', '\u2208'],
  ['\\infty', '\u221e'],
  ['\\int', '\u222b'],
  ['\\iota', '\u03b9'],
  ['\\kappa', '\u03ba'],
  ['\\lambda', '\u03bb'],
  ['\\Lambda', '\u039b'],
  ['\\langle', '\u27e8'],
  ['\\lceil', '\u2308'],
  ['\\ldots', '\u2026'],
  ['\\le', '\u2264'],
  ['\\leftarrow', '\u2190'],
  ['\\Leftarrow', '\u21d0'],
  ['\\leftrightarrow', '\u2194'],
  ['\\Leftrightarrow', '\u21d4'],
  ['\\leq', '\u2264'],
  ['\\lfloor', '\u230a'],
  ['\\ll', '\u226a'],
  ['\\mapsto', '\u21a6'],
  ['\\models', '\u22a8'],
  ['\\mp', '\u2213'],
  ['\\mu', '\u03bc'],
  ['\\nabla', '\u2207'],
  ['\\ne', '\u2260'],
  ['\\nearrow', '\u2197'],
  ['\\neg', '\u00ac'],
  ['\\neq', '\u2260'],
  ['\\ni', '\u220b'],
  ['\\notin', '\u2209'],
  ['\\nu', '\u03bd'],
  ['\\nwarrow', '\u2196'],
  ['\\odot', '\u2299'],
  ['\\oint', '\u222e'],
  ['\\omega', '\u03c9'],
  ['\\Omega', '\u03a9'],
  ['\\ominus', '\u2296'],
  ['\\oplus', '\u2295'],
  ['\\otimes', '\u2297'],
  ['\\parallel', '\u2225'],
  ['\\partial', '\u2202'],
  ['\\perp', '\u22a5'],
  ['\\phi', '\u03d5'],
  ['\\Phi', '\u03a6'],
  ['\\pi', '\u03c0'],
  ['\\Pi', '\u03a0'],
  ['\\pm', '\u00b1'],
  ['\\prec', '\u227a'],
  ['\\prod', '\u220f'],
  ['\\propto', '\u221d'],
  ['\\psi', '\u03c8'],
  ['\\Psi', '\u03a8'],
  ['\\qdrt', '\u221c'],
  ['\\rangle', '\u27e9'],
  ['\\rceil', '\u2309'],
  ['\\Re', '\u211c'],
  ['\\rfloor', '\u230b'],
  ['\\rho', '\u03c1'],
  ['\\rightarrow', '\u2192'],
  ['\\Rightarrow', '\u21d2'],
  ['\\searrow', '\u2198'],
  ['\\setminus', '\u2216'],
  ['\\sigma', '\u03c3'],
  ['\\Sigma', '\u03a3'],
  ['\\sim', '\u223c'],
  ['\\simeq', '\u2243'],
  ['\\sqrt', '\u221a'],
  ['\\star', '\u22c6'],
  ['\\subset', '\u2282'],
  ['\\subseteq', '\u2286'],
  ['\\succ', '\u227b'],
  ['\\sum', '\u2211'],
  ['\\superset', '\u2283'],
  ['\\superseteq', '\u2287'],
  ['\\swarrow', '\u2199'],
  ['\\tau', '\u03c4'],
  ['\\therefore', '\u2234'],
  ['\\theta', '\u03b8'],
  ['\\Theta', '\u0398'],
  ['\\times', '\u00d7'],
  ['\\to', '\u2192'],
  ['\\top', '\u22a4'],
  ['\\uparrow', '\u2191'],
  ['\\Uparrow', '\u21d1'],
  ['\\updownarrow', '\u2195'],
  ['\\upsilon', '\u03c5'],
  ['\\Upsilon', '\u03a5'],
  ['\\varepsilon', '\u03b5'],
  ['\\varphi', '\u03c6'],
  ['\\varpi', '\u03d6'],
  ['\\varrho', '\u03f1'],
  ['\\varsigma', '\u03c2'],
  ['\\vartheta', '\u03d1'],
  ['\\vdash', '\u22a2'],
  ['\\vdots', '\u22ee'],
  ['\\vee', '\u2228'],
  ['\\wedge', '\u2227'],
  ['\\wp', '\u2118'],
  ['\\xi', '\u03be'],
  ['\\Xi', '\u039e'],
  ['\\zeta', '\u03b6'],
]

/** abbreviations whose period does not end a sentence */
export const DEFAULT_FIRST_LETTER_EXCEPTIONS = [
  'a.m.',
  'approx.',
  'apr.',
  'aug.',
  'cf.',
  'co.',
  'dec.',
  'dr.',
  'e.g.',
  'etc.',
  'feb.',
  'fig.',
  'i.e.',
  'inc.',
  'jan.',
  'jr.',
  'jul.',
  'jun.',
  'ltd.',
  'mar.',
  'mr.',
  'mrs.',
  'ms.',
  'no.',
  'nov.',
  'oct.',
  'p.m.',
  'prof.',
  'sep.',
  'sept.',
  'sr.',
  'st.',
  'u.k.',
  'u.s.',
  'vs.',
]

/** words that really start with two capitals */
export const DEFAULT_INITIAL_CAPS_EXCEPTIONS = ['GHz', 'IDs', 'JQuery', 'MHz', 'OAuth', 'VMware']

/** names an equation sets upright when they are typed without a backslash */
export const DEFAULT_MATH_FUNCTIONS = [
  'arccos',
  'arccot',
  'arccsc',
  'arcsec',
  'arcsin',
  'arctan',
  'arg',
  'cos',
  'cosh',
  'cot',
  'coth',
  'csc',
  'csch',
  'deg',
  'det',
  'dim',
  'exp',
  'gcd',
  'hom',
  'inf',
  'ker',
  'lg',
  'lim',
  'ln',
  'log',
  'max',
  'min',
  'Pr',
  'sec',
  'sech',
  'sin',
  'sinh',
  'sup',
  'tan',
  'tanh',
]

export interface AutocorrectLists {
  replace: ReplaceEntry[]
  math: ReplaceEntry[]
  firstLetter: string[]
  initialCaps: string[]
  functions: string[]
}

export type AutocorrectListName = keyof AutocorrectLists

const LIST_DEFAULTS: AutocorrectLists = {
  replace: DEFAULT_REPLACE_ENTRIES,
  math: DEFAULT_MATH_ENTRIES,
  firstLetter: DEFAULT_FIRST_LETTER_EXCEPTIONS,
  initialCaps: DEFAULT_INITIAL_CAPS_EXCEPTIONS,
  functions: DEFAULT_MATH_FUNCTIONS,
}

const isEntry = (v: unknown): v is ReplaceEntry =>
  Array.isArray(v) &&
  (v.length === 2 || v.length === 3) &&
  v.every((part) => typeof part === 'string')

let listsCache: { raw: string | null; lists: AutocorrectLists } | null = null

/** the lists as edited in the dialog; one never edited is the built-in one */
export function autocorrectLists(): AutocorrectLists {
  let raw: string | null = null
  try {
    raw = globalThis.localStorage?.getItem(LISTS_KEY) ?? null
  } catch {
    /* built-in lists */
  }
  if (listsCache && listsCache.raw === raw) return listsCache.lists
  const stored = readJson(LISTS_KEY)
  const entries = (name: 'replace' | 'math') => {
    const v = stored[name]
    return Array.isArray(v) ? v.filter(isEntry) : LIST_DEFAULTS[name]
  }
  const words = (name: 'firstLetter' | 'initialCaps' | 'functions') => {
    const v = stored[name]
    return Array.isArray(v)
      ? v.filter((w): w is string => typeof w === 'string')
      : LIST_DEFAULTS[name]
  }
  const lists = {
    replace: entries('replace'),
    math: entries('math'),
    firstLetter: words('firstLetter'),
    initialCaps: words('initialCaps'),
    functions: words('functions'),
  }
  listsCache = { raw, lists }
  return lists
}

export function setAutocorrectList<K extends AutocorrectListName>(
  name: K,
  value: AutocorrectLists[K],
): void {
  writeJson(LISTS_KEY, { ...readJson(LISTS_KEY), [name]: value })
}

/** add `from`, or change what an entry of that spelling is replaced with; sorted like Word's table */
export function upsertEntry(
  list: ReplaceEntry[],
  from: string,
  to: string,
  html?: string,
): ReplaceEntry[] {
  const rest = list.filter(([f]) => f !== from)
  const entry: ReplaceEntry = html ? [from, to, html] : [from, to]
  return [...rest, entry].sort((a, b) => a[0].localeCompare(b[0]))
}
