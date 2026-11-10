/// Excel stores post-2007 function names in file formulas behind an
/// `_xlfn.` marker (worksheet-scope dynamic-array functions behind
/// `_xlfn._xlws.`) and shows the plain name in the UI. The sidecar strips
/// the markers on read; this restores them when a formula is serialized
/// back into sheet XML — without the marker Excel repairs the call to
/// #NAME?.

/// Functions Excel 2007 does not know, keyed by the marker they need.
const XLWS_FUNCTIONS = new Set(['FILTER', 'SORT'])

const XLFN_FUNCTIONS = new Set([
  // Excel 2010
  'AGGREGATE',
  'BETA.DIST',
  'BETA.INV',
  'BINOM.DIST',
  'BINOM.INV',
  'CEILING.PRECISE',
  'CHISQ.DIST',
  'CHISQ.DIST.RT',
  'CHISQ.INV',
  'CHISQ.INV.RT',
  'CHISQ.TEST',
  'CONFIDENCE.NORM',
  'CONFIDENCE.T',
  'COVARIANCE.P',
  'COVARIANCE.S',
  'ERF.PRECISE',
  'ERFC.PRECISE',
  'EXPON.DIST',
  'F.DIST',
  'F.DIST.RT',
  'F.INV',
  'F.INV.RT',
  'F.TEST',
  'FLOOR.PRECISE',
  'GAMMA.DIST',
  'GAMMA.INV',
  'GAMMALN.PRECISE',
  'HYPGEOM.DIST',
  'LOGNORM.DIST',
  'LOGNORM.INV',
  'MODE.MULT',
  'MODE.SNGL',
  'NEGBINOM.DIST',
  'NETWORKDAYS.INTL',
  'NORM.DIST',
  'NORM.INV',
  'NORM.S.DIST',
  'NORM.S.INV',
  'PERCENTILE.EXC',
  'PERCENTILE.INC',
  'PERCENTRANK.EXC',
  'PERCENTRANK.INC',
  'POISSON.DIST',
  'QUARTILE.EXC',
  'QUARTILE.INC',
  'RANK.AVG',
  'RANK.EQ',
  'STDEV.P',
  'STDEV.S',
  'T.DIST',
  'T.DIST.2T',
  'T.DIST.RT',
  'T.INV',
  'T.INV.2T',
  'T.TEST',
  'VAR.P',
  'VAR.S',
  'WEIBULL.DIST',
  'WORKDAY.INTL',
  'Z.TEST',
  // Excel 2013
  'ACOT',
  'ACOTH',
  'ARABIC',
  'BASE',
  'BINOM.DIST.RANGE',
  'BITAND',
  'BITLSHIFT',
  'BITOR',
  'BITRSHIFT',
  'BITXOR',
  'CEILING.MATH',
  'COMBINA',
  'COT',
  'COTH',
  'CSC',
  'CSCH',
  'DAYS',
  'DECIMAL',
  'ENCODEURL',
  'FILTERXML',
  'FLOOR.MATH',
  'FORMULATEXT',
  'GAMMA',
  'GAUSS',
  'IFNA',
  'IMCOSH',
  'IMCOT',
  'IMCSC',
  'IMCSCH',
  'IMSEC',
  'IMSECH',
  'IMSINH',
  'IMTAN',
  'ISFORMULA',
  'ISOWEEKNUM',
  'MUNIT',
  'NUMBERVALUE',
  'PDURATION',
  'PERMUTATIONA',
  'PHI',
  'RRI',
  'SEC',
  'SECH',
  'SHEET',
  'SHEETS',
  'SKEW.P',
  'UNICHAR',
  'UNICODE',
  'WEBSERVICE',
  'XOR',
  // Excel 2016
  'CONCAT',
  'FORECAST.ETS',
  'FORECAST.ETS.CONFINT',
  'FORECAST.ETS.SEASONALITY',
  'FORECAST.ETS.STAT',
  'FORECAST.LINEAR',
  'IFS',
  'MAXIFS',
  'MINIFS',
  'SWITCH',
  'TEXTJOIN',
  // Microsoft 365
  'BYCOL',
  'BYROW',
  'CHOOSECOLS',
  'CHOOSEROWS',
  'DROP',
  'EXPAND',
  'HSTACK',
  'IMAGE',
  'ISOMITTED',
  'LAMBDA',
  'LET',
  'MAKEARRAY',
  'MAP',
  'RANDARRAY',
  'REDUCE',
  'SCAN',
  'SEQUENCE',
  'SORTBY',
  'STOCKHISTORY',
  'TAKE',
  'TEXTAFTER',
  'TEXTBEFORE',
  'TEXTSPLIT',
  'TOCOL',
  'TOROW',
  'UNIQUE',
  'VSTACK',
  'WRAPCOLS',
  'WRAPROWS',
  'XLOOKUP',
  'XMATCH',
  // Storage forms of the @ implicit-intersection and # spill operators.
  'SINGLE',
  'ANCHORARRAY',
])

/// Functions whose result spills into neighbouring cells. Excel 365 only
/// spills a stored formula that is marked as a dynamic array (`t="array"` on
/// the <f> plus `cm="1"` on the cell); an unmarked call is read as `@FILTER`
/// and returns one value.
const SPILL_FUNCTIONS = new Set([
  'BYCOL',
  'BYROW',
  'CHOOSECOLS',
  'CHOOSEROWS',
  'DROP',
  'EXPAND',
  'FILTER',
  'HSTACK',
  'MAKEARRAY',
  'MAP',
  'RANDARRAY',
  'SCAN',
  'SEQUENCE',
  'SORT',
  'SORTBY',
  'TAKE',
  'TEXTSPLIT',
  'TOCOL',
  'TOROW',
  'UNIQUE',
  'VSTACK',
  'WRAPCOLS',
  'WRAPROWS',
])

interface FormulaCall {
  start: number
  end: number
  name: string
  marked: boolean
}

function isNameCharacter(character: string): boolean {
  return (
    character === '_' ||
    character === '.' ||
    character === '\\' ||
    character === '$' ||
    /[\p{L}\p{N}]/u.test(character)
  )
}

function quotedEnd(formula: string, start: number): number {
  const quote = formula[start]
  let index = start + 1
  while (index < formula.length) {
    if (formula[index] === quote) {
      if (formula[index + 1] === quote) {
        index += 2
        continue
      }
      return index + 1
    }
    index += 1
  }
  return formula.length
}

function functionCalls(formula: string): FormulaCall[] {
  const calls: FormulaCall[] = []
  let cursor = 0
  while (cursor < formula.length) {
    const character = formula[cursor]!
    if (character === '"' || character === "'") {
      cursor = quotedEnd(formula, cursor)
      continue
    }
    const codePoint = formula.codePointAt(cursor)!
    const symbol = String.fromCodePoint(codePoint)
    if (!isNameCharacter(symbol)) {
      cursor += symbol.length
      continue
    }
    const start = cursor
    let end = cursor
    while (end < formula.length) {
      const current = String.fromCodePoint(formula.codePointAt(end)!)
      if (!isNameCharacter(current)) break
      end += current.length
    }
    let after = end
    while (after < formula.length && /\s/.test(formula[after]!)) after += 1
    if (formula[after] === '(') {
      const token = formula.slice(start, end)
      const marked = /^_xlfn\.(?:_xlws\.)?/i.test(token) || /^_xlws\./i.test(token)
      const name = token.replace(/^_xlfn\.(?:_xlws\.)?/i, '').replace(/^_xlws\./i, '')
      calls.push({ start, end, name, marked })
    }
    cursor = end
  }
  return calls
}

export function spillsDynamicArray(formula: string): boolean {
  return functionCalls(formula).some((call) => SPILL_FUNCTIONS.has(call.name.toUpperCase()))
}

/// Prefixes future-function calls with their storage markers, outside
/// string literals. Marked calls store the canonical uppercase name.
export function withFutureFunctionMarkers(formula: string): string {
  let out = ''
  let cursor = 0
  for (const call of functionCalls(formula)) {
    out += formula.slice(cursor, call.start)
    const token = formula.slice(call.start, call.end)
    const canonical = call.name.toUpperCase()
    if (call.marked) {
      out += token
    } else if (XLWS_FUNCTIONS.has(canonical)) {
      out += `_xlfn._xlws.${canonical}`
    } else if (XLFN_FUNCTIONS.has(canonical)) {
      out += `_xlfn.${canonical}`
    } else {
      out += token
    }
    cursor = call.end
  }
  return out + formula.slice(cursor)
}
