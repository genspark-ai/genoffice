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

// Lookahead keeps the "(" unconsumed so back-to-back calls (SORT(FILTER(…)
// both match; the lead guard rejects already-marked calls (the "." before
// the name) and sheet-qualified identifiers. Case-insensitive: Excel
// accepts lowercase input, and the marker must still be written.
const FUNCTION_CALL_PATTERN = /(^|[^A-Za-z0-9_."'!])([A-Za-z][A-Za-z0-9.]*)(?=\s*\()/g

/// Prefixes future-function calls with their storage markers, outside
/// string literals. Marked calls store the canonical uppercase name.
export function withFutureFunctionMarkers(formula: string): string {
  const segments = formula.split('"')
  for (let index = 0; index < segments.length; index += 2) {
    const segment = segments[index]
    if (segment === undefined) continue
    segments[index] = segment.replace(FUNCTION_CALL_PATTERN, (full, lead: string, name: string) => {
      const canonical = name.toUpperCase()
      if (XLWS_FUNCTIONS.has(canonical)) return `${lead}_xlfn._xlws.${canonical}`
      if (XLFN_FUNCTIONS.has(canonical)) return `${lead}_xlfn.${canonical}`
      return full
    })
  }
  return segments.join('"')
}
