const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

export type MathJc = 'centerGroup' | 'center' | 'left' | 'right'
export type MathLimLoc = 'subSup' | 'undOvr'

/** where the limits of an n-ary operator without its own m:limLoc go */
export interface MathLimits {
  intLim?: MathLimLoc
  naryLim?: MathLimLoc
}

/** settings.xml m:mathPr: the document's equation defaults */
export interface MathPr {
  /** m:dispDef: display equations use defJc and the margins; off = they follow their paragraph */
  dispDef: boolean
  defJc: MathJc
  lMarginTwips: number
  rMarginTwips: number
  intLim: MathLimLoc
  naryLim: MathLimLoc
  mathFont?: string
  brkBin?: 'before' | 'after' | 'repeat'
  brkBinSub?: '--' | '-+' | '+-'
  smallFrac?: boolean
  wrapIndentTwips?: number
  wrapRight?: boolean
}

/** Word's values for a document whose settings carry no m:mathPr */
export const DEFAULT_MATH_PR: MathPr = {
  dispDef: true,
  defJc: 'centerGroup',
  lMarginTwips: 0,
  rMarginTwips: 0,
  intLim: 'subSup',
  naryLim: 'undOvr',
}

/** the fields saveMathPr writes; the rest of m:mathPr is carried untouched */
export type MathPrEdit = Pick<
  MathPr,
  | 'dispDef'
  | 'defJc'
  | 'lMarginTwips'
  | 'rMarginTwips'
  | 'intLim'
  | 'naryLim'
  | 'mathFont'
  | 'smallFrac'
  | 'brkBin'
  | 'brkBinSub'
  | 'wrapIndentTwips'
  | 'wrapRight'
>

/** Word's values for the line-break children a part leaves out */
export const DEFAULT_MATH_BREAKS = {
  brkBin: 'before',
  brkBinSub: '--',
  wrapIndentTwips: 1440,
} as const

/** CT_MathPr child sequence (ECMA-376 22.1.2.62); wrapIndent / wrapRight are one choice */
const MATH_PR_CHILD_ORDER = [
  'mathFont',
  'brkBin',
  'brkBinSub',
  'smallFrac',
  'dispDef',
  'lMargin',
  'rMargin',
  'defJc',
  'preSp',
  'postSp',
  'interSp',
  'intraSp',
  'wrapIndent',
  'wrapRight',
  'intLim',
  'naryLim',
]

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/** the prefix the part binds to the math namespace; null when it declares none */
function mathPrefixOf(xml: string): string | null {
  return (
    new RegExp(`\\sxmlns:([A-Za-z0-9._-]+)\\s*=\\s*(["'])${escapeRe(MATH_NS)}\\2`).exec(xml)?.[1] ??
    null
  )
}

interface MathPrSpan {
  start: number
  end: number
  /** between the open and close tags; '' for a self-closed element */
  inner: string
  open: string
}

function mathPrSpan(xml: string, prefix: string): MathPrSpan | null {
  const p = escapeRe(prefix)
  const m = new RegExp(`<${p}:mathPr(?=[\\s/>])([^>]*?)(/>|>([\\s\\S]*?)</${p}:mathPr\\s*>)`).exec(
    xml,
  )
  if (!m) return null
  return {
    start: m.index,
    end: m.index + m[0].length,
    inner: m[3] ?? '',
    open: `<${prefix}:mathPr${m[1]}>`,
  }
}

function childRe(prefix: string, name: string): RegExp {
  const p = escapeRe(prefix)
  return new RegExp(`<${p}:${name}(?=[\\s/>])([^>]*?)(?:/>|>\\s*</${p}:${name}\\s*>)`)
}

/** the child's m:val: undefined when the child is absent, '' when it carries no value */
function childVal(inner: string, prefix: string, name: string): string | undefined {
  const m = childRe(prefix, name).exec(inner)
  if (!m) return undefined
  const val = new RegExp(`\\s${escapeRe(prefix)}:val\\s*=\\s*(["'])([^"']*)\\1`).exec(m[1])
  return val ? val[2] : ''
}

const isOff = (val: string) => ['0', 'false', 'off'].includes(val.toLowerCase())

function oneOf<T extends string>(val: string | undefined, allowed: readonly T[]): T | undefined {
  return allowed.find((a) => a === val)
}

function twips(val: string | undefined): number | undefined {
  if (val === undefined || !/^\d+$/.test(val)) return undefined
  return parseInt(val, 10)
}

/** m:mathPr of a settings part; undefined when the part has none */
export function readMathPr(settingsXml: string): MathPr | undefined {
  const prefix = mathPrefixOf(settingsXml) ?? 'm'
  const span = mathPrSpan(settingsXml, prefix)
  if (!span) return undefined
  const val = (name: string) => childVal(span.inner, prefix, name)
  const dispDef = val('dispDef')
  const smallFrac = val('smallFrac')
  const wrapRight = val('wrapRight')
  const mathFont = val('mathFont')
  const brkBin = oneOf(val('brkBin'), ['before', 'after', 'repeat'] as const)
  const brkBinSub = oneOf(val('brkBinSub'), ['--', '-+', '+-'] as const)
  const wrapIndent = twips(val('wrapIndent'))
  const limLocs = ['subSup', 'undOvr'] as const
  return {
    dispDef: dispDef === undefined || !isOff(dispDef),
    defJc:
      oneOf(val('defJc'), ['centerGroup', 'center', 'left', 'right'] as const) ??
      DEFAULT_MATH_PR.defJc,
    lMarginTwips: twips(val('lMargin')) ?? 0,
    rMarginTwips: twips(val('rMargin')) ?? 0,
    intLim: oneOf(val('intLim'), limLocs) ?? DEFAULT_MATH_PR.intLim,
    naryLim: oneOf(val('naryLim'), limLocs) ?? DEFAULT_MATH_PR.naryLim,
    ...(mathFont ? { mathFont } : {}),
    ...(brkBin ? { brkBin } : {}),
    ...(brkBinSub ? { brkBinSub } : {}),
    ...(smallFrac !== undefined ? { smallFrac: !isOff(smallFrac) } : {}),
    ...(wrapIndent !== undefined ? { wrapIndentTwips: wrapIndent } : {}),
    ...(wrapRight !== undefined ? { wrapRight: !isOff(wrapRight) } : {}),
  }
}

const CHILD_LOCAL_NAME = /<(?:[A-Za-z0-9._-]+:)?([A-Za-z0-9._-]+)[\s/>]/g

function setChild(inner: string, prefix: string, name: string, childXml: string): string {
  const existing = childRe(prefix, name).exec(inner)
  if (existing) {
    return (
      inner.slice(0, existing.index) + childXml + inner.slice(existing.index + existing[0].length)
    )
  }
  const rank = MATH_PR_CHILD_ORDER.indexOf(name)
  for (const m of inner.matchAll(CHILD_LOCAL_NAME)) {
    if (MATH_PR_CHILD_ORDER.indexOf(m[1]) > rank) {
      return inner.slice(0, m.index) + childXml + inner.slice(m.index)
    }
  }
  return inner + childXml
}

function dropChild(inner: string, prefix: string, name: string): string {
  return inner.replace(childRe(prefix, name), '')
}

/**
 * Write the edited fields into the part's m:mathPr, creating it through
 * `insertChild` (which knows the CT_Settings position) when the part has none.
 * Only a field whose value differs from what the part already says is written,
 * so every other child, and an unchanged part, stay byte-identical.
 */
export function saveMathPr(
  settingsXml: string,
  edit: MathPrEdit,
  insertChild: (xml: string, localName: string, childXml: string) => string,
): string {
  const current = readMathPr(settingsXml) ?? DEFAULT_MATH_PR
  const prefix = mathPrefixOf(settingsXml) ?? 'm'
  const el = (name: string, val?: string | number) =>
    val === undefined ? `<${prefix}:${name}/>` : `<${prefix}:${name} ${prefix}:val="${val}"/>`
  const writes: Array<[string, string | null]> = []
  if (edit.mathFont && edit.mathFont !== current.mathFont) {
    const font = edit.mathFont.replace(/[<>&"]/g, '')
    if (font) writes.push(['mathFont', el('mathFont', font)])
  }
  const brkBin = edit.brkBin ?? DEFAULT_MATH_BREAKS.brkBin
  if (brkBin !== (current.brkBin ?? DEFAULT_MATH_BREAKS.brkBin)) {
    writes.push(['brkBin', el('brkBin', brkBin)])
  }
  const brkBinSub = edit.brkBinSub ?? DEFAULT_MATH_BREAKS.brkBinSub
  if (brkBinSub !== (current.brkBinSub ?? DEFAULT_MATH_BREAKS.brkBinSub)) {
    writes.push(['brkBinSub', el('brkBinSub', brkBinSub)])
  }
  if (!!edit.smallFrac !== !!current.smallFrac) {
    writes.push(['smallFrac', edit.smallFrac ? el('smallFrac') : el('smallFrac', '0')])
  }
  if (edit.dispDef !== current.dispDef) {
    writes.push(['dispDef', edit.dispDef ? el('dispDef') : el('dispDef', '0')])
  }
  const margin = (v: number) => Math.max(0, Math.round(v))
  if (margin(edit.lMarginTwips) !== current.lMarginTwips) {
    writes.push(['lMargin', el('lMargin', margin(edit.lMarginTwips))])
  }
  if (margin(edit.rMarginTwips) !== current.rMarginTwips) {
    writes.push(['rMargin', el('rMargin', margin(edit.rMarginTwips))])
  }
  if (edit.defJc !== current.defJc) writes.push(['defJc', el('defJc', edit.defJc)])
  const wrapIndent = (pr: Pick<MathPr, 'wrapIndentTwips'>) =>
    Math.max(0, Math.round(pr.wrapIndentTwips ?? DEFAULT_MATH_BREAKS.wrapIndentTwips))
  if (!!edit.wrapRight !== !!current.wrapRight) {
    // one choice in CT_MathPr: the other child leaves
    if (edit.wrapRight) writes.push(['wrapIndent', null], ['wrapRight', el('wrapRight')])
    else writes.push(['wrapRight', null], ['wrapIndent', el('wrapIndent', wrapIndent(edit))])
  } else if (!edit.wrapRight && wrapIndent(edit) !== wrapIndent(current)) {
    writes.push(['wrapIndent', el('wrapIndent', wrapIndent(edit))])
  }
  if (edit.intLim !== current.intLim) writes.push(['intLim', el('intLim', edit.intLim)])
  if (edit.naryLim !== current.naryLim) writes.push(['naryLim', el('naryLim', edit.naryLim)])
  if (writes.length === 0) return settingsXml

  const span = mathPrSpan(settingsXml, prefix)
  let inner = span?.inner ?? ''
  for (const [name, childXml] of writes) {
    inner =
      childXml === null ? dropChild(inner, prefix, name) : setChild(inner, prefix, name, childXml)
  }
  if (span) {
    return (
      settingsXml.slice(0, span.start) +
      `${span.open}${inner}</${prefix}:mathPr>` +
      settingsXml.slice(span.end)
    )
  }
  if (mathPrefixOf(settingsXml)) {
    return insertChild(settingsXml, 'mathPr', `<${prefix}:mathPr>${inner}</${prefix}:mathPr>`)
  }
  // a root that binds m: to something else keeps it: the element declares the namespace itself
  const rootBindsM = /<(?:[A-Za-z0-9._-]+:)?settings\b[^>]*\sxmlns:m\s*=/.test(settingsXml)
  if (rootBindsM) {
    return insertChild(settingsXml, 'mathPr', `<m:mathPr xmlns:m="${MATH_NS}">${inner}</m:mathPr>`)
  }
  const declared = settingsXml.replace(
    /<((?:[A-Za-z0-9._-]+:)?settings)\b/,
    `<$1 xmlns:m="${MATH_NS}"`,
  )
  return insertChild(declared, 'mathPr', `<m:mathPr>${inner}</m:mathPr>`)
}
