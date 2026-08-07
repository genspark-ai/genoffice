/// Declarative data-validation save: the renderer snapshots the full Univer
/// rule set of a dirty sheet and this module rewrites the worksheet's
/// `<dataValidations>` section from it (mirroring the CF/filter recipe).
/// Mappings are the exact inverse of the read-side install in App.tsx.

export class DvEditError extends Error {}

export interface DvCellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/// One rule in the Univer data-validation model shape (validated structurally
/// here — unknown shapes fail the save rather than guess).
export interface DvWireRule {
  readonly ranges: readonly DvCellArea[]
  readonly rule: Record<string, unknown>
}

const DV_TYPES = new Set(['whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'])
const DV_OPERATORS = new Set([
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])
/// Univer DataValidationErrorStyle: INFO=0, STOP=1 (OOXML default), WARNING=2.
const DV_ERROR_STYLE_NAMES: Record<number, string | undefined> = {
  0: 'information',
  1: undefined,
  2: 'warning',
}

export function applyDvRules(worksheetXml: string, rules: readonly DvWireRule[]): string {
  if (/<x14:dataValidation\b/.test(worksheetXml)) {
    throw new DvEditError(
      'This sheet has extended (x14) data validation — editing its rules is not ' +
        'supported yet.',
    )
  }
  const xml = worksheetXml.replace(
    /<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>|<dataValidations\b[^>]*\/>/g,
    '',
  )
  if (rules.length === 0) return xml

  const body = rules.map(serializeRule).join('')
  const section = `<dataValidations count="${rules.length}">${body}</dataValidations>`
  const anchor =
    /<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<picture\b|<oleObjects\b|<tableParts\b|<extLst\b/.exec(
      xml,
    )
  if (anchor) {
    return xml.slice(0, anchor.index) + section + xml.slice(anchor.index)
  }
  const end = xml.lastIndexOf('</worksheet>')
  if (end === -1) throw new DvEditError('Worksheet has no closing element.')
  return xml.slice(0, end) + section + xml.slice(end)
}

function serializeRule(wireRule: DvWireRule): string {
  if (wireRule.ranges.length === 0) {
    throw new DvEditError('A data-validation rule has no ranges.')
  }
  let rule = wireRule.rule
  let rawType = String(rule.type ?? '')
  if (rawType === 'listMultiple') {
    throw new DvEditError(
      'Multi-select list rules are Univer-only and cannot be saved to xlsx — ' +
        'delete the rule before saving.',
    )
  }
  if (rawType === 'checkbox') {
    // Checkbox is Univer-only; degrade to a two-value list so Excel keeps the
    // constraint. The default 1/0 pair round-trips back to a checkbox on load.
    const checked =
      rule.formula1 === undefined || rule.formula1 === '' ? '1' : String(rule.formula1)
    const unchecked =
      rule.formula2 === undefined || rule.formula2 === '' ? '0' : String(rule.formula2)
    rule = {
      ...rule,
      type: 'list',
      operator: undefined,
      formula1: `${checked},${unchecked}`,
      formula2: undefined,
    }
    rawType = 'list'
  }
  // 'any' is the read-side mapping of OOXML type="none" (no constraint, just
  // messages); it round-trips back to the default, attribute-less form.
  const type = rawType === 'any' || rawType === 'none' ? undefined : rawType
  if (type !== undefined && !DV_TYPES.has(type)) {
    throw new DvEditError(`Unsupported data-validation type "${rawType}".`)
  }

  const attrs: string[] = []
  if (type !== undefined) attrs.push(`type="${type}"`)
  const operator = rule.operator === undefined ? undefined : String(rule.operator)
  if (operator !== undefined && operator !== '') {
    if (!DV_OPERATORS.has(operator)) {
      throw new DvEditError(`Unsupported data-validation operator "${operator}".`)
    }
    // "between" is the OOXML default; only operator-carrying types keep it.
    if (operator !== 'between' && type !== undefined && type !== 'list' && type !== 'custom') {
      attrs.push(`operator="${operator}"`)
    }
  }
  if (rule.allowBlank === true) attrs.push('allowBlank="1"')
  // OOXML's showDropDown="1" SUPPRESSES the in-cell dropdown (inverted name);
  // Univer's showDropDown means what it says.
  if (type === 'list' && rule.showDropDown === false) attrs.push('showDropDown="1"')
  if (rule.showInputMessage === true) attrs.push('showInputMessage="1"')
  if (rule.showErrorMessage === true) attrs.push('showErrorMessage="1"')
  const errorStyle = errorStyleName(rule.errorStyle)
  if (errorStyle !== undefined) attrs.push(`errorStyle="${errorStyle}"`)
  for (const [key, attribute] of [
    ['errorTitle', 'errorTitle'],
    ['error', 'error'],
    ['promptTitle', 'promptTitle'],
    ['prompt', 'prompt'],
  ] as const) {
    const value = rule[key]
    if (typeof value === 'string' && value.length > 0) {
      attrs.push(`${attribute}="${escapeXmlAttribute(value)}"`)
    }
  }
  attrs.push(`sqref="${wireRule.ranges.map(toRef).join(' ')}"`)

  const formulas = serializeFormulas(type, rule)
  return formulas === ''
    ? `<dataValidation ${attrs.join(' ')}/>`
    : `<dataValidation ${attrs.join(' ')}>${formulas}</dataValidation>`
}

function errorStyleName(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined
  const style = Number(value)
  if (!(style in DV_ERROR_STYLE_NAMES)) {
    throw new DvEditError(`Unsupported data-validation error style "${String(value)}".`)
  }
  return DV_ERROR_STYLE_NAMES[style]
}

function serializeFormulas(type: string | undefined, rule: Record<string, unknown>): string {
  const formula1 = formulaText(type, rule.formula1)
  const formula2 = formulaText(type, rule.formula2)
  return (
    (formula1 === undefined ? '' : `<formula1>${escapeXmlText(formula1)}</formula1>`) +
    (formula2 === undefined ? '' : `<formula2>${escapeXmlText(formula2)}</formula2>`)
  )
}

/// Inverse of the install-side formula transforms: list literals regain their
/// quotes, `=`-prefixed references/formulas lose the prefix, and panel-edited
/// date/time strings become Excel serial numbers.
function formulaText(type: string | undefined, raw: unknown): string | undefined {
  if (raw === undefined || raw === null) return undefined
  const text = String(raw)
  if (text === '') return undefined
  if (type === 'list') {
    return text.startsWith('=') ? text.slice(1) : `"${text}"`
  }
  if (type === 'custom') {
    return text.startsWith('=') ? text.slice(1) : text
  }
  if (type === 'date') {
    const serial = dateToSerial(text)
    if (serial !== undefined) return String(serial)
  }
  if (type === 'time') {
    const fraction = timeToFraction(text)
    if (fraction !== undefined) return String(fraction)
  }
  return text.startsWith('=') ? text.slice(1) : text
}

/// 'YYYY-MM-DD[ HH:mm[:ss]]' (or slashes) → Excel serial (days since
/// 1899-12-30). Plain numbers and references pass through untouched.
function dateToSerial(text: string): number | undefined {
  const match =
    /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/.exec(
      text.trim(),
    )
  if (!match) return undefined
  const [, year, month, day, hour, minute, second] = match
  const days =
    (Date.UTC(Number(year), Number(month) - 1, Number(day)) - Date.UTC(1899, 11, 30)) / 86_400_000
  const seconds = Number(hour ?? 0) * 3600 + Number(minute ?? 0) * 60 + Number(second ?? 0)
  return seconds === 0 ? days : days + seconds / 86_400
}

function timeToFraction(text: string): number | undefined {
  const match = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/.exec(text.trim())
  if (!match) return undefined
  const [, hour, minute, second] = match
  return (Number(hour) * 3600 + Number(minute) * 60 + Number(second ?? 0)) / 86_400
}

function toRef(range: DvCellArea): string {
  const start = `${columnToLetters(range.startColumn)}${range.startRow + 1}`
  return range.startRow === range.endRow && range.startColumn === range.endColumn
    ? start
    : `${start}:${columnToLetters(range.endColumn)}${range.endRow + 1}`
}

function columnToLetters(column: number): string {
  let letters = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26)
  }
  return letters
}

function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll('\n', '&#10;')
}
