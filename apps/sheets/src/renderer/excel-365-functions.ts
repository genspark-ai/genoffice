/**
 * Microsoft 365 functions with Excel's signatures. Univer ships REGEXEXTRACT
 * and REGEXREPLACE with the Google Sheets arity (2 and 3 arguments, capture
 * groups as the extract result) and has no REGEXTEST or PERCENTOF, so an
 * Excel workbook using them recalculates to #NAME? or an argument error.
 * Excel's patterns are PCRE2; JavaScript RegExp covers the common subset,
 * plus a leading (?i) which PCRE2 users write for case-insensitive matching.
 */
import {
  ArrayValueObject,
  BaseFunction,
  BooleanValueObject,
  ErrorType,
  ErrorValueObject,
  FunctionType,
  IFunctionService,
  NumberValueObject,
  StringValueObject,
} from '@univerjs/engine-formula'
import type { BaseValueObject, IFunctionInfo } from '@univerjs/engine-formula'

import type { UniverRuntime } from './univer-state'

const LEADING_CASE_FLAG = /^\(\?i\)/

type Result = BaseValueObject

function scalar(value: BaseValueObject | undefined): BaseValueObject | undefined {
  if (!value?.isArray()) return value
  return (value as unknown as ArrayValueObject).getFirstCell()
}

function textOf(value: BaseValueObject): string {
  if (value.isNull()) return ''
  if (value.isBoolean()) return value.getValue() ? 'TRUE' : 'FALSE'
  return `${value.getValue()}`
}

/// Optional integer argument: undefined or blank means the default, and a
/// value that is not a number is #VALUE!.
function option(value: BaseValueObject | undefined, fallback: number): number | Result {
  const cell = scalar(value)
  if (!cell || cell.isNull()) return fallback
  if (cell.isError()) return cell
  const number = Number(cell.isBoolean() ? (cell.getValue() ? 1 : 0) : cell.getValue())
  if (!Number.isFinite(number)) return ErrorValueObject.create(ErrorType.VALUE)
  return Math.trunc(number)
}

function compile(patternValue: BaseValueObject, caseValue: number): RegExp | Result {
  let pattern = textOf(patternValue)
  let flags = 'g'
  if (LEADING_CASE_FLAG.test(pattern)) {
    pattern = pattern.replace(LEADING_CASE_FLAG, '')
    flags += 'i'
  } else if (caseValue === 1) {
    flags += 'i'
  }
  try {
    return new RegExp(pattern, flags)
  } catch {
    return ErrorValueObject.create(ErrorType.VALUE)
  }
}

function isResult(value: unknown): value is Result {
  return typeof value === 'object' && value !== null && !(value instanceof RegExp)
}

/// Text arrays apply elementwise, like Excel's lifting of scalar arguments.
function perText(text: BaseValueObject, each: (cell: BaseValueObject) => Result): Result {
  if (!text.isArray()) return text.isError() ? text : each(text)
  const mapped = (text as unknown as ArrayValueObject).mapValue((cell) =>
    cell.isError() ? cell : each(cell),
  )
  if (mapped.getRowCount() === 1 && mapped.getColumnCount() === 1) return mapped.get(0, 0)!
  return mapped
}

function row(values: string[], fn: BaseFunction): Result {
  if (values.length === 1) return StringValueObject.create(values[0]!)
  return ArrayValueObject.create({
    calculateValueList: [values.map((value) => StringValueObject.create(value))],
    rowCount: 1,
    columnCount: values.length,
    unitId: fn.unitId ?? '',
    sheetId: fn.subUnitId ?? '',
    row: fn.row,
    column: fn.column,
  })
}

export class RegexTest extends BaseFunction {
  override minParams = 2
  override maxParams = 3

  override calculate(
    text: BaseValueObject,
    pattern: BaseValueObject,
    caseSensitivity?: BaseValueObject,
  ): Result {
    const patternCell = scalar(pattern)!
    if (patternCell.isError()) return patternCell
    const caseValue = option(caseSensitivity, 0)
    if (isResult(caseValue)) return caseValue
    const regExp = compile(patternCell, caseValue)
    if (isResult(regExp)) return regExp
    return perText(text, (cell) => {
      regExp.lastIndex = 0
      return BooleanValueObject.create(regExp.test(textOf(cell)))
    })
  }
}

export class RegexExtract extends BaseFunction {
  override minParams = 2
  override maxParams = 4

  override calculate(
    text: BaseValueObject,
    pattern: BaseValueObject,
    returnMode?: BaseValueObject,
    caseSensitivity?: BaseValueObject,
  ): Result {
    const patternCell = scalar(pattern)!
    if (patternCell.isError()) return patternCell
    const mode = option(returnMode, 0)
    if (isResult(mode)) return mode
    if (mode < 0 || mode > 2) return ErrorValueObject.create(ErrorType.VALUE)
    const caseValue = option(caseSensitivity, 0)
    if (isResult(caseValue)) return caseValue
    const regExp = compile(patternCell, caseValue)
    if (isResult(regExp)) return regExp
    return perText(text, (cell) => {
      const value = textOf(cell)
      if (mode === 1) {
        const all = [...value.matchAll(regExp)].map((match) => match[0])
        return all.length === 0 ? ErrorValueObject.create(ErrorType.NA) : row(all, this)
      }
      regExp.lastIndex = 0
      const first = regExp.exec(value)
      if (!first) return ErrorValueObject.create(ErrorType.NA)
      if (mode === 0) return StringValueObject.create(first[0])
      const groups = first.slice(1).map((group) => group ?? '')
      return groups.length === 0 ? ErrorValueObject.create(ErrorType.NA) : row(groups, this)
    })
  }
}

export class RegexReplace extends BaseFunction {
  override minParams = 3
  override maxParams = 5

  override calculate(
    text: BaseValueObject,
    pattern: BaseValueObject,
    replacement: BaseValueObject,
    occurrence?: BaseValueObject,
    caseSensitivity?: BaseValueObject,
  ): Result {
    const patternCell = scalar(pattern)!
    if (patternCell.isError()) return patternCell
    const replacementCell = scalar(replacement)!
    if (replacementCell.isError()) return replacementCell
    const nth = option(occurrence, 0)
    if (isResult(nth)) return nth
    const caseValue = option(caseSensitivity, 0)
    if (isResult(caseValue)) return caseValue
    const regExp = compile(patternCell, caseValue)
    if (isResult(regExp)) return regExp
    const replaceWith = textOf(replacementCell)
    return perText(text, (cell) => {
      const value = textOf(cell)
      if (nth === 0) return StringValueObject.create(value.replace(regExp, replaceWith))
      const matches = [...value.matchAll(regExp)]
      const target = matches[nth > 0 ? nth - 1 : matches.length + nth]
      if (!target || target.index === undefined) return StringValueObject.create(value)
      const single = new RegExp(regExp.source, regExp.flags.replace('g', ''))
      const replaced = target[0].replace(single, replaceWith)
      return StringValueObject.create(
        value.slice(0, target.index) + replaced + value.slice(target.index + target[0].length),
      )
    })
  }
}

/// PERCENTOF(data_subset, data_all) is SUM(data_subset) / SUM(data_all).
export class PercentOf extends BaseFunction {
  override minParams = 2
  override maxParams = 2

  override calculate(subset: BaseValueObject, all: BaseValueObject): Result {
    const part = total(subset)
    if (part.isError()) return part
    const whole = total(all)
    if (whole.isError()) return whole
    if (Number(whole.getValue()) === 0) return ErrorValueObject.create(ErrorType.DIV_BY_ZERO)
    return NumberValueObject.create(Number(part.getValue()) / Number(whole.getValue()))
  }
}

function total(value: BaseValueObject): BaseValueObject {
  let variant = value
  if (variant.isString()) variant = variant.convertToNumberObjectValue()
  if (variant.isError()) return variant
  if (variant.isArray()) variant = (variant as unknown as ArrayValueObject).sum()
  if (variant.isNull()) return NumberValueObject.create(0)
  return variant
}

function param(name: string, detail: string, require = true) {
  return { name, detail, example: '', require: require ? 1 : 0, repeat: 0 } as const
}

const CASE_PARAM = param(
  'case_sensitivity',
  '0 for case-sensitive (default), 1 for case-insensitive.',
  false,
)

export const EXCEL_365_DESCRIPTIONS: IFunctionInfo[] = [
  {
    functionName: 'REGEXTEST',
    functionType: FunctionType.Text,
    description: 'Checks whether any part of the text matches the regular expression.',
    abstract: 'Checks whether text matches a pattern',
    functionParameter: [
      param('text', 'The text to test.'),
      param('pattern', 'The regular expression.'),
      CASE_PARAM,
    ],
  },
  {
    functionName: 'REGEXEXTRACT',
    functionType: FunctionType.Text,
    description: 'Extracts the text that matches the regular expression.',
    abstract: 'Extracts text that matches a pattern',
    functionParameter: [
      param('text', 'The text to extract from.'),
      param('pattern', 'The regular expression.'),
      param(
        'return_mode',
        '0 for the first match (default), 1 for all matches, 2 for the capture groups of the first match.',
        false,
      ),
      CASE_PARAM,
    ],
  },
  {
    functionName: 'REGEXREPLACE',
    functionType: FunctionType.Text,
    description:
      'Replaces the text that matches the regular expression. Use $1, $2 for capture groups.',
    abstract: 'Replaces text that matches a pattern',
    functionParameter: [
      param('text', 'The text to replace in.'),
      param('pattern', 'The regular expression.'),
      param('replacement', 'The replacement text.'),
      param(
        'occurrence',
        '0 replaces every match (default), n the nth match, -n the nth from the end.',
        false,
      ),
      CASE_PARAM,
    ],
  },
  {
    functionName: 'PERCENTOF',
    functionType: FunctionType.Math,
    description: 'Returns the sum of a subset of values as a fraction of the sum of all values.',
    abstract: 'Returns the percentage a subset makes up',
    functionParameter: [
      param('data_subset', 'The values in the subset.'),
      param('data_all', 'All the values.'),
    ],
  },
]

export function createExcel365Executors(): BaseFunction[] {
  return [
    new RegexTest('REGEXTEST'),
    new RegexExtract('REGEXEXTRACT'),
    new RegexReplace('REGEXREPLACE'),
    new PercentOf('PERCENTOF'),
  ]
}

export function installExcel365Functions(runtime: UniverRuntime): { dispose(): void } {
  const functionService = runtime.univer.__getInjector().get(IFunctionService)
  const executors = createExcel365Executors()
  const originals = executors.map(
    (executor) => functionService.getExecutor(executor.name) ?? undefined,
  )
  const originalDescriptions = EXCEL_365_DESCRIPTIONS.map(
    (info) => functionService.getDescription(info.functionName) ?? undefined,
  )
  const register = () => {
    const stale = executors.filter(
      (executor) => functionService.getExecutor(executor.name) !== executor,
    )
    if (stale.length > 0) functionService.registerExecutors(...stale)
    functionService.registerDescriptions(...EXCEL_365_DESCRIPTIONS)
  }
  register()
  // The engine registers its builtins (and the sheets-formula plugin its
  // descriptions) at onReady, which can run after this installer.
  const timer = setTimeout(register, 0)
  return {
    dispose() {
      clearTimeout(timer)
      originals.forEach((original, at) => {
        if (original) functionService.registerExecutors(original)
        else functionService.unregisterExecutors(executors[at]!.name)
      })
      originalDescriptions.forEach((info, at) => {
        if (info) functionService.registerDescriptions(info)
        else functionService.unregisterDescriptions(EXCEL_365_DESCRIPTIONS[at]!.functionName)
      })
    },
  }
}
