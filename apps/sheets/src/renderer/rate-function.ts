/**
 * RATE bisection fallback. Univer's Newton solver gives up the moment
 * an iterate crosses rate <= -1, so deep-negative solutions (each period
 * losing nearly everything, e.g. CAGR of a collapsing series) come back
 * #NUM! even though a root exists just above -1. When the builtin errors,
 * re-solve by bracketing on (-1, 1e9] and bisecting; only a genuinely
 * root-less equation stays #NUM!.
 */
import {
  BaseFunction,
  ErrorType,
  FUNCTION_NAMES_FINANCIAL,
  IFunctionService,
  NumberValueObject,
  functionFinancial,
} from '@univerjs/engine-formula'

import type { UniverRuntime } from './univer-state'

function annuityResidual(nper: number, pmt: number, pv: number, fv: number, type: number) {
  return (rate: number): number => {
    if (Math.abs(rate) < 1e-12) return pv + pmt * nper + fv
    const growth = (1 + rate) ** nper
    return pv * growth + pmt * (1 / rate + type) * (growth - 1) + fv
  }
}

/// Scan grid: dense toward -1 (where the collapsing-value roots live), then
/// coarse toward large positive rates.
const SCAN_POINTS: readonly number[] = [
  ...Array.from({ length: 12 }, (unused, at) => -1 + 10 ** (at - 12)),
  -0.8,
  -0.6,
  -0.4,
  -0.2,
  -0.05,
  0.05,
  0.2,
  0.5,
  1,
  2,
  5,
  10,
  100,
  1e4,
  1e6,
  1e9,
]

export function solveRateByBisection(
  nper: number,
  pmt: number,
  pv: number,
  fv: number,
  type: number,
): number | null {
  const residual = annuityResidual(nper, pmt, pv, fv, type)
  let previous: { rate: number; value: number } | null = null
  for (const rate of SCAN_POINTS) {
    const value = residual(rate)
    if (!Number.isFinite(value)) {
      previous = null
      continue
    }
    if (value === 0) return rate
    if (previous && Math.sign(value) !== Math.sign(previous.value)) {
      return bisect(residual, previous.rate, rate, previous.value)
    }
    previous = { rate, value }
  }
  return null
}

function bisect(
  residual: (rate: number) => number,
  low: number,
  high: number,
  lowValue: number,
): number {
  let a = low
  let b = high
  let fa = lowValue
  for (let i = 0; i < 200 && b - a > Number.EPSILON * (1 + Math.abs(a)); i += 1) {
    const mid = (a + b) / 2
    const fm = residual(mid)
    if (fm === 0) return mid
    if (Math.sign(fm) === Math.sign(fa)) {
      a = mid
      fa = fm
    } else {
      b = mid
    }
  }
  return (a + b) / 2
}

interface RateResultLike {
  isError?(): boolean
  getValue(): unknown
}

/// The builtin keeps its solver in _getResult(nper, pmt, pv, fv, type,
/// guess, rowIndex, columnIndex); the subclass reuses all of its argument
/// plumbing and only steps in when that solver says #NUM!.
export function createRateExecutor(): BaseFunction | null {
  const RateBuiltin = functionFinancial.find(
    ([, name]) => name === FUNCTION_NAMES_FINANCIAL.RATE,
  )?.[0] as (new (name: string) => BaseFunction) | undefined
  const builtinSolve = (RateBuiltin?.prototype as Record<string, unknown> | undefined)?.[
    '_getResult'
  ] as ((...args: number[]) => RateResultLike) | undefined
  if (!RateBuiltin || typeof builtinSolve !== 'function') return null
  class RateWithBisection extends RateBuiltin {
    _getResult(
      nper: number,
      pmt: number,
      pv: number,
      fv: number,
      type: number,
      guess: number,
      rowIndex: number,
      columnIndex: number,
    ): RateResultLike {
      const result = builtinSolve!.call(this, nper, pmt, pv, fv, type, guess, rowIndex, columnIndex)
      if (!(result.isError?.() && result.getValue() === ErrorType.NUM)) return result
      const root = solveRateByBisection(nper, pmt, pv, fv, type)
      if (root === null) return result
      return NumberValueObject.create(root, rowIndex === 0 && columnIndex === 0 ? '0%' : undefined)
    }
  }
  return new RateWithBisection(FUNCTION_NAMES_FINANCIAL.RATE)
}

export function installRateFallback(runtime: UniverRuntime): { dispose(): void } {
  const functionService = runtime.univer.__getInjector().get(IFunctionService)
  const name = FUNCTION_NAMES_FINANCIAL.RATE
  const original = functionService.getExecutor(name) ?? undefined
  const executor = createRateExecutor()
  if (!executor) return { dispose() {} }
  functionService.registerExecutors(executor)
  const timer = setTimeout(() => {
    if (functionService.getExecutor(name) !== executor) {
      functionService.registerExecutors(executor)
    }
  }, 0)
  return {
    dispose() {
      clearTimeout(timer)
      if (original) functionService.registerExecutors(original)
    },
  }
}
