import { NumberValueObject } from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import { createRateExecutor, solveRateByBisection } from '../src/renderer/rate-function'

function residual(nper: number, pmt: number, pv: number, fv: number, rate: number): number {
  const growth = (1 + rate) ** nper
  return pv * growth + pmt * (1 / rate + 0) * (growth - 1) + fv
}

describe('solveRateByBisection', () => {
  it('finds deep-negative roots the Newton solver overshoots', () => {
    const root = solveRateByBisection(5, -10, -100, 10.0000001, 0)
    expect(root).not.toBeNull()
    expect(root!).toBeGreaterThan(-1)
    expect(root!).toBeLessThan(-0.99)
    expect(Math.abs(residual(5, -10, -100, 10.0000001, root!))).toBeLessThan(1e-6)
  })

  it('returns null when the equation has no root', () => {
    expect(solveRateByBisection(3, -10, -100, 9, 0)).toBeNull()
  })
})

describe('RATE executor with bisection fallback', () => {
  const executor = createRateExecutor()
  const n = (value: number) => NumberValueObject.create(value)
  const run = (...args: number[]) =>
    (executor!.calculate as (...v: unknown[]) => { getValue(): unknown })(...args.map(n)).getValue()

  it('builds on the current Univer bundle', () => {
    expect(executor).not.toBeNull()
  })

  it('keeps ordinary solutions identical to the builtin', () => {
    expect(run(10, -100, 1000)).toBeCloseTo(0, 10)
    expect(run(5, 10, -100, 0.0001) as number).toBeCloseTo(-0.194017827877, 9)
  })

  it('converges near -100% instead of #NUM! (issue regression)', () => {
    // =RATE(5,0,-100,0.0001): each period loses ~94%
    expect(run(5, 0, -100, 0.0001) as number).toBeCloseTo(-0.93690426555, 9)
  })

  it('rescues pmt≠0 cases where Newton walks past -1', () => {
    const value = run(5, -10, -100, 10.0000001)
    expect(typeof value).toBe('number')
    expect(value as number).toBeLessThan(-0.99)
    expect(value as number).toBeGreaterThan(-1)
  })

  it('still errors when no rate exists', () => {
    expect(run(3, -10, -100, 9)).toBe('#NUM!')
  })
})
