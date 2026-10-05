import { describe, expect, it } from 'vitest'

import { scatterAxisBounds, valueAxisScale } from '../src/domain/chart-visual'

/// A value axis that cannot show its own data: an all-negative series used to
/// scale 0..1, because `valueAxisScale` only ever saw the maximum and derived
/// the lower bound as `min = explicit?.min ?? 0`.
describe('valueAxisScale with below-zero data', () => {
  it('scales an all-negative series around zero instead of 0..1', () => {
    // Renderers floor the maximum at 0, so an all-negative chart arrives as
    // dataMax 0; the minimum is what tells the axis to go down.
    const axis = valueAxisScale(0, undefined, -500)
    expect(axis.min).toBe(-500)
    expect(axis.max).toBe(0)
    expect(axis.min).toBeLessThanOrEqual(-500)
    expect(axis.max).toBeGreaterThanOrEqual(-100)
  })

  it('keeps every all-negative value inside the axis', () => {
    const values = [-500, -300, -100]
    const axis = valueAxisScale(Math.max(...values), undefined, Math.min(...values))
    for (const value of values) {
      expect(axis.min).toBeLessThanOrEqual(value)
      expect(axis.max).toBeGreaterThanOrEqual(value)
    }
  })

  it('derives a negative floor from a negative maximum alone', () => {
    // No dataMin supplied: the maximum still cannot sit above the axis.
    const axis = valueAxisScale(-100)
    expect(axis.min).toBeLessThanOrEqual(-100)
    expect(axis.max).toBe(0)
  })

  it('matches scatterAxisBounds on an all-negative series', () => {
    // scatterAxisBounds is the sibling that already handled negatives via
    // -niceCeiling(-dataMin); the two must not drift apart.
    const axis = valueAxisScale(0, undefined, -500)
    const scatter = scatterAxisBounds([-500, -300, -100])
    expect(axis.min).toBe(scatter.min)
    expect(axis.max).toBe(scatter.max)
  })

  it('scales mixed positive/negative data across zero', () => {
    const axis = valueAxisScale(300, undefined, -500)
    expect(axis.min).toBeLessThanOrEqual(-500)
    expect(axis.max).toBeGreaterThanOrEqual(300)
    expect(axis.min).toBeLessThan(0)
    expect(axis.max).toBeGreaterThan(0)
  })

  it('leaves an all-positive series on the original 0-based scale', () => {
    // Unchanged behaviour: no dataMin argument at all.
    expect(valueAxisScale(877)).toEqual({
      min: 0,
      max: 1000,
      ticks: [0, 100, 200, 300, 400, 500, 600, 700, 800, 900, 1000],
    })
  })

  it('leaves a positive series untouched when its minimum is passed', () => {
    // dataMin >= 0 must not move the axis off 0.
    expect(valueAxisScale(877, undefined, 10)).toEqual(valueAxisScale(877))
    expect(valueAxisScale(18, undefined, 0)).toEqual({
      min: 0,
      max: 20,
      ticks: [0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20],
    })
  })

  it('gives a constant negative series a range instead of a zero-width axis', () => {
    const axis = valueAxisScale(0, undefined, -5)
    expect(axis.max).toBeGreaterThan(axis.min)
    expect(axis.min).toBeLessThanOrEqual(-5)
    expect(axis.max).toBeGreaterThanOrEqual(-5)
    // Same shape as the sibling for the same constant input.
    const scatter = scatterAxisBounds([-5, -5, -5])
    expect(axis.min).toBe(scatter.min)
    expect(axis.max).toBe(scatter.max)
  })

  it('keeps the flat 0..1 axis for all-zero data', () => {
    expect(valueAxisScale(0)).toEqual({
      min: 0,
      max: 1,
      ticks: [0, 0.2, 0.4, 0.6, 0.8, 1],
    })
    expect(valueAxisScale(0, undefined, 0)).toEqual(valueAxisScale(0))
  })

  it('honours an explicit min/max over the data minimum', () => {
    const axis = valueAxisScale(0, { min: -200, max: 50 }, -500)
    expect(axis.min).toBe(-200)
    expect(axis.max).toBe(50)
  })

  it('still honours an explicit majorUnit below zero', () => {
    const axis = valueAxisScale(0, { majorUnit: 100 }, -300)
    expect(axis.min).toBe(-300)
    expect(axis.max).toBe(0)
    expect(axis.ticks).toEqual([-300, -200, -100, 0])
  })

  it('does not scale to -Infinity when a caller has no values', () => {
    const axis = valueAxisScale(Math.max(...[], 0), undefined, Math.min(...[]))
    expect(Number.isFinite(axis.min)).toBe(true)
    expect(Number.isFinite(axis.max)).toBe(true)
  })
})
