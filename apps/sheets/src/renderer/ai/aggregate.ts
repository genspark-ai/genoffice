/// Streaming range statistics for the aggregate_range AI tool: the answer to
/// "how many distinct suppliers" on an 88k-row sheet must come from batched
/// reads, never from read_range loops (token cost) or COUNTIF-style array
/// formulas (quadratic main-thread evaluation — the app-freeze incident).

import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'

/// Distinct tracking stops (and the result says so) past this many unique
/// values; counts and numeric stats stay exact.
const MAX_DISTINCT_TRACKED = 200_000

export interface RangeAggregate {
  readonly cells: number
  readonly nonEmpty: number
  /** null when distinct tracking overflowed MAX_DISTINCT_TRACKED */
  readonly distinct: number | null
  readonly numericCount: number
  readonly sum: number
  readonly min: number | null
  readonly max: number | null
  readonly average: number | null
  /** most frequent values, descending; empty when tracking overflowed */
  readonly topValues: readonly { value: string; count: number }[]
  /**
   * Cells the reader withheld. They are counted here and nowhere else: every
   * statistic above is computed over the remaining cells only.
   *
   * A withheld cell cannot simply be masked in the output. `sum` is exact, so
   * one withheld cell inside a column is solvable — sum minus everything else
   * gives the very value the reader hid. Excluding it from the arithmetic is
   * the only fix; the count is then reported honestly so the model does not
   * present a total as if the column were whole.
   */
  readonly withheld: number
}

export interface RangeAggregator {
  add(value: CellScalar): void
  addRepeated(value: CellScalar, count: number): void
  addEmpty(count: number): void
  /** record `count` cells the reader withheld; they contribute to nothing else */
  addWithheld(count: number): void
  finish(topValueCount: number): RangeAggregate
}

export function createRangeAggregator(): RangeAggregator {
  let cells = 0
  let nonEmpty = 0
  let numericCount = 0
  let sum = 0
  let min: number | null = null
  let max: number | null = null
  let overflowed = false
  let withheld = 0
  const counts = new Map<string, number>()

  const addRepeated = (value: CellScalar, count: number): void => {
    if (count <= 0) return
    cells += count
    if (value === null || value === '') return
    nonEmpty += count
    if (typeof value === 'number') {
      numericCount += count
      sum += value * count
      min = min === null ? value : Math.min(min, value)
      max = max === null ? value : Math.max(max, value)
    }
    if (overflowed) return
    const key = `${typeof value}:${String(value)}`
    const existing = counts.get(key)
    if (existing !== undefined) {
      counts.set(key, existing + count)
    } else if (counts.size >= MAX_DISTINCT_TRACKED) {
      overflowed = true
      counts.clear()
    } else {
      counts.set(key, count)
    }
  }

  return {
    add(value: CellScalar): void {
      addRepeated(value, 1)
    },
    addRepeated,
    addEmpty(count: number): void {
      if (count <= 0) return
      cells += count
    },
    addWithheld(count: number): void {
      if (count <= 0) return
      // Counted as a cell so the range's own total stays honest, but never as
      // a value: nonEmpty, distinct, the numeric stats and the frequency table
      // all have to be blind to it or the number can be back-solved.
      cells += count
      withheld += count
    },
    finish(topValueCount: number): RangeAggregate {
      const topValues = overflowed
        ? []
        : [...counts.entries()]
            .sort((left, right) => right[1] - left[1])
            .slice(0, Math.max(0, topValueCount))
            .map(([key, count]) => ({ value: key.slice(key.indexOf(':') + 1), count }))
      return {
        cells,
        nonEmpty,
        distinct: overflowed ? null : counts.size,
        numericCount,
        sum,
        min,
        max,
        average: numericCount > 0 ? sum / numericCount : null,
        topValues,
        withheld,
      }
    },
  }
}

const formatNumber = (value: number): string =>
  Number.isInteger(value)
    ? value.toLocaleString('en-US')
    : value.toLocaleString('en-US', { maximumFractionDigits: 4 })

/** Tool-output rendering (English, like all other tool outputs). */
export function formatRangeAggregate(
  rangeLabel: string,
  aggregate: RangeAggregate,
  topValueCount: number,
): string {
  const lines = [
    `Statistics for ${rangeLabel}:`,
    `- cells: ${formatNumber(aggregate.cells)}, non-empty: ${formatNumber(aggregate.nonEmpty)}`,
    aggregate.distinct === null
      ? `- distinct values: more than ${formatNumber(MAX_DISTINCT_TRACKED)} (tracking stopped)`
      : `- distinct values: ${formatNumber(aggregate.distinct)}`,
  ]
  if (aggregate.withheld > 0) {
    // Stated up front, because every number below is a total over the cells
    // that remain: without this the model would report a sum that silently
    // omits the reader's own data.
    lines.push(
      `- withheld: ${formatNumber(aggregate.withheld)} cell(s) hidden from the model; ` +
        'every statistic below covers the remaining cells only',
    )
  }
  if (aggregate.numericCount > 0) {
    lines.push(
      `- numeric cells: ${formatNumber(aggregate.numericCount)}, sum: ${formatNumber(aggregate.sum)}, ` +
        `average: ${formatNumber(aggregate.average ?? 0)}, min: ${formatNumber(aggregate.min ?? 0)}, max: ${formatNumber(aggregate.max ?? 0)}`,
    )
  }
  const shown = aggregate.topValues.slice(0, Math.max(0, topValueCount))
  if (shown.length > 0) {
    lines.push('- top values by frequency:')
    for (const { value, count } of shown) {
      lines.push(
        `  ${value.length > 80 ? `${value.slice(0, 80)}…` : value}: ${formatNumber(count)}`,
      )
    }
  }
  return lines.join('\n')
}
