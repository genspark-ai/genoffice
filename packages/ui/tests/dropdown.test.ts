import { describe, expect, it } from 'vitest'
import { filterOptions, nextEnabledIndex, reconcileActiveIndex } from '../src/dropdown'

describe('Dropdown keyboard navigation', () => {
  const options = [{ disabled: true }, {}, { disabled: true }, {}]

  it('skips disabled options in both directions', () => {
    expect(nextEnabledIndex(options, 0, 1)).toBe(1)
    expect(nextEnabledIndex(options, 2, -1)).toBe(1)
  })

  it('finds the enabled Home and End targets and reports none when all are disabled', () => {
    expect(nextEnabledIndex(options, 0, 1)).toBe(1)
    expect(nextEnabledIndex(options, options.length - 1, -1)).toBe(3)
    expect(nextEnabledIndex([{ disabled: true }], 0, 1)).toBe(-1)
  })

  it('reconciles active indexes when options shrink or become disabled', () => {
    const choices = [
      { value: 'auto', disabled: true },
      { value: 'left' },
      { value: 'right', disabled: true },
      { value: 'center' },
    ]
    expect(reconcileActiveIndex(choices, 1, 'left')).toBe(1)
    expect(reconcileActiveIndex(choices, 0, 'auto')).toBe(1)
    expect(reconcileActiveIndex(choices, 2, 'right')).toBe(3)
    expect(reconcileActiveIndex(choices, 8, 'center')).toBe(3)
    expect(reconcileActiveIndex([{ value: 'auto', disabled: true }], 0, 'auto')).toBe(-1)
  })
})

describe('filterOptions', () => {
  const options = [
    { value: 'anthropic/claude-opus-5.5', label: 'anthropic/claude-opus-5.5' },
    { value: 'openai/gpt-6-astra', label: 'openai/gpt-6-astra' },
    { value: 'moonshotai/kimi-k3', label: 'moonshotai/kimi-k3' },
  ]

  it('matches case-insensitively on any substring', () => {
    expect(filterOptions(options, 'OPUS').map((o) => o.value)).toEqual([
      'anthropic/claude-opus-5.5',
    ])
  })

  it('requires every whitespace-separated term', () => {
    expect(filterOptions(options, 'openai astra')).toHaveLength(1)
    expect(filterOptions(options, 'openai kimi')).toHaveLength(0)
  })
})
