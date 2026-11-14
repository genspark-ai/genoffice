import { describe, expect, it } from 'vitest'

import {
  applyFlashFillTemplate,
  flashFillFirstSource,
  inferFlashFillTemplate,
} from '@genoffice/xlsx-gateway/domain/flash-fill'

describe('inferFlashFillTemplate', () => {
  it('learns a join with literal glue from one example', () => {
    const template = inferFlashFillTemplate([{ source: ['John', 'Sales'], output: 'John·Sales' }])
    expect(template).not.toBeNull()
    expect(applyFlashFillTemplate(template!, ['Mary', 'Marketing'])).toBe('Mary·Marketing')
  })

  it('learns prefixes and suffixes', () => {
    const template = inferFlashFillTemplate([
      { source: ['A101', 'Alice'], output: 'Alice (A101)' },
      { source: ['B202', 'Bob'], output: 'Bob (B202)' },
    ])
    expect(template).not.toBeNull()
    expect(applyFlashFillTemplate(template!, ['C303', 'Carol'])).toBe('Carol (C303)')
  })

  it('prefers the longest field when fields overlap', () => {
    const template = inferFlashFillTemplate([
      { source: ['York', 'Yorkshire'], output: 'Yorkshire-York' },
    ])
    expect(template).not.toBeNull()
    expect(applyFlashFillTemplate(template!, ['Kent', 'Kentwood'])).toBe('Kentwood-Kent')
  })

  it('rejects examples that contradict each other', () => {
    expect(
      inferFlashFillTemplate([
        { source: ['a1', 'b1'], output: 'a1-b1' },
        { source: ['a2', 'b2'], output: 'b2/a2' },
      ]),
    ).toBeNull()
  })

  it('rejects outputs with no recognizable field', () => {
    expect(inferFlashFillTemplate([{ source: ['abc', 'def'], output: 'xyz' }])).toBeNull()
  })

  it('trims source whitespace when matching and applying', () => {
    const template = inferFlashFillTemplate([
      { source: [' John ', ' 13800000000 '], output: 'John: 13800000000' },
    ])
    expect(template).not.toBeNull()
    expect(applyFlashFillTemplate(template!, ['Mary', '13911111111'])).toBe('Mary: 13911111111')
  })
})

describe('flashFillFirstSource', () => {
  const filled =
    (...columns: number[]) =>
    (column: number) =>
      columns.includes(column)

  it('reaches the first column of a wide list, not only the six nearest', () => {
    expect(flashFillFirstSource(9, filled(0, 1, 2, 3, 4, 5, 6, 7, 8))).toBe(0)
  })

  it('keeps the six nearest columns, across an empty one inside the list', () => {
    expect(flashFillFirstSource(3, filled(0, 2))).toBe(0)
    expect(flashFillFirstSource(6, filled(0, 1, 3, 4, 5))).toBe(0)
    expect(flashFillFirstSource(4, filled(0, 1))).toBe(0)
  })

  it('stops a walk past the six nearest at an empty column', () => {
    const list = Array.from({ length: 20 }, (_, column) => column).filter((column) => column !== 11)
    expect(flashFillFirstSource(20, filled(...list))).toBe(12)
  })

  it('walks no further than 64 columns', () => {
    expect(flashFillFirstSource(100, () => true)).toBe(36)
  })
})
