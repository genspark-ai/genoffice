import { describe, expect, it } from 'vitest'
import { liveSections, sameSectionInfos } from '../src/renderer/pagination-sections'

type SectionInfo = Parameters<typeof liveSections>[0][number]
type BlockBox = Parameters<typeof liveSections>[1][number]

const section = (first: number, last: number): SectionInfo =>
  ({
    firstBlockIndex: first,
    lastBlockIndex: last,
    startType: 'nextPage',
  }) as unknown as SectionInfo
const present = (...docxIndexes: number[]): BlockBox[] =>
  docxIndexes.map((docxIndex) => ({ docxIndex }) as unknown as BlockBox)

describe('sameSectionInfos', () => {
  const sections = [section(0, 9), section(10, 19), section(20, 29)]

  it('treats the merged list a streaming open recomputes per chunk as unchanged', () => {
    // the second section's break block has not streamed in yet: it merges into the third
    const a = liveSections(sections, present(9, 29))
    const b = liveSections(sections, present(9, 29))
    expect(a).not.toBe(b)
    expect(a).not.toBe(sections)
    expect(sameSectionInfos(a, b)).toBe(true)
  })

  it('reports a change once another break block lands', () => {
    const merged = liveSections(sections, present(9, 29))
    const complete = liveSections(sections, present(9, 19, 29))
    expect(complete).toBe(sections)
    expect(sameSectionInfos(merged, complete)).toBe(false)
  })

  it('is identity for the same array and false for a different length', () => {
    expect(sameSectionInfos(sections, sections)).toBe(true)
    expect(sameSectionInfos(sections, sections.slice(1))).toBe(false)
  })
})
