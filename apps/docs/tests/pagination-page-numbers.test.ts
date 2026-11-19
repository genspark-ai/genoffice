import { describe, expect, it } from 'vitest'
import type { SectionInfo } from '@genoffice/docx-engine'
import { pageNoAfterOf, pageNumbers } from '../src/renderer/pagination'
import type { PageSlice, SectionGeom } from '../src/renderer/pagination-types'

// deterministic LCG so a failing seed can be replayed
const rng = (seed: number) => () => {
  seed = (seed * 1664525 + 1013904223) >>> 0
  return seed / 2 ** 32
}

const START_TYPES: Array<SectionInfo['startType'] | undefined> = [
  undefined,
  'nextPage',
  'continuous',
  'evenPage',
  'oddPage',
  'nextColumn',
]

function randomCase(rand: () => number): { sections: SectionInfo[]; slices: PageSlice[] } {
  const sectionCount = 1 + Math.floor(rand() * 5)
  const sections = Array.from({ length: sectionCount }, (_, i) => {
    const sec = {
      startType: START_TYPES[Math.floor(rand() * START_TYPES.length)],
      ...(rand() < 0.4 ? { pageNumberStart: Math.floor(rand() * 7) } : {}),
    }
    return (i === 0 && rand() < 0.5
      ? { ...sec, startType: undefined }
      : sec) as unknown as SectionInfo
  })
  const slices: PageSlice[] = []
  let y = 0
  for (let s = 0; s < sectionCount; s++) {
    const pages = 1 + Math.floor(rand() * 4)
    for (let p = 0; p < pages; p++) {
      slices.push({ start: y, end: y + 100, section: s })
      y += 100
    }
  }
  return { sections, slices }
}

/** the slicer's incremental page-number walk (pagination-slices.ts) */
function walkPageNumbers(slices: PageSlice[], sections: SectionInfo[]): number[] {
  const geomOf = (s: number): SectionGeom => ({
    contentHeight: 1000,
    forceBreak: false,
    startType: sections[s]?.startType,
    ...(sections[s]?.pageNumberStart !== undefined
      ? { pageNumberStart: sections[s].pageNumberStart }
      : {}),
  })
  let prevNo = 0
  let prevSection = -1
  return slices.map((s) => {
    const no = pageNoAfterOf(geomOf, prevNo, prevSection, s.section)
    prevNo = no
    prevSection = s.section
    return no
  })
}

describe('pageNumbers == pageNoAfterOf walk', () => {
  it('agrees on restarts, parity skips and continuation (1000 random section layouts)', () => {
    const rand = rng(20261030)
    for (let n = 0; n < 1000; n++) {
      const { sections, slices } = randomCase(rand)
      expect(walkPageNumbers(slices, sections), JSON.stringify(sections)).toEqual(
        pageNumbers(slices, sections),
      )
    }
  })

  it('pins the parity skip: two evenPage sections after page 1 show 2 and 4', () => {
    const sections = [
      {},
      { startType: 'evenPage' },
      { startType: 'evenPage' },
    ] as unknown as SectionInfo[]
    const slices: PageSlice[] = [
      { start: 0, end: 1, section: 0 },
      { start: 1, end: 2, section: 1 },
      { start: 2, end: 3, section: 2 },
    ]
    expect(pageNumbers(slices, sections)).toEqual([1, 2, 4])
    expect(walkPageNumbers(slices, sections)).toEqual([1, 2, 4])
  })
})
