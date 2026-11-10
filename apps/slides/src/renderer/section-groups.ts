/**
 * Thumbnail-rail section grouping shared by the rail, the section context menu
 * and the section actions. Mirrors the engine's positional normalization: section
 * i covers [start_i, start_{i+1}); slides before the first section form a lead
 * group with id null and no name (PowerPoint's "Default Section", which the file
 * never stores).
 */
import type { SectionInfo } from '../shared/ipc'

export interface SectionGroup {
  id: string | null
  name: string
  start: number
  end: number
}

export function groupSections(sections: SectionInfo[], total: number): SectionGroup[] | null {
  if (!sections.length || !total) return null
  const starts = new Array<number>(sections.length)
  let nextStart = total
  for (let i = sections.length - 1; i >= 0; i--) {
    const own = sections[i]!.slideIndices.length
      ? Math.min(...sections[i]!.slideIndices)
      : nextStart
    starts[i] = Math.min(own, nextStart)
    nextStart = starts[i]!
  }
  const groups: SectionGroup[] = []
  if (starts[0]! > 0) groups.push({ id: null, name: '', start: 0, end: starts[0]! })
  sections.forEach((s, i) => {
    groups.push({
      id: s.id,
      name: s.name,
      start: starts[i]!,
      end: i + 1 < sections.length ? starts[i + 1]! : total,
    })
  })
  return groups
}

export function setAllCollapsed(sections: SectionInfo[], collapsed: boolean): Set<string> {
  return new Set(collapsed ? sections.map((s) => s.id) : [])
}

/** Current slide after removing [group.start, group.end): the hole is now filled by the slide that followed it. */
export function currentAfterRemoval(
  current: number,
  group: SectionGroup,
  remaining: number,
): number {
  if (current >= group.end) return current - (group.end - group.start)
  if (current >= group.start) return Math.min(group.start, remaining - 1)
  return current
}

export function indexRange(start: number, end: number): number[] {
  const out: number[] = []
  for (let i = start; i < end; i++) out.push(i)
  return out
}
