import type { TabStop } from '@genoffice/docx-engine'
import { directTabStops } from './Ruler'

/** the paragraph's tab stops attribute, split into the stops a user sets and the ones a w:ptab stands for */
export function parseTabStops(raw: unknown): { stops: TabStop[]; relStops: TabStop[] } {
  if (typeof raw !== 'string') return { stops: [], relStops: [] }
  try {
    const parsed: unknown = JSON.parse(raw)
    if (!Array.isArray(parsed)) return { stops: [], relStops: [] }
    const all = parsed as TabStop[]
    return { stops: all.filter((s) => !s.rel), relStops: all.filter((s) => s.rel) }
  } catch {
    return { stops: [], relStops: [] }
  }
}

/** the attribute for an edited set of stops: inherited ones removed are cleared; w:ptab stops and
 *  the clears already there are kept, or a style stop removed earlier would come back */
export function tabStopsAttr(raw: unknown, edited: TabStop[]): string | null {
  const { stops, relStops } = parseTabStops(raw)
  const keptClears = stops.filter((s) => s.val === 'clear' && !edited.some((e) => e.pos === s.pos))
  const direct = directTabStops(stops, [...edited, ...keptClears])
  return direct.length > 0 || relStops.length > 0 ? JSON.stringify([...direct, ...relStops]) : null
}

/** the stops with one set: a stop already at that position is replaced */
export function withTabStop(stops: TabStop[], stop: TabStop): TabStop[] {
  return [...stops.filter((s) => s.pos !== stop.pos), stop].sort((a, b) => a.pos - b.pos)
}
