import type { TabStop } from '@genoffice/docx-engine'
import { parseTwips } from './lengths'

export interface TabSpec {
  /** twips, or a length string ("1in", "6.5in") */
  pos: number | string
  align?: 'left' | 'center' | 'right' | 'decimal' | 'bar'
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'middleDot' | 'heavy'
  /** cancel the style-inherited stop at this position instead of defining one */
  clear?: boolean
}

const TAB_ALIGNS = ['left', 'center', 'right', 'decimal', 'bar'] as const
const TAB_LEADERS = ['none', 'dot', 'hyphen', 'underscore', 'middleDot', 'heavy'] as const

/** shared by setParagraphFormat and define_style → paragraph.tabs */
export function validateTabs(tabs: unknown, where: string): string | null {
  if (tabs === null) return null
  if (!Array.isArray(tabs))
    return `${where}: tabs must be an array of { pos, align?, leader?, clear? } or null`
  for (const t of tabs as Record<string, unknown>[]) {
    if (!t || typeof t !== 'object') return `${where}: each tab must be an object`
    if (parseTwips(t.pos) === undefined)
      return `${where}: tab pos must be twips or a length like "6.5in"`
    if (t.align !== undefined && !(TAB_ALIGNS as readonly unknown[]).includes(t.align))
      return `${where}: tab align must be ${TAB_ALIGNS.join(' / ')}`
    if (t.leader !== undefined && !(TAB_LEADERS as readonly unknown[]).includes(t.leader))
      return `${where}: tab leader must be ${TAB_LEADERS.join(' / ')}`
    if (t.clear !== undefined && typeof t.clear !== 'boolean')
      return `${where}: tab clear must be a boolean`
  }
  return null
}

/** validated TabSpec list → engine stops (w:tabs entries); null/[] → [] */
export function tabSpecsToStops(tabs: TabSpec[] | null): TabStop[] {
  return (tabs ?? []).map((t) => {
    const stop: TabStop = { pos: parseTwips(t.pos)!, val: t.clear ? 'clear' : (t.align ?? 'left') }
    if (t.leader && t.leader !== 'none' && !t.clear) stop.leader = t.leader
    return stop
  })
}
