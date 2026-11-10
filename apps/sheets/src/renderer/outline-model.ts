/**
 * Pure row/column outline model (Excel "Group"): the per-line level map
 * becomes nested groups, and the gutter buttons translate into hide/show
 * spans plus the summary-line collapsed flags the file stores.
 */

export interface OutlineEntry {
  level: number
  collapsed: boolean
  /// hidden="1" straight from the file, for lines the grid has not loaded yet.
  fileHidden?: boolean
}

export interface OutlineGroup {
  readonly level: number
  readonly start: number
  readonly end: number
  /// Summary line the [-]/[+] button sits on; -1 when it would fall before line 0.
  readonly summary: number
  /// Every detail line is hidden (what the button shows).
  readonly collapsed: boolean
  /// The summary line's own collapsed flag — the user folded this group
  /// itself, not just an outer one around it.
  readonly folded: boolean
}

export interface IndexSpan {
  readonly start: number
  readonly end: number
}

export interface SheetOutlineState {
  readonly rows: Map<number, OutlineEntry>
  readonly cols: Map<number, OutlineEntry>
  summaryBelow: boolean
  summaryRight: boolean
  /// Bumped on every change so the gutter can repaint without diffing maps.
  version: number
  /// File rows the sidecar index already delivered; a later poll must not
  /// re-seed a row the user has since cleared.
  readonly seededRows: Set<number>
}

export const MAX_OUTLINE_LEVEL = 7
export const OUTLINE_LEVEL_STEP = 13
export const OUTLINE_GUTTER_PAD = 3

export function outlineMaxLevel(entries: ReadonlyMap<number, OutlineEntry>): number {
  let max = 0
  for (const entry of entries.values()) if (entry.level > max) max = entry.level
  return Math.min(MAX_OUTLINE_LEVEL, max)
}

/// Gutter thickness for a given depth: one column per level button (1..max+1).
export function outlineGutterThickness(maxLevel: number): number {
  return maxLevel === 0 ? 0 : (maxLevel + 1) * OUTLINE_LEVEL_STEP + OUTLINE_GUTTER_PAD * 2
}

/// Center of the gutter column for 1-based level button `level`.
export function outlineColumnCenter(level: number): number {
  return OUTLINE_GUTTER_PAD + (level - 1) * OUTLINE_LEVEL_STEP + OUTLINE_LEVEL_STEP / 2
}

function sortedIndices(entries: ReadonlyMap<number, OutlineEntry>): number[] {
  const indices: number[] = []
  for (const [index, entry] of entries) if (entry.level > 0) indices.push(index)
  return indices.sort((left, right) => left - right)
}

/// Excel's grouping rule: at each level L, every maximal run of consecutive
/// lines whose level is >= L is one group of level L.
export function outlineGroups(
  entries: ReadonlyMap<number, OutlineEntry>,
  summaryAfter: boolean,
  isHidden: (index: number) => boolean,
): OutlineGroup[] {
  const indices = sortedIndices(entries)
  const maxLevel = outlineMaxLevel(entries)
  const groups: OutlineGroup[] = []
  for (let level = 1; level <= maxLevel; level += 1) {
    let runStart = -1
    let runEnd = -1
    const close = (): void => {
      if (runStart < 0) return
      let hidden = true
      for (let index = runStart; index <= runEnd && hidden; index += 1) {
        if (!isHidden(index)) hidden = false
      }
      const summary = summaryAfter ? runEnd + 1 : runStart - 1
      groups.push({
        level,
        start: runStart,
        end: runEnd,
        summary,
        collapsed: hidden,
        folded: entries.get(summary)?.collapsed === true,
      })
      runStart = -1
    }
    for (const index of indices) {
      if ((entries.get(index)?.level ?? 0) < level) continue
      if (runStart >= 0 && index === runEnd + 1) {
        runEnd = index
        continue
      }
      close()
      runStart = index
      runEnd = index
    }
    close()
  }
  return groups
}

/// Lines with a level >= `level`, as runs — the detail hidden by level button `level`.
export function outlineSpansAtLeast(
  entries: ReadonlyMap<number, OutlineEntry>,
  level: number,
): IndexSpan[] {
  const spans: IndexSpan[] = []
  let current: { start: number; end: number } | null = null
  for (const index of sortedIndices(entries)) {
    if ((entries.get(index)?.level ?? 0) < level) continue
    if (current && index === current.end + 1) {
      current.end = index
      continue
    }
    if (current) spans.push(current)
    current = { start: index, end: index }
  }
  if (current) spans.push(current)
  return spans
}

export interface LevelButtonPlan {
  readonly hide: IndexSpan[]
  readonly show: IndexSpan[]
  /// Summary lines whose collapsed flag changes (groups at >= level fold).
  readonly summaries: { index: number; collapsed: boolean }[]
}

/// Level button N: detail at levels >= N folds, shallower detail unfolds.
export function levelButtonPlan(
  entries: ReadonlyMap<number, OutlineEntry>,
  groups: readonly OutlineGroup[],
  level: number,
): LevelButtonPlan {
  const hide = outlineSpansAtLeast(entries, level)
  const show: IndexSpan[] = []
  for (const span of outlineSpansAtLeast(entries, 1)) {
    let cursor = span.start
    for (const hidden of hide) {
      if (hidden.end < span.start || hidden.start > span.end) continue
      if (hidden.start > cursor) show.push({ start: cursor, end: hidden.start - 1 })
      cursor = hidden.end + 1
    }
    if (cursor <= span.end) show.push({ start: cursor, end: span.end })
  }
  const summaries = new Map<number, boolean>()
  for (const group of groups) {
    if (group.summary < 0) continue
    const collapsed = group.level >= level
    // The shallowest group on a shared summary line decides its flag.
    if (!summaries.has(group.summary) || !collapsed) summaries.set(group.summary, collapsed)
  }
  return {
    hide,
    show,
    summaries: [...summaries].map(([index, collapsed]) => ({ index, collapsed })),
  }
}

export interface GroupTogglePlan {
  readonly hide: IndexSpan[]
  readonly show: IndexSpan[]
  readonly summaries: { index: number; collapsed: boolean }[]
}

/// [-] hides the whole span; [+] shows it except nested groups the user
/// folded themselves (their summary flag), not ones hidden only by the outer fold.
export function groupTogglePlan(
  groups: readonly OutlineGroup[],
  group: OutlineGroup,
  collapsed: boolean,
): GroupTogglePlan {
  const summaries = group.summary < 0 ? [] : [{ index: group.summary, collapsed }]
  if (collapsed) return { hide: [{ start: group.start, end: group.end }], show: [], summaries }
  const nested = groups.filter(
    (other) =>
      other.level > group.level &&
      other.start >= group.start &&
      other.end <= group.end &&
      // A nested group on the same summary line shares the flag we are about
      // to clear, so it cannot be told apart from an outer-only fold: unfold it.
      other.summary !== group.summary &&
      other.folded,
  )
  return {
    hide: nested.map((other) => ({ start: other.start, end: other.end })),
    show: [{ start: group.start, end: group.end }],
    summaries,
  }
}

/// Re-key the level map after rows/columns are inserted (count > 0) or
/// removed (count < 0) at `index`. Lines inserted inside a group take the
/// level of the line above them (Excel), so the group stays one run;
/// returns that level (0 when nothing was filled).
export function shiftOutlineEntries(
  entries: Map<number, OutlineEntry>,
  index: number,
  count: number,
): number {
  if (count === 0) return 0
  const above = index > 0 ? (entries.get(index - 1)?.level ?? 0) : 0
  const below = entries.get(index)?.level ?? 0
  const moved: [number, OutlineEntry][] = []
  for (const [at, entry] of entries) {
    if (at < index) continue
    entries.delete(at)
    if (count < 0 && at < index - count) continue
    moved.push([at + count, entry])
  }
  for (const [at, entry] of moved) entries.set(at, entry)
  const fill = count > 0 && above > 0 && below > 0 ? Math.min(above, below) : 0
  for (let at = index; fill > 0 && at < index + count; at += 1) {
    entries.set(at, { level: fill, collapsed: false })
  }
  return fill
}

/// Re-key after rows [index, index+count) move before pre-move row `before`:
/// the two adjacent blocks swap places.
export function moveOutlineEntries(
  entries: Map<number, OutlineEntry>,
  index: number,
  count: number,
  before: number,
): void {
  if (count <= 0 || (before >= index && before <= index + count)) return
  const first =
    before > index ? { start: index, end: index + count - 1 } : { start: before, end: index - 1 }
  const second =
    before > index
      ? { start: index + count, end: before - 1 }
      : { start: index, end: index + count - 1 }
  const moved: [number, OutlineEntry][] = []
  for (const [at, entry] of entries) {
    if (at >= first.start && at <= first.end) {
      entries.delete(at)
      moved.push([at + (second.end - second.start + 1), entry])
    } else if (at >= second.start && at <= second.end) {
      entries.delete(at)
      moved.push([at - (first.end - first.start + 1), entry])
    }
  }
  for (const [at, entry] of moved) entries.set(at, entry)
}

/// Group/Ungroup over a selection: every run of equal levels shifts by ±1.
export function groupLevelOps(
  entries: ReadonlyMap<number, OutlineEntry>,
  start: number,
  end: number,
  delta: 1 | -1,
): { start: number; end: number; level: number }[] {
  const ops: { start: number; end: number; level: number }[] = []
  let runStart = start
  let runLevel = entries.get(start)?.level ?? 0
  const closeRun = (runEnd: number): void => {
    const level = Math.min(MAX_OUTLINE_LEVEL, Math.max(0, runLevel + delta))
    if (level !== runLevel) ops.push({ start: runStart, end: runEnd, level })
  }
  for (let index = start + 1; index <= end; index += 1) {
    const level = entries.get(index)?.level ?? 0
    if (level !== runLevel) {
      closeRun(index - 1)
      runStart = index
      runLevel = level
    }
  }
  closeRun(end)
  return ops
}
