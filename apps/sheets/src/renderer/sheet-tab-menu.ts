/// Pure index math behind the Excel-style sheet-tab context menu: the Unhide
/// list, the Move or Copy target list and the reorder plan that reproduces a
/// final tab order through Univer's one-sheet-at-a-time order mutation.

export interface SheetTabEntry {
  readonly id: string
  readonly name: string
  readonly hidden: boolean
  /// Excel state="veryHidden": never offered in Unhide (only VBA may show it).
  readonly veryHidden: boolean
}

/// Excel's Unhide dialog lists hidden sheets but not veryHidden ones.
export function unhideCandidates<T extends SheetTabEntry>(sheets: readonly T[]): T[] {
  return sheets.filter((sheet) => sheet.hidden && !sheet.veryHidden)
}

/// "Before sheet" choices: every tab the user can see or unhide, in tab order.
export function moveTargets<T extends SheetTabEntry>(sheets: readonly T[]): T[] {
  return sheets.filter((sheet) => !sheet.veryHidden)
}

/// Visible sheets left after hiding or deleting `ids` (Excel keeps >= 1).
export function visibleAfterRemoving(
  sheets: readonly SheetTabEntry[],
  ids: ReadonlySet<string>,
): number {
  return sheets.filter((sheet) => !sheet.hidden && !ids.has(sheet.id)).length
}

/// Excel inserts a new worksheet in front of the active one.
export function insertIndex(order: readonly string[], activeId: string): number {
  const index = order.indexOf(activeId)
  return index < 0 ? order.length : index
}

/// The tab order after moving `moving` (kept in their current relative order)
/// in front of `beforeId`, or to the end when `beforeId` is null. A target
/// that is itself being moved resolves to the first unmoved sheet after it.
export function finalOrder(
  order: readonly string[],
  moving: ReadonlySet<string>,
  beforeId: string | null,
): string[] {
  const moved = order.filter((id) => moving.has(id))
  const rest = order.filter((id) => !moving.has(id))
  if (beforeId === null) return [...rest, ...moved]
  let anchor = order.indexOf(beforeId)
  while (anchor >= 0 && anchor < order.length && moving.has(order[anchor]!)) anchor += 1
  const anchorId = anchor >= 0 && anchor < order.length ? order[anchor]! : null
  const at = anchorId === null ? rest.length : rest.indexOf(anchorId)
  return [...rest.slice(0, at), ...moved, ...rest.slice(at)]
}

export interface ReorderStep {
  readonly id: string
  /// Index in the order with `id` already spliced out (Univer's `toOrder`).
  readonly to: number
}

/// Steps that turn `order` into `target` by relocating only the `moving`
/// sheets, each a single splice-out/splice-in exactly like
/// `sheet.mutation.set-worksheet-order`. Unmoved sheets and already placed
/// ones are "settled"; each step drops the sheet right after the last
/// settled sheet that precedes it in `target`, so later steps never
/// disturb earlier ones.
export function planReorder(
  order: readonly string[],
  target: readonly string[],
  moving: ReadonlySet<string>,
): ReorderStep[] {
  const steps: ReorderStep[] = []
  const current = [...order]
  const rank = new Map(target.map((id, index) => [id, index]))
  const settled = new Set(order.filter((id) => !moving.has(id)))
  for (const id of target) {
    if (!moving.has(id)) continue
    const from = current.indexOf(id)
    if (from < 0) continue
    current.splice(from, 1)
    const mine = rank.get(id) ?? 0
    let to = 0
    current.forEach((other, index) => {
      if (settled.has(other) && (rank.get(other) ?? 0) < mine) to = index + 1
    })
    current.splice(to, 0, id)
    settled.add(id)
    if (from !== to) steps.push({ id, to })
  }
  return steps
}

/// Applies `steps` to `order` with the mutation's semantics (test helper and
/// the simulation the copy path uses to predict where a clone will land).
export function applyReorder(order: readonly string[], steps: readonly ReorderStep[]): string[] {
  const current = [...order]
  for (const step of steps) {
    const from = current.indexOf(step.id)
    if (from < 0) continue
    current.splice(from, 1)
    current.splice(step.to, 0, step.id)
  }
  return current
}

/// Ids between (and including) `fromId` and `toId` in tab order, skipping
/// hidden sheets — Excel's Shift+click tab range.
export function tabRange(sheets: readonly SheetTabEntry[], fromId: string, toId: string): string[] {
  const order = sheets.map((sheet) => sheet.id)
  const a = order.indexOf(fromId)
  const b = order.indexOf(toId)
  if (a < 0 || b < 0) return [toId]
  const [start, end] = a <= b ? [a, b] : [b, a]
  return sheets
    .slice(start, end + 1)
    .filter((sheet) => !sheet.hidden)
    .map((sheet) => sheet.id)
}
