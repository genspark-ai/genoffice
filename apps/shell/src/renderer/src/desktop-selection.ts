import type { DesktopItem } from './desktop-items'

// ── the desktop's OS-grade selection layer ──────────────
// Pure geometry and set semantics for everything DesktopView's selection
// does: the rubber-band marquee's rectangle hits and modifier merge, the
// icon-click modifiers, the Shift range over the grid's visible-geometry
// order, the arrow keys' geometry walk, and the Ask AI path harvest. No DOM,
// no bridge — every rule here is unit-testable without mounting the view.

/** a surface-local axis-aligned rectangle */
export interface SelectionRect {
  readonly x: number
  readonly y: number
  readonly w: number
  readonly h: number
}

/** one icon's footprint on the surface, at the same coordinates a rect lives in */
export interface SelectionTarget extends SelectionRect {
  readonly id: string
}

/** the marquee rectangle two pointer positions span (either drag direction) */
export function marqueeRect(x0: number, y0: number, x1: number, y1: number): SelectionRect {
  return {
    x: Math.min(x0, x1),
    y: Math.min(y0, y1),
    w: Math.abs(x1 - x0),
    h: Math.abs(y1 - y0),
  }
}

/**
 * Whether two rectangles share positive area — the marquee's intersection
 * rule. Edges that merely touch and a zero-area band claim nothing: the
 * overlap has to be real on both axes before an icon joins the selection.
 */
export function rectsIntersect(a: SelectionRect, b: SelectionRect): boolean {
  const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)
  const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y)
  return w > 0 && h > 0
}

/** the ids whose footprints the rectangle intersects, in the given order */
export function marqueeHits(rect: SelectionRect, targets: readonly SelectionTarget[]): string[] {
  const out: string[] = []
  for (const target of targets) {
    if (rectsIntersect(rect, target)) out.push(target.id)
  }
  return out
}

// ── the modifier merge semantics ────────────────────────

/**
 * How a gesture's hit set combines with the selection it started from:
 * a plain press replaces, Shift unions onto it, Cmd/Ctrl toggles each hit
 * (subtracting the already-selected, adding the unselected — the symmetric
 * difference of the base and the hits).
 */
export type MarqueeMergeMode = 'replace' | 'add' | 'toggle'

/** the selection a marquee currently shows: base ∘ hits under the mode */
export function mergeMarqueeSelection(
  base: ReadonlySet<string>,
  hits: readonly string[],
  mode: MarqueeMergeMode,
): ReadonlySet<string> {
  if (mode === 'replace') return new Set(hits)
  const next = new Set(base)
  for (const hit of hits) {
    if (mode === 'add') next.add(hit)
    else if (next.has(hit)) next.delete(hit)
    else next.add(hit)
  }
  return next
}

/** Cmd/Ctrl+click on one icon: flip that id's membership, keep the rest */
export function toggleSelectionMember(prev: ReadonlySet<string>, id: string): ReadonlySet<string> {
  const next = new Set(prev)
  if (next.has(id)) next.delete(id)
  else next.add(id)
  return next
}

// ── Shift+click: the range over the grid's visible geometry ──

/**
 * The ids between the anchor and the clicked icon in the order the caller
 * passes — the view hands it the grid's visible-geometry order below, so the
 * span is what the eye sees. A missing anchor (nothing clicked before, or it
 * left the surface) falls back to the click alone, the way an OS list treats
 * a Shift-click without an anchor.
 */
export function rangeSelection(
  order: readonly string[],
  anchor: string | null,
  target: string,
): ReadonlySet<string> {
  const to = order.indexOf(target)
  if (to < 0) return new Set([target])
  const anchorAt = anchor !== null && order.includes(anchor) ? order.indexOf(anchor) : to
  const [from, till] = anchorAt <= to ? [anchorAt, to] : [to, anchorAt]
  return new Set(order.slice(from, till + 1))
}

// ── the arrow keys' geometry walk ───────────────────────

/** where an icon sits on the grid, in whole cells */
export interface GridCell {
  readonly id: string
  readonly col: number
  readonly row: number
}

export type FocusDirection = 'left' | 'right' | 'up' | 'down'

/**
 * The icon the focus moves to. Geometry, not list order: a left/right step
 * picks the nearest icon in a strictly different column (ties by nearest
 * row), an up/down step the nearest icon in a strictly different row (ties
 * by nearest column) — so saved positions, not the provider order, decide
 * what "next to" means. With nothing focused, down/right enter from the
 * top-left-most icon and up/left from the bottom-right-most one.
 */
export function moveGridFocus(
  cells: readonly GridCell[],
  current: string | null,
  dir: FocusDirection,
): string | null {
  if (cells.length === 0) return null
  const me = current !== null ? cells.find((cell) => cell.id === current) : undefined
  if (me === undefined) {
    // nothing focused: enter from the reading corner the direction comes
    // from — down/right from the top-left-most icon, up/left from the
    // bottom-right-most one (the far cell of the direction's own axis)
    const horizontal = dir === 'left' || dir === 'right'
    const forward = dir === 'right' || dir === 'down'
    let corner = cells[0]!
    for (const cell of cells) {
      const better = horizontal
        ? forward
          ? cell.col < corner.col || (cell.col === corner.col && cell.row < corner.row)
          : cell.col > corner.col || (cell.col === corner.col && cell.row > corner.row)
        : forward
          ? cell.row < corner.row || (cell.row === corner.row && cell.col < corner.col)
          : cell.row > corner.row || (cell.row === corner.row && cell.col > corner.col)
      if (better) corner = cell
    }
    return corner.id
  }
  const horizontal = dir === 'left' || dir === 'right'
  const forward = dir === 'right' || dir === 'down'
  let best: GridCell | null = null
  let bestPrimary = Number.POSITIVE_INFINITY
  let bestSecondary = Number.POSITIVE_INFINITY
  for (const cell of cells) {
    if (cell.id === me.id) continue
    const primary = horizontal ? cell.col - me.col : cell.row - me.row
    if (forward ? primary <= 0 : primary >= 0) continue
    const primaryDistance = Math.abs(primary)
    const secondaryDistance = Math.abs(horizontal ? cell.row - me.row : cell.col - me.col)
    if (
      primaryDistance < bestPrimary ||
      (primaryDistance === bestPrimary && secondaryDistance < bestSecondary)
    ) {
      best = cell
      bestPrimary = primaryDistance
      bestSecondary = secondaryDistance
    }
  }
  return best === null ? null : best.id
}

/**
 * The grid's visible-geometry order, derived row-first from the resolved
 * (col, row) cells — rows top to bottom, columns left to right within a row.
 * The Shift range walks this order so a span tracks what the eye sees on the
 * surface: the saved positions decide the geometry (exactly like the arrow
 * keys' walk), not the provider order, which a Custom placement can scatter
 * across the grid far from its list neighbors.
 */
export function geometryOrderOf(cells: readonly GridCell[]): string[] {
  return [...cells].sort((a, b) => a.row - b.row || a.col - b.col).map((cell) => cell.id)
}

// ── the Ask AI corpus harvest ───────────────────────────

/**
 * The real paths behind a selection, for the Ask AI handoff: files and
 * folders carry a path on disk; app tiles are furniture and skip. Order is
 * the caller's (the view passes the grid's flow order so the corpus list is
 * deterministic), and duplicates cannot survive a one-path-per-icon model
 * but are dropped anyway.
 */
export function askPathsOf(items: readonly DesktopItem[]): string[] {
  const out: string[] = []
  for (const item of items) {
    if (item.kind !== 'file' && item.kind !== 'folder') continue
    if (item.path === null || out.includes(item.path)) continue
    out.push(item.path)
  }
  return out
}

/** the empty selection, shared so toggles never allocate a fresh singleton */
export const EMPTY_SELECTION: ReadonlySet<string> = new Set()
