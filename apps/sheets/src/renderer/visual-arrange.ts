/**
 * Pure arrange math for floating objects: Excel's Align / Distribute on
 * pixel boxes, z-order moves on an ordered id list, and the rotated AABB
 * the anchor must hold for a rotated frame.
 */

export interface ArrangeBox {
  readonly id: string
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

export type AlignMode = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'
export type DistributeAxis = 'horizontal' | 'vertical'
export type ZOrderOp = 'front' | 'back' | 'forward' | 'backward'

export interface BoxMove {
  readonly id: string
  readonly x: number
  readonly y: number
}

/// Excel aligns a multi-selection to its outermost edge (or the center of
/// the selection's bounding box); a single object has nothing to align to.
export function alignBoxes(boxes: readonly ArrangeBox[], mode: AlignMode): BoxMove[] {
  if (boxes.length < 2) return []
  const left = Math.min(...boxes.map((box) => box.x))
  const right = Math.max(...boxes.map((box) => box.x + box.width))
  const top = Math.min(...boxes.map((box) => box.y))
  const bottom = Math.max(...boxes.map((box) => box.y + box.height))
  return boxes
    .map((box) => {
      switch (mode) {
        case 'left':
          return { id: box.id, x: left, y: box.y }
        case 'center':
          return { id: box.id, x: (left + right) / 2 - box.width / 2, y: box.y }
        case 'right':
          return { id: box.id, x: right - box.width, y: box.y }
        case 'top':
          return { id: box.id, x: box.x, y: top }
        case 'middle':
          return { id: box.id, x: box.x, y: (top + bottom) / 2 - box.height / 2 }
        case 'bottom':
          return { id: box.id, x: box.x, y: bottom - box.height }
      }
    })
    .filter((move, index) => move.x !== boxes[index]!.x || move.y !== boxes[index]!.y)
}

/// Even gaps between the objects along one axis; the outermost two stay put.
export function distributeBoxes(boxes: readonly ArrangeBox[], axis: DistributeAxis): BoxMove[] {
  if (boxes.length < 3) return []
  const horizontal = axis === 'horizontal'
  const start = (box: ArrangeBox): number => (horizontal ? box.x : box.y)
  const size = (box: ArrangeBox): number => (horizontal ? box.width : box.height)
  const sorted = [...boxes].sort((left, right) => start(left) - start(right))
  const first = sorted[0]!
  const last = sorted[sorted.length - 1]!
  const span = start(last) + size(last) - start(first)
  const gap = (span - sorted.reduce((total, box) => total + size(box), 0)) / (sorted.length - 1)
  const moves: BoxMove[] = []
  let cursor = start(first) + size(first) + gap
  for (const box of sorted.slice(1, -1)) {
    const next = horizontal
      ? { id: box.id, x: cursor, y: box.y }
      : { id: box.id, x: box.x, y: cursor }
    if (next.x !== box.x || next.y !== box.y) moves.push(next)
    cursor += size(box) + gap
  }
  return moves
}

/// Axis-aligned bounds of a width×height frame rotated by `degrees`.
export function rotatedAabb(
  width: number,
  height: number,
  degrees: number,
): { width: number; height: number } {
  const radians = (degrees * Math.PI) / 180
  const cos = Math.abs(Math.cos(radians))
  const sin = Math.abs(Math.sin(radians))
  return { width: width * cos + height * sin, height: width * sin + height * cos }
}

export const normalizeDegrees = (degrees: number): number => ((degrees % 360) + 360) % 360

/// New document order (back → front) after moving `id`; unchanged when the
/// move is a no-op (already frontmost / backmost).
export function reorderZ(order: readonly string[], id: string, op: ZOrderOp): string[] {
  const at = order.indexOf(id)
  if (at < 0) return [...order]
  const target =
    op === 'front'
      ? order.length - 1
      : op === 'back'
        ? 0
        : op === 'forward'
          ? Math.min(order.length - 1, at + 1)
          : Math.max(0, at - 1)
  if (target === at) return [...order]
  const next = [...order]
  next.splice(at, 1)
  next.splice(target, 0, id)
  return next
}

/// Positions that actually changed between two orders — the only ones the
/// journal needs to carry.
export function changedPositions(
  before: readonly string[],
  after: readonly string[],
): Map<string, number> {
  const changes = new Map<string, number>()
  after.forEach((id, index) => {
    if (before[index] !== id) changes.set(id, index)
  })
  return changes
}
