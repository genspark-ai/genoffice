/**
 * Shift+drag axis lock (PowerPoint): the move follows whichever axis has the larger displacement
 * from the drag origin. Callers re-run it on every drag move, so the lock re-evaluates as the
 * pointer travels; an exact tie keeps the move horizontal.
 */
export function constrainAxis(dx: number, dy: number): { dx: number; dy: number } {
  return Math.abs(dy) > Math.abs(dx) ? { dx: 0, dy } : { dx, dy: 0 }
}
