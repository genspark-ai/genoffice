/**
 * Univer's border extension asks, per left/right border segment, whether an
 * overflowing text run covers that edge — and answers by walking EVERY row of
 * the overflow cache to find the one row it cares about. On a bordered sheet
 * whose text overflows (long labels in default-width columns) that scan ran
 * for thousands of segments a frame and was the single largest render cost
 * while scrolling. Look the row up directly instead.
 */
import { Border } from '@univerjs/engine-render'

interface OverflowSpan {
  startColumn: number
  endColumn: number
}

interface OverflowCache {
  getRow(row: number): Record<string, OverflowSpan> | undefined
}

export function overflowCoversBorder(
  overflowCache: OverflowCache | null | undefined,
  type: string,
  borderRow: number,
  borderColumn: number,
): boolean {
  if (type !== 'l' && type !== 'r') return false
  const spans = overflowCache?.getRow(borderRow)
  if (!spans) return false
  for (const key in spans) {
    const span = spans[key]
    if (!span) continue
    if (type === 'l' && borderColumn > span.startColumn && borderColumn <= span.endColumn) {
      return true
    }
    if (type === 'r' && borderColumn >= span.startColumn && borderColumn < span.endColumn) {
      return true
    }
  }
  return false
}

let installed = false

export function installBorderOverflowExclusionFix(): void {
  if (installed) return
  installed = true
  const proto = Border.prototype as unknown as {
    _getOverflowExclusion(
      overflowCache: OverflowCache | null | undefined,
      type: string,
      borderRow: number,
      borderColumn: number,
    ): boolean
  }
  if (typeof proto._getOverflowExclusion !== 'function') return
  proto._getOverflowExclusion = overflowCoversBorder
}
