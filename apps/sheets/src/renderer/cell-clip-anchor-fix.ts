/**
 * Excel clips overflowing cell text at the BOTTOM: when a multi-line text
 * block is taller than its cell, the FIRST line stays visible regardless of
 * the cell's vertical alignment (verified on production files — a
 * bottom-aligned wrapped cell at a fixed one-line height shows line 1 in
 * Excel). Univer anchors the block by its alignment, so bottom/middle cells
 * showed the LAST line(s). Two render paths need the fix:
 *
 * - Rich/multi-line documents (cell `p` data): Documents._verticalHandler
 *   returns a negative page offset — clamp it to the cell top (multi-line
 *   blocks only; single lines keep their native alignment on both paths).
 * - Plain strings (wrapped by style): the static Text.drawWith fast path
 *   bottom-anchors its line block — flip an overflowing multi-line block to
 *   TOP so trailing lines are the ones dropped.
 */
import { VerticalAlign } from '@univerjs/core'
import { DocSimpleSkeleton, Documents, LineType, Text } from '@univerjs/engine-render'

export function excelClipOffsetTop(offsetTop: number, lineCount: number): number {
  return offsetTop < 0 && lineCount > 1 ? 0 : offsetTop
}

interface SkeletonDataLike {
  pages?: { sections?: { columns?: { lines?: { type?: LineType }[] }[] }[] }[]
}

export function countSkeletonTextLines(skeletonData: SkeletonDataLike | undefined): number {
  let count = 0
  for (const page of skeletonData?.pages ?? [])
    for (const section of page.sections ?? [])
      for (const column of section.columns ?? [])
        for (const line of column.lines ?? [])
          if (line.type !== LineType.BLOCK && ++count > 1) return count
  return count
}

/// Single lines keep their native alignment: a slightly-short row clips a
/// bottom-aligned line at its top in Excel too, so only multi-line overflow
/// re-anchors.
export function shouldAnchorTextTop(
  lineCount: number,
  totalHeight: number,
  cellHeight: number,
): boolean {
  return lineCount > 1 && totalHeight > cellHeight
}

export interface TextDrawProps {
  text: string
  fontStyle: string
  warp?: boolean
  vAlign?: VerticalAlign
  width: number
  height: number
}

export function effectiveTextDrawProps<T extends TextDrawProps>(
  props: T,
  lineCount: number,
  totalHeight: number,
): T {
  if (shouldAnchorTextTop(lineCount, totalHeight, props.height)) {
    return { ...props, vAlign: VerticalAlign.TOP }
  }
  return props
}

interface SimpleSkeleton {
  calculate(): unknown[]
  getTotalHeight(): number
}

let installed = false

export function installCellClipAnchorFix(): void {
  if (installed) return
  installed = true

  const documentsProto = Documents.prototype as unknown as {
    _verticalHandler(
      pageHeight: number,
      pagePaddingTop: number,
      pagePaddingBottom: number,
      verticalAlign: unknown,
    ): number
    getSkeleton?(): { getSkeletonData?(): SkeletonDataLike | undefined } | undefined
  }
  const originalVertical = documentsProto._verticalHandler
  if (typeof originalVertical === 'function') {
    documentsProto._verticalHandler = function (
      this: typeof documentsProto,
      pageHeight: number,
      pagePaddingTop: number,
      pagePaddingBottom: number,
      verticalAlign: unknown,
    ): number {
      const offsetTop = originalVertical.call(
        this,
        pageHeight,
        pagePaddingTop,
        pagePaddingBottom,
        verticalAlign,
      )
      if (offsetTop >= 0) return offsetTop
      // Same single-line rule as shouldAnchorTextTop: a near-fit one-liner
      // keeps its native alignment instead of re-anchoring to the top.
      return excelClipOffsetTop(
        offsetTop,
        countSkeletonTextLines(this.getSkeleton?.()?.getSkeletonData?.()),
      )
    }
  }

  const textClass = Text as unknown as {
    drawWith(ctx: unknown, props: TextDrawProps, skeleton?: SimpleSkeleton): void
  }
  const originalDrawWith = textClass.drawWith
  if (typeof originalDrawWith !== 'function') return
  textClass.drawWith = function (
    this: unknown,
    ctx: unknown,
    props: TextDrawProps,
    skeleton?: SimpleSkeleton,
  ): void {
    if (props.vAlign !== VerticalAlign.TOP) {
      // Measure with an unbounded skeleton — the same layout the original
      // builds for non-TOP alignment, so handing it through costs nothing.
      const measured =
        skeleton ??
        (new DocSimpleSkeleton(
          props.text,
          props.fontStyle,
          Boolean(props.warp),
          props.width,
          Number.POSITIVE_INFINITY,
        ) as unknown as SimpleSkeleton)
      const lines = measured.calculate()
      const effective = effectiveTextDrawProps(props, lines.length, measured.getTotalHeight())
      if (effective !== props) {
        // TOP builds its own height-bounded skeleton that drops the lines
        // beyond the cell, so the unbounded probe must not be passed along.
        return originalDrawWith.call(this, ctx, effective, undefined)
      }
      return originalDrawWith.call(this, ctx, props, measured)
    }
    return originalDrawWith.call(this, ctx, props, skeleton)
  }
}
