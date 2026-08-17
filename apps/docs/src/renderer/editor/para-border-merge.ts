/**
 * Word merges adjacent paragraphs whose border and shading settings are
 * identical into one border group (ECMA-376 §17.3.1.24): top/bottom lines draw
 * only at the group's edges (no w:between = no inner lines). Display-only —
 * the merge is applied as decoration classes; the borders attrs (which the
 * save path serializes back to w:pBdr) are never touched.
 */

/** the paragraph attrs Word's border-group merging compares */
export interface ParaBorderAttrs {
  borders?: string | null
  /** JSON per-side {color?,szPt?} as stored in the borderLines attr */
  borderLines?: string | null
  shadingFill?: string | null
}

/** canonical per-side form of the borderLines JSON (key-order insensitive) */
function normLines(raw: string | null | undefined): string {
  if (!raw) return ''
  try {
    const parsed = JSON.parse(raw) as Record<string, { color?: string; szPt?: number } | undefined>
    return (['t', 'b', 'l', 'r'] as const)
      .map((side) => {
        const l = parsed?.[side]
        return l ? `${side}:${l.color ?? ''}/${l.szPt ?? ''}` : ''
      })
      .join('|')
  } catch {
    return raw
  }
}

export function sameBorderGroup(a: ParaBorderAttrs, b: ParaBorderAttrs): boolean {
  if (!a.borders || !b.borders) return false
  return (
    a.borders === b.borders &&
    (a.shadingFill ?? null) === (b.shadingFill ?? null) &&
    normLines(a.borderLines) === normLines(b.borderLines)
  )
}

/**
 * Per-paragraph border suppression for a run of adjacent top-level paragraphs:
 * same group as the next → its bottom border is an inner boundary (suppress);
 * same group as the previous → suppress its top border.
 */
export function borderMergeFlags(
  paras: ParaBorderAttrs[],
): Array<{ suppressTop: boolean; suppressBottom: boolean }> {
  return paras.map((p, i) => ({
    suppressTop: i > 0 && sameBorderGroup(paras[i - 1], p),
    suppressBottom: i < paras.length - 1 && sameBorderGroup(p, paras[i + 1]),
  }))
}
