/** Assign x-sorted same-baseline extents to z-layers. Two runs drawn over each other
    (an earlier multi-line edit that overflowed onto the next line) must not be read as
    one line: heavy x-overlap forbids sharing a chain, and a non-overlapping extent
    continues the chain whose tail is nearest — glyph spacing beats a cross-layer jump,
    so a longer stacked run keeps its own tail instead of leaking into the other layer.

    `tieMargin`: when two chains' tail gaps are this close, the chain with more members
    wins — a short overlay ending inside a word gap must not steal the line's next word
    (the real text line is a many-span chain, an overlay is usually one span). */
export function chainLayers(extents: { left: number; right: number }[], tieMargin = 0): number[] {
  const assign: number[] = []
  const tails: { left: number; right: number; size: number }[] = []
  for (const r of extents) {
    const allowed: { chain: number; gap: number; size: number }[] = []
    for (let c = 0; c < tails.length; c++) {
      const t = tails[c]!
      const overlap = Math.min(t.right, r.right) - Math.max(t.left, r.left)
      if (overlap > 0.3 * Math.min(t.right - t.left, r.right - r.left)) continue
      allowed.push({ chain: c, gap: Math.abs(r.left - t.right), size: t.size })
    }
    allowed.sort((a, b) => a.gap - b.gap)
    let pick = allowed[0]
    for (const a of allowed.slice(1)) {
      if (a.gap - allowed[0]!.gap > tieMargin) break
      if (a.size > pick!.size) pick = a
    }
    if (!pick) {
      tails.push({ left: r.left, right: r.right, size: 1 })
      assign.push(tails.length - 1)
    } else {
      const t = tails[pick.chain]!
      tails[pick.chain] = { left: r.left, right: r.right, size: t.size + 1 }
      assign.push(pick.chain)
    }
  }
  return assign
}
