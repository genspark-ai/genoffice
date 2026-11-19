import type { PageCut } from '../pagination-cuts'

/** A break that shows its own labelled mark gets no rule; every other page start does */
export interface DraftBreak {
  page: number
  /** the block that opens the page, when the break falls between two blocks */
  block?: number
  /** flow y of the break (px at zoom 1 from the first page's content top) */
  y: number
}

interface DraftBlock {
  section?: number
}

export interface DraftBreakMarks {
  /** the block ends in a page break character */
  endsInPageBreak(block: number): boolean
  /** the block shows a mark for its own page break: one inherited from a style has none */
  startsWithMark(block: number): boolean
}

const NO_MARKS: DraftBreakMarks = { endsInPageBreak: () => false, startsWithMark: () => false }

export function draftAutoBreaks(
  cuts: readonly PageCut[],
  blocks: readonly DraftBlock[],
  marks: DraftBreakMarks = NO_MARKS,
): DraftBreak[] {
  const out: DraftBreak[] = []
  cuts.forEach((cut, page) => {
    if (page === 0 || cut.lead.kind === 'tail') return
    // a blank page and the page after it start at the same place
    if (out.length > 0 && Math.abs(out[out.length - 1].y - cut.start) < 0.5) return
    if (cut.lead.kind !== 'block') {
      out.push({ page, y: cut.start })
      return
    }
    const i = cut.lead.block
    const marked =
      marks.startsWithMark(i) ||
      (i > 0 &&
        ((blocks[i]?.section ?? 0) !== (blocks[i - 1]?.section ?? 0) ||
          marks.endsInPageBreak(i - 1)))
    if (!marked) out.push({ page, block: i, y: cut.start })
  })
  return out
}

export function syncDraftBreaks(layer: HTMLElement, tops: readonly number[]): void {
  while (layer.children.length > tops.length) layer.lastElementChild?.remove()
  tops.forEach((top, i) => {
    let rule = layer.children[i] as HTMLElement | undefined
    if (!rule) {
      rule = document.createElement('div')
      rule.className = 'draft-break'
      layer.appendChild(rule)
    }
    rule.style.top = `${top.toFixed(1)}px`
  })
}
