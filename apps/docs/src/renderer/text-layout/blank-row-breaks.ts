import type { Node as PmNode } from '@tiptap/pm/model'

/** content offsets of soft breaks opening a text-less row (first child, or right
 *  after another soft break): `sized` carry their own run size, `inherited` carry
 *  neither size nor face and take the paragraph's inherited face in Word;
 *  `leadingEnd` is the offset where the paragraph's leading soft breaks stop (a
 *  page / column break ends them: the text after it is a continuation line) */
export function blankRowBreaks(node: PmNode): {
  sized: number[]
  inherited: number[]
  leadingEnd: number
} {
  const sized: number[] = []
  const inherited: number[] = []
  let leadingEnd = 0
  let rowStart = true
  for (let i = 0, off = 0; i < node.childCount; i++) {
    const child = node.child(i)
    const isBreak = child.type.name === 'hardBreak'
    const soft = isBreak && !child.attrs.pageBreak && !child.attrs.colBreak
    if (soft && rowStart) {
      const style = child.marks.find((m) => m.type.name === 'docTextStyle')
      if (style?.attrs.sizeHalfPoints != null) sized.push(off)
      else if (!style?.attrs.font && !style?.attrs.fontAscii) inherited.push(off)
    }
    // a page / column break ends its row too, so the next break opens a blank one
    rowStart = isBreak
    off += child.nodeSize
    if (soft && leadingEnd === off - child.nodeSize) leadingEnd = off
  }
  return { sized, inherited, leadingEnd }
}
