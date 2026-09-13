import type { ParaFrame, TextFlowDirection } from '@genoffice/docx-engine'

/** CSS writing mode for a w:textDirection value (frames and section blocks share it) */
export function textFlowCss(dir: TextFlowDirection | null | undefined): string[] {
  if (dir === 'tbRl') return ['writing-mode:vertical-rl']
  if (dir === 'tbRlV') return ['writing-mode:vertical-rl', 'text-orientation:sideways']
  if (dir === 'btLr') return ['writing-mode:sideways-lr']
  return []
}

const pt = (twips: number) => `${(twips / 20).toFixed(2)}pt`

/**
 * Inline styles + classes for a w:framePr paragraph. The frame is a fixed-width
 * box floated at its horizontal position (page-anchored x is measured from the
 * page edge, so the section's left margin is subtracted through the
 * --doc-margin-left variable); body text wraps beside it for wrap=around/
 * through/tight/auto. Vertical anchors other than "text" are approximated by the
 * frame's flow position.
 */
export function paraFrameCss(
  frame: ParaFrame,
  dir: TextFlowDirection | null | undefined,
): { classes: string[]; styles: string[] } {
  const styles: string[] = ['box-sizing:border-box', `width:${pt(frame.wTwips)}`]
  const flow = textFlowCss(dir)
  if (frame.hTwips) {
    // sideways text: the height is the line length, so it is fixed even for
    // atLeast (an auto orthogonal box would take the page height instead)
    const fixed = frame.hRule === 'exact' || flow.length > 0
    styles.push(`${fixed ? 'height' : 'min-height'}:${pt(frame.hTwips)}`)
    if (frame.hRule === 'exact') styles.push('overflow:hidden')
  }
  styles.push(...flow)
  // a centered frame cannot float (auto margins do not center floats): it takes its own line
  const wraps =
    frame.wrap === 'around' ||
    frame.wrap === 'through' ||
    frame.wrap === 'tight' ||
    frame.wrap === 'auto'
  const floats = wraps && frame.xAlign !== 'center'
  const side = frame.xAlign === 'right' || frame.xAlign === 'outside' ? 'right' : 'left'
  const classes = ['doc-para-frame']
  if (floats) {
    classes.push('doc-para-frame-float')
    styles.push(`float:${side}`)
    if (frame.hSpaceTwips)
      styles.push(`margin-${side === 'left' ? 'right' : 'left'}:${pt(frame.hSpaceTwips)}`)
  }
  if (frame.xAlign === 'center') {
    styles.push('margin-left:auto', 'margin-right:auto')
  } else if (!frame.xAlign) {
    const x = pt(frame.xTwips)
    styles.push(
      frame.hAnchor === 'margin' || frame.hAnchor === 'text'
        ? `margin-left:${x}`
        : `margin-left:calc(${x} - var(--doc-margin-left, 0px))`,
    )
  }
  if (frame.vAnchor === 'text' && frame.yTwips > 0) styles.push(`margin-top:${pt(frame.yTwips)}`)
  if (frame.vSpaceTwips) styles.push(`margin-bottom:${pt(frame.vSpaceTwips)}`)
  return { classes, styles }
}
