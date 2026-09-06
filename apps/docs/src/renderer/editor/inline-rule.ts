import type { Run } from '@genoffice/docx-engine'

export type InlineRule = NonNullable<NonNullable<Run['image']>['rule']>

export const INLINE_RULE_CLASS = 'doc-inline-rule'

/** inline declarations of a VML horizontal rule run: document stroke color/size
 *  (not chrome) and the run's own font size, which sets the rule line's height */
export function inlineRuleDecls(rule: InlineRule): string[] {
  return [
    rule.colorHex ? `--doc-rule-color:#${rule.colorHex}` : '',
    rule.thicknessPx ? `--doc-rule-h:${rule.thicknessPx}px` : '',
    rule.sizeHalfPoints ? `font-size:${rule.sizeHalfPoints / 2}pt` : '',
    rule.widthPx ? `width:${rule.widthPx}px;max-width:100%` : '',
    rule.widthPx && rule.align ? 'margin-left:auto' : '',
    rule.widthPx && rule.align === 'center' ? 'margin-right:auto' : '',
  ].filter(Boolean)
}

/** the same declarations as a React style object */
export function inlineRuleStyle(rule: InlineRule): Record<string, string> {
  const out: Record<string, string> = {}
  if (rule.colorHex) out['--doc-rule-color'] = `#${rule.colorHex}`
  if (rule.thicknessPx) out['--doc-rule-h'] = `${rule.thicknessPx}px`
  if (rule.sizeHalfPoints) out.fontSize = `${rule.sizeHalfPoints / 2}pt`
  if (rule.widthPx) {
    out.width = `${rule.widthPx}px`
    out.maxWidth = '100%'
    if (rule.align) out.marginLeft = 'auto'
    if (rule.align === 'center') out.marginRight = 'auto'
  }
  return out
}

/** rule attrs restored from a pasted rule span (renderHTML bakes them as data attributes) */
export function parseInlineRuleEl(el: HTMLElement): Record<string, unknown> | false {
  const raw = el.getAttribute('data-inline-rule')
  const xml = el.getAttribute('data-xml')
  if (!raw || !xml) return false
  try {
    const rule = JSON.parse(raw) as InlineRule
    return { dataUrl: '', xml, rule, leadRule: el.hasAttribute('data-lead-rule') }
  } catch {
    return false
  }
}
