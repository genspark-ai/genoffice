/// Snapshots of the floating visuals (charts, pictures, shapes) for the print
/// layout: the live float DOM is cloned with the stylesheet rules it matches,
/// custom properties resolved to their light-theme values so a dark UI prints
/// the way Excel would. The layout places each snapshot by its anchor frame.

import type { InstalledVisualFrame } from './WorkbookVisuals'

export interface PrintVisual {
  readonly id: string
  readonly fromRow: number
  readonly fromColumn: number
  readonly toRow: number
  readonly toColumn: number
  /// px inside the anchor cell
  readonly offsetXPx: number
  readonly offsetYPx: number
  readonly widthPx: number
  readonly heightPx: number
  readonly html: string
}

export interface PrintVisualSnapshot {
  readonly visuals: readonly PrintVisual[]
  readonly css: string
}

const SETTLE_POLL_MS = 60

function visualNode(doc: Document, id: string): HTMLElement | null {
  const node = doc.querySelector<HTMLElement>(`[data-print-visual="${id}"]`)
  return node?.isConnected && node.firstElementChild ? node : null
}

function laidOut(doc: Document, id: string): boolean {
  const child = visualNode(doc, id)?.firstElementChild
  return child !== undefined && child !== null && child.getBoundingClientRect().width > 0
}

/// Float DOM attaches a beat after install; wait for every frame's node to
/// exist and lay out (charts that never paint, e.g. an empty figure, still
/// have their frame). Gives up quietly: a missing node prints nothing.
export async function settleVisualNodes(
  doc: Document,
  frames: readonly InstalledVisualFrame[],
  timeoutMs = 2000,
): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    if (frames.every((frame) => laidOut(doc, frame.visual.id))) return
    await new Promise((resolve) => setTimeout(resolve, SETTLE_POLL_MS))
  }
}

export function snapshotPrintVisuals(
  doc: Document,
  frames: readonly InstalledVisualFrame[],
): PrintVisualSnapshot {
  const visuals: PrintVisual[] = []
  const roots: Element[] = []
  for (const frame of frames) {
    const node = visualNode(doc, frame.visual.id)
    if (!node) continue
    roots.push(node)
    const clone = node.cloneNode(true) as HTMLElement
    stripInteractive(clone)
    visuals.push({
      id: frame.visual.id,
      fromRow: frame.fromRow,
      fromColumn: frame.fromColumn,
      toRow: frame.toRow,
      toColumn: frame.toColumn,
      offsetXPx: frame.marginX,
      offsetYPx: frame.marginY,
      widthPx: frame.width,
      heightPx: frame.height,
      html: clone.outerHTML,
    })
  }
  return { visuals, css: roots.length ? collectPrintCss(doc, roots) : '' }
}

function stripInteractive(root: HTMLElement): void {
  for (const el of Array.from(
    root.querySelectorAll('button, input, textarea, select, [data-print-skip]'),
  )) {
    el.remove()
  }
  for (const el of Array.from(root.querySelectorAll('*'))) {
    el.removeAttribute('contenteditable')
    el.removeAttribute('tabindex')
    el.removeAttribute('draggable')
  }
}

const STATE_PSEUDO = /:(hover|focus|focus-within|focus-visible|active)\b/
const PSEUDO_ELEMENT = /::?(before|after|marker|placeholder|selection|scrollbar[\w-]*)\b/g

/// The rules of the document's stylesheets that apply inside `roots`,
/// conditional groups kept, `var()` resolved against the `:root` (light) values.
export function collectPrintCss(doc: Document, roots: readonly Element[]): string {
  const out: string[] = []
  const vars = new Map<string, string>()
  for (const sheet of Array.from(doc.styleSheets)) {
    let rules: CSSRuleList
    try {
      rules = sheet.cssRules
    } catch {
      continue
    }
    collectRules(Array.from(rules), roots, out, vars)
  }
  return resolveVars(out.join('\n'), vars)
}

function collectRules(
  rules: readonly CSSRule[],
  roots: readonly Element[],
  out: string[],
  vars: Map<string, string>,
): void {
  for (const rule of rules) {
    if (isStyleRule(rule)) {
      if (rule.selectorText === ':root') {
        for (let i = 0; i < rule.style.length; i += 1) {
          const name = rule.style.item(i)
          if (name.startsWith('--')) vars.set(name, rule.style.getPropertyValue(name).trim())
        }
        continue
      }
      const kept = rule.selectorText
        .split(',')
        .map((part) => part.trim())
        .filter((part) => part && !STATE_PSEUDO.test(part) && matchesAny(part, roots))
      if (kept.length === 0) continue
      const body = rule.cssText.slice(rule.cssText.indexOf('{'))
      out.push(`${kept.join(', ')} ${body}`)
      continue
    }
    if (isGroupingRule(rule)) {
      const inner: string[] = []
      collectRules(Array.from(rule.cssRules), roots, inner, vars)
      if (inner.length === 0) continue
      const head = rule.cssText.slice(0, rule.cssText.indexOf('{')).trim()
      out.push(`${head} {\n${inner.join('\n')}\n}`)
    }
  }
}

function isStyleRule(rule: CSSRule): rule is CSSStyleRule {
  return typeof (rule as CSSStyleRule).selectorText === 'string'
}

function isGroupingRule(rule: CSSRule): rule is CSSGroupingRule {
  return (
    !isStyleRule(rule) &&
    typeof (rule as CSSGroupingRule).cssRules === 'object' &&
    !rule.cssText.startsWith('@keyframes') &&
    !rule.cssText.startsWith('@font-face')
  )
}

function matchesAny(selector: string, roots: readonly Element[]): boolean {
  const plain = selector.replace(PSEUDO_ELEMENT, '')
  if (!plain) return false
  try {
    return roots.some((root) => root.matches(plain) || root.querySelector(plain) !== null)
  } catch {
    return false
  }
}

const VAR = /var\(\s*(--[\w-]+)\s*(?:,\s*([^()]*(?:\([^()]*\))?[^()]*))?\)/g

function resolveVars(css: string, vars: ReadonlyMap<string, string>): string {
  let text = css
  for (let pass = 0; pass < 4 && VAR.test(text); pass += 1) {
    VAR.lastIndex = 0
    text = text.replace(VAR, (whole, name: string, fallback: string | undefined) => {
      const value = vars.get(name)
      if (value !== undefined && value !== '') return value
      return fallback !== undefined ? fallback.trim() : whole
    })
    VAR.lastIndex = 0
  }
  return text
}
