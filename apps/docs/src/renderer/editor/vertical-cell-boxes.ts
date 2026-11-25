import type { EditorView } from '@tiptap/pm/view'

export type VerticalCellMode = 'vertical-rl' | 'sideways-lr'

export interface BoxSides {
  top: string
  right: string
  bottom: string
  left: string
}

/**
 * The sides of a horizontal paragraph box as they lie once the text is turned:
 * vertical-rl turns it a quarter right (space after ends up on the left),
 * sideways-lr a quarter left.
 */
export function turnedSides(mode: VerticalCellMode, h: BoxSides): BoxSides {
  return mode === 'vertical-rl'
    ? { top: h.left, right: h.top, bottom: h.right, left: h.bottom }
    : { top: h.right, right: h.bottom, bottom: h.left, left: h.top }
}

const WRAPPERS = '.cell-vert, .cell-clip[style*="writing-mode"]'
const PROBE = 'cell-flow-probe'
const SIDES = ['top', 'right', 'bottom', 'left'] as const

function readSides(cs: CSSStyleDeclaration, prop: 'margin' | 'padding'): BoxSides {
  const out = {} as BoxSides
  for (const s of SIDES) out[s] = cs.getPropertyValue(`${prop}-${s}`) || '0px'
  return out
}

/** the inline declarations that carry a turned box; the stylesheet applies them through [data-vbox] */
export function turnedBoxVars(mode: VerticalCellMode, margin: BoxSides, padding: BoxSides): string {
  const m = turnedSides(mode, margin)
  const p = turnedSides(mode, padding)
  return SIDES.map((s) => `--vb-m${s[0]}:${m[s]};--vb-p${s[0]}:${p[s]}`).join(';')
}

function modeOf(wrapper: HTMLElement): VerticalCellMode | null {
  const wm = wrapper.style.writingMode
  return wm === 'vertical-rl' || wm === 'sideways-lr' ? wm : null
}

/**
 * Paragraph spacing and indents are physical margins and paddings, so in a
 * vertical cell the space after a paragraph would shorten its lines instead of
 * parting it from the next one. Each pass reads the boxes with the cell laid
 * out horizontally and writes them back turned with the text. True when a box
 * changed, so the caller measures again.
 */
export function syncVerticalCellBoxes(root: HTMLElement, view?: EditorView | null): boolean {
  const wrappers = Array.from(root.querySelectorAll<HTMLElement>(WRAPPERS)).filter(modeOf)
  if (wrappers.length === 0) return false
  const obs = (view as unknown as { domObserver?: { stop(): void; start(): void } } | null)
    ?.domObserver
  obs?.stop()
  let changed = false
  try {
    for (const w of wrappers) w.classList.add(PROBE)
    const next: { el: HTMLElement; vars: string }[] = []
    for (const w of wrappers) {
      const mode = modeOf(w)!
      for (const el of Array.from(w.children) as HTMLElement[]) {
        if (el.tagName === 'TABLE' || !(el instanceof HTMLElement)) continue
        const cs = getComputedStyle(el)
        next.push({
          el,
          vars: turnedBoxVars(mode, readSides(cs, 'margin'), readSides(cs, 'padding')),
        })
      }
    }
    for (const w of wrappers) w.classList.remove(PROBE)
    for (const { el, vars } of next) {
      if (el.dataset.vbox === vars && el.style.getPropertyValue('--vb-mt')) continue
      for (const decl of vars.split(';')) {
        const i = decl.indexOf(':')
        el.style.setProperty(decl.slice(0, i), decl.slice(i + 1))
      }
      el.dataset.vbox = vars
      changed = true
    }
  } finally {
    obs?.start()
  }
  return changed
}
