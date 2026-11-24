import { wrapBigCaptions } from './ribbon-caption-wrap'

/**
 * Word keeps the ribbon inside the window by taking each group down a ladder
 * as the window narrows: full-height buttons, then a stack of small captioned
 * controls, then icons alone, and at last one menu button named after the
 * group. Which step comes first is Word's own table, written as ranks in one
 * number space per tab: `data-wide` (buttons to a stack) and `data-collapse`
 * (captions to icons) on a stack, `data-icon` on a big button that goes
 * straight to its icon, `data-drop` on a control that leaves the ribbon, and
 * `data-menu` on a group that becomes a menu button. `data-tight` ranks a
 * group Word does not have, whose big buttons give up their captions. Stacks
 * without a rank lose their captions from the right, after the ranks below
 * UNRANKED.
 */

const WIDE = 'rb-col-wide'
const TIGHT = 'rb-col-tight'
const BIG_TIGHT = 'rb-big-tight'
const GONE = 'rb-fit-gone'
const GROUP_TIGHT = 'rb-group-tight'
export const GROUP_MENU = 'rb-group-menu'
const WIDE_CAPTION = `.${WIDE} .rb-small-label`
const UNRANKED = 1000
const fitted = new WeakMap<HTMLElement, string>()

export interface FitStep {
  el: HTMLElement
  to: 'stack' | 'icons' | 'icon' | 'tight' | 'gone' | 'menu'
}

/** the steps down the ladder in the order they are taken */
export function fitSteps(els: HTMLElement[]): FitStep[] {
  const rank = (value: string | undefined) => Number(value ?? UNRANKED)
  // at one rank: buttons to stacks, captions to icons, and only then a whole group
  const weight = { stack: 0, icons: 1, icon: 1, tight: 1, gone: 2, menu: 3 }
  return els
    .flatMap((el, index): (FitStep & { index: number; at: number })[] => {
      const { wide, collapse, icon, drop, menu, tight } = el.dataset
      if (menu) return [{ el, to: 'menu', index, at: rank(menu) }]
      if (icon) return [{ el, to: 'icon', index, at: rank(icon) }]
      if (drop) return [{ el, to: 'gone', index, at: rank(drop) }]
      if (tight) return [{ el, to: 'tight', index, at: rank(tight) }]
      // a stack has to be a stack before it can lose its captions
      const icons = {
        el,
        to: 'icons' as const,
        index,
        at: Math.max(wide ? rank(wide) : -Infinity, rank(collapse)),
      }
      return wide ? [{ el, to: 'stack', index, at: rank(wide) }, icons] : [icons]
    })
    .sort((a, b) => a.at - b.at || weight[a.to] - weight[b.to] || b.index - a.index)
    .map(({ el, to }) => ({ el, to }))
}

const RANKED = '[data-icon], [data-drop], .ribbon-group[data-menu], .ribbon-group[data-tight]'

export function fitRibbon(body: HTMLElement): void {
  // in document order, which is what "from the right" is counted in
  const rungs = [...body.querySelectorAll<HTMLElement>(`.rb-col, ${RANKED}`)].filter(
    (el) => el.matches(RANKED) || el.querySelector('.rb-small-label'),
  )
  const ranked = rungs.filter((el) => el.matches(RANKED))
  const stacks = rungs.filter((el) => !el.matches(RANKED))
  const groups = [...body.querySelectorAll<HTMLElement>('.ribbon-group')].filter((group) =>
    group.querySelector('.rb-big'),
  )
  // the compact ribbon has no captions to give up
  const compact = !!body.closest('.ribbon-compact')
  const state = `${body.clientWidth}|${compact}|${stacks.length}|${ranked.length}|${groups.length}|${body.textContent?.length}`
  if (fitted.get(body) === state) return
  fitted.set(body, state)
  const captions = (stack: HTMLElement) => stack.querySelectorAll<HTMLElement>('.rb-small-label')
  for (const stack of stacks) {
    stack.classList.remove(TIGHT)
    stack.classList.toggle(WIDE, !compact && !!stack.dataset.wide)
    if (compact) for (const caption of captions(stack)) caption.style.maxWidth = ''
  }
  for (const el of ranked) el.classList.remove(BIG_TIGHT, GONE, GROUP_MENU, GROUP_TIGHT)
  for (const group of groups) group.classList.remove(GROUP_TIGHT)
  if (compact) return
  wrapBigCaptions(body, WIDE_CAPTION)
  for (const { el, to } of fitSteps(rungs)) {
    if (body.scrollWidth <= body.clientWidth) return
    // what a group's menu holds keeps its captions
    if (to !== 'menu' && el.closest(`.${GROUP_MENU}`)) continue
    if (to === 'icons') el.classList.add(TIGHT)
    else if (to === 'icon') el.classList.add(BIG_TIGHT)
    else if (to === 'gone') el.classList.add(GONE)
    else if (to === 'tight') el.classList.add(GROUP_TIGHT)
    else if (to === 'menu') {
      el.classList.add(GROUP_MENU)
      for (const inner of el.querySelectorAll(`.${TIGHT}, .${BIG_TIGHT}`))
        inner.classList.remove(TIGHT, BIG_TIGHT)
    } else {
      el.classList.remove(WIDE)
      for (const caption of captions(el)) caption.style.maxWidth = ''
    }
  }
  // nothing left on the ladder: the big buttons still showing give up their captions, from the right
  for (const group of groups.reverse()) {
    if (body.scrollWidth <= body.clientWidth) return
    if (group.classList.contains(GROUP_MENU)) continue
    group.classList.add(GROUP_TIGHT)
  }
}

/** a style card and the gap after it */
const STYLE_CARD = 77
/** the opener at the end of the cards, and what separates it from them */
const STYLE_MORE = 18
const STRIP_GAP = 7
/** Word shows its Styles button rather than a card or two */
const MIN_CARDS = 3

/**
 * How many style cards fit in place of the Styles button without making the ribbon
 * scroll; null when the tab showing has no Styles group to measure.
 */
export function styleCardRoom(body: HTMLElement): number | null {
  const anchor = body.querySelector<HTMLElement>('.style-gallery-wrap')
  if (!anchor) return null
  if (
    body.closest('.ribbon-compact') ||
    body.querySelector(`.${TIGHT}, .${GROUP_TIGHT}, .${BIG_TIGHT}, .${GONE}, .${GROUP_MENU}`)
  )
    return 0
  const last = [...body.querySelectorAll<HTMLElement>('.ribbon-group')].at(-1)
  if (!last) return 0
  // measured along the reading direction: in a right-to-left UI the groups run from the right
  const rtl = getComputedStyle(body).direction === 'rtl'
  const start = (el: Element) => {
    const box = el.getBoundingClientRect()
    return rtl ? -box.right : box.left
  }
  const end = (el: Element) => {
    const box = el.getBoundingClientRect()
    return rtl ? -box.left : box.right
  }
  // what the cards take is how far they push the Styles button: strip, margin and gap alike
  const strip = anchor.querySelector<HTMLElement>('.style-gallery-inline')
  const next = strip?.nextElementSibling
  const shown = strip && next ? start(next) - start(strip) : 0
  // the cards stand in for the Styles button, and only their narrow opener stays
  const button = anchor.querySelector<HTMLElement>('.rb-styles-gallery')
  const swapped =
    strip || !button ? 0 : button.getBoundingClientRect().width - STYLE_MORE - STRIP_GAP
  const padding = parseFloat(getComputedStyle(body).paddingInlineEnd) || 0
  const used = end(last) - start(body) + padding - shown - swapped
  const cards = Math.floor((body.clientWidth - used - 4) / STYLE_CARD)
  return cards < MIN_CARDS ? 0 : Math.min(8, cards)
}
