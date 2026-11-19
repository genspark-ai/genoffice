import { wrapBigCaptions } from './ribbon-caption-wrap'

/**
 * Word keeps the ribbon inside the window by taking each group down a ladder
 * as the window narrows: full-height buttons, then a stack of small captioned
 * controls, then icons alone. Which stack goes first is Word's own table:
 * `data-wide` ranks the step from buttons to a stack, `data-collapse` the
 * step from captions to icons, and the unranked go last, from the right.
 */

const WIDE = 'rb-col-wide'
const TIGHT = 'rb-col-tight'
const WIDE_CAPTION = `.${WIDE} .rb-small-label`
const fitted = new WeakMap<HTMLElement, string>()

export interface FitStep {
  stack: HTMLElement
  to: 'stack' | 'icons'
}

/** the steps down the ladder in the order they are taken */
export function fitSteps(stacks: HTMLElement[]): FitStep[] {
  const rank = (value: string | undefined) => Number(value ?? Infinity)
  return stacks
    .flatMap((stack, index) => {
      const wide = stack.dataset.wide ? rank(stack.dataset.wide) : -Infinity
      // a stack has to be a stack before it can lose its captions
      const icons = {
        stack,
        to: 'icons' as const,
        index,
        at: Math.max(wide, rank(stack.dataset.collapse)),
      }
      return stack.dataset.wide
        ? [{ stack, to: 'stack' as const, index, at: wide }, icons]
        : [icons]
    })
    .sort(
      (a, b) =>
        a.at - b.at || Number(a.to === 'icons') - Number(b.to === 'icons') || b.index - a.index,
    )
    .map(({ stack, to }) => ({ stack, to }))
}

const GROUP_TIGHT = 'rb-group-tight'

export function fitRibbon(body: HTMLElement): void {
  const stacks = [...body.querySelectorAll<HTMLElement>('.rb-col')].filter((stack) =>
    stack.querySelector('.rb-small-label'),
  )
  const groups = [...body.querySelectorAll<HTMLElement>('.ribbon-group')].filter((group) =>
    group.querySelector('.rb-big'),
  )
  // the compact ribbon has no captions to give up
  const compact = !!body.closest('.ribbon-compact')
  const state = `${body.clientWidth}|${compact}|${stacks.length}|${groups.length}|${body.textContent?.length}`
  if (fitted.get(body) === state) return
  fitted.set(body, state)
  const captions = (stack: HTMLElement) => stack.querySelectorAll<HTMLElement>('.rb-small-label')
  for (const stack of stacks) {
    stack.classList.remove(TIGHT)
    stack.classList.toggle(WIDE, !compact && !!stack.dataset.wide)
    if (compact) for (const caption of captions(stack)) caption.style.maxWidth = ''
  }
  for (const group of groups) group.classList.remove(GROUP_TIGHT)
  if (compact) return
  wrapBigCaptions(body, WIDE_CAPTION)
  for (const { stack, to } of fitSteps(stacks)) {
    if (body.scrollWidth <= body.clientWidth) return
    if (to === 'icons') stack.classList.add(TIGHT)
    else {
      stack.classList.remove(WIDE)
      for (const caption of captions(stack)) caption.style.maxWidth = ''
    }
  }
  // still too wide with every stack down to icons: the big buttons go next, from the right
  for (const group of groups.reverse()) {
    if (body.scrollWidth <= body.clientWidth) return
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
  if (body.closest('.ribbon-compact') || body.querySelector(`.${TIGHT}, .${GROUP_TIGHT}`)) return 0
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
