/**
 * Word breaks a big button's caption over two lines and makes the button as
 * narrow as the wider line. CSS can wrap at a width but cannot find that
 * width, so it is measured here and handed over as the caption's max-width.
 */

/** the narrowest width that still fits the words on two lines */
export function twoLineWidth(words: number[], space: number): number {
  let best = Infinity
  for (let cut = 1; cut < words.length; cut += 1) {
    const line = (part: number[]) => part.reduce((sum, w) => sum + w, 0) + space * (part.length - 1)
    best = Math.min(best, Math.max(line(words.slice(0, cut)), line(words.slice(cut))))
  }
  return best
}

const widths = new Map<string, number>()
let pen: CanvasRenderingContext2D | null = null

const BIG_CAPTION = '.rb-big > span:not(.rb-big-icon, .rb-group-menu-name)'

export function wrapBigCaptions(root: HTMLElement, captions = BIG_CAPTION): void {
  // nothing is laid out (a hidden ribbon, a test DOM): there is nothing to measure against
  const laidOut = root.offsetWidth > 0
  for (const caption of root.querySelectorAll<HTMLElement>(captions)) {
    const words = (caption.textContent ?? '').trim().split(/\s+/)
    let next = ''
    if (words.length >= 2) {
      if (!laidOut) continue
      const font = getComputedStyle(caption).font
      const key = `${font}|${words.join(' ')}`
      let width = widths.get(key)
      if (width === undefined) {
        pen ??= document.createElement('canvas').getContext('2d')
        if (!pen) return
        pen.font = font
        const measure = (text: string) => pen!.measureText(text).width
        width = Math.ceil(twoLineWidth(words.map(measure), measure(' '))) + 1
        widths.set(key, width)
      }
      next = `${width}px`
    }
    // this runs on every render: write only a change, so a settled ribbon is not laid out again
    if (caption.style.maxWidth !== next) caption.style.maxWidth = next
  }
}
