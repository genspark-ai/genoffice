import { setCommentPresentation } from './comment-presentation'

export interface PageComment {
  bubble: HTMLElement
  desiredTop: number
  drawLeader: (top: number) => void
}

/** Measure and arrange complete discussions within one physical page. */
export function layoutCommentPage(
  viewport: HTMLElement,
  content: HTMLElement,
  items: PageComment[],
  height: number,
  gap: number,
): void {
  let heights: number[] = []
  let needed = 0
  for (const lines of [15, 5, 3, 1] as const) {
    heights = items.map(({ bubble }) =>
      setCommentPresentation(
        bubble,
        lines,
        bubble.classList.contains('active') || bubble.classList.contains('comment-popup'),
        height,
      ),
    )
    needed = heights.reduce((sum, value) => sum + value, 0) + Math.max(0, items.length - 1) * gap
    if (needed <= height) break
  }

  const scrolling = needed > height
  const tops: number[] = []
  let bottom = 0
  for (let i = 0; i < items.length; i++) {
    tops[i] = scrolling
      ? bottom
      : Math.max(bottom, Math.min(items[i].desiredTop, height - heights[i]))
    bottom = tops[i] + heights[i] + gap
  }
  if (!scrolling) {
    bottom = height
    for (let i = items.length - 1; i >= 0; i--) {
      tops[i] = Math.min(tops[i], bottom - heights[i])
      bottom = tops[i] - gap
    }
  }
  content.style.height = `${Math.max(height, needed)}px`
  viewport.classList.toggle('comment-page-scrolling', scrolling)
  items.forEach(({ bubble }, i) => {
    bubble.style.top = `${tops[i]}px`
  })
  const draw = () => items.forEach((item, i) => item.drawLeader(tops[i] - viewport.scrollTop))
  viewport.onscroll = draw
  draw()
}
