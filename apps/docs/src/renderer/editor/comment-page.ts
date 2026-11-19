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
  scrollTop = viewport.scrollTop,
): void {
  const expanded = items.map(
    ({ bubble }) =>
      bubble.classList.contains('active') || bubble.classList.contains('comment-popup'),
  )
  const contentOffsets = items.map(({ bubble }) => bubble.scrollTop)
  const totalHeight = (values: number[]) =>
    values.reduce((sum, value) => sum + value, 0) + Math.max(0, items.length - 1) * gap
  let heights: number[] = []
  let needed = 0
  for (const lines of [15, 5, 3, 1] as const) {
    heights = items.map(({ bubble }, i) =>
      setCommentPresentation(bubble, lines, expanded[i], height),
    )
    needed = totalHeight(heights)
    if (needed <= height) break
  }

  if (needed > height && expanded.some(Boolean)) {
    const activeCount = expanded.filter(Boolean).length
    const available =
      (height - totalHeight(heights.map((value, i) => (expanded[i] ? 0 : value)))) / activeCount
    for (let i = 0; i < items.length; i++) {
      if (!expanded[i]) continue
      // Retain a readable discussion even when its neighbors require page scrolling.
      const minimum = setCommentPresentation(items[i].bubble, 3, false, height)
      heights[i] = setCommentPresentation(
        items[i].bubble,
        1,
        true,
        Math.min(height, Math.max(minimum, available)),
      )
    }
    needed = totalHeight(heights)
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
    bubble.scrollTop = contentOffsets[i]
  })
  viewport.scrollTop = scrollTop
  items.forEach(({ bubble }, i) => {
    if (!expanded[i]) return
    if (tops[i] < viewport.scrollTop) viewport.scrollTop = tops[i]
    else if (tops[i] + bubble.offsetHeight > viewport.scrollTop + height)
      viewport.scrollTop = tops[i] + bubble.offsetHeight - height
  })
  const draw = () => items.forEach((item, i) => item.drawLeader(tops[i] - viewport.scrollTop))
  viewport.onscroll = draw
  draw()
}
