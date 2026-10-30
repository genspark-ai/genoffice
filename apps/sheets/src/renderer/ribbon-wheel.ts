/// True when something between `target` and `ribbon` scrolls vertically itself
/// (an open dropdown's list), so the wheel belongs to it.
function scrollsItself(target: Element, ribbon: Element): boolean {
  for (let el: Element | null = target; el && el !== ribbon; el = el.parentElement) {
    if (el.scrollHeight <= el.clientHeight) continue
    const overflow = getComputedStyle(el).overflowY
    if (overflow === 'auto' || overflow === 'scroll') return true
  }
  return false
}

/// The ribbon overflows sideways, and its scrollbar is thin (genoffice#2334): a
/// mouse's vertical wheel over it scrolls it horizontally. Trackpads scroll it
/// natively, and Ctrl/⌘ + wheel keeps zooming.
export function scrollRibbonOnWheel(event: WheelEvent): void {
  if (event.ctrlKey || event.metaKey) return
  if (Math.abs(event.deltaY) <= Math.abs(event.deltaX)) return
  const target = event.target as Element | null
  const ribbon = target?.closest?.('[data-ribbon-body]')
  if (!target || !ribbon || ribbon.scrollWidth <= ribbon.clientWidth) return
  if (scrollsItself(target, ribbon)) return
  event.preventDefault()
  ribbon.scrollLeft += event.deltaY
}
