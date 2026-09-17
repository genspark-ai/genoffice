/**
 * A context-menu submenu (Clear, Insert, ...) that the pointer leaves and
 * re-enters within Univer's 500 ms close delay never shows again: the
 * re-enter resets the position-ready flag while the submenu is still
 * mounted, and the positioning effect only re-runs on a visibility toggle
 * or a window resize/scroll, so nothing sets the flag back (genoffice#337).
 *
 * A submenu still hidden after the frame Univer uses for its first
 * measurement is stuck; a synthetic scroll event on it reaches the effect's
 * capture listener on window, which re-measures and shows it.
 */
const SUBMENU_SELECTOR = '[data-u-context-menu-submenu]'

export function installContextSubmenuReopenFix(root: HTMLElement = document.body): {
  dispose(): void
} {
  let frame = 0
  const nudged = new WeakSet<Element>()

  const check = () => {
    frame = 0
    for (const submenu of root.querySelectorAll<HTMLElement>(SUBMENU_SELECTOR)) {
      if (submenu.style.visibility !== 'hidden') {
        nudged.delete(submenu)
        continue
      }
      if (nudged.has(submenu)) continue
      nudged.add(submenu)
      submenu.dispatchEvent(new Event('scroll'))
    }
  }

  const schedule = () => {
    if (frame) return
    // a fresh submenu mounts hidden and measures itself on the next frame
    frame = requestAnimationFrame(() => {
      frame = requestAnimationFrame(check)
    })
  }

  const observer = new MutationObserver(schedule)
  observer.observe(root, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ['style'],
  })
  return {
    dispose() {
      observer.disconnect()
      if (frame) cancelAnimationFrame(frame)
      frame = 0
    },
  }
}
