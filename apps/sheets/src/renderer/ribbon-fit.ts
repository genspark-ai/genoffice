/// Adaptive ribbon, in two steps that only start once the band would scroll:
///
/// 1. The band goes tight (TIGHT_CLASS): narrower padding, gaps and icon
///    buttons. Where everything fits it is never set, so the band keeps the
///    metrics it shares with Docs and Slides.
/// 2. If it still overflows, groups show their small labeled buttons and menu
///    triggers as icons only, the rightmost group first, stopping as soon as
///    the band fits. Every affected control carries `data-tip`, so its name
///    still shows on hover.
///
/// Widen the window and both steps undo. Large buttons are left to the compact
/// density (⌥⌘K), which keeps its own metrics.

export const TIGHT_CLASS = 'ribbon-tight'
export const ICONS_ONLY_CLASS = 'icons-only'

interface FitClassList {
  readonly classList: Pick<DOMTokenList, 'add' | 'remove'>
}

interface FitRibbon extends FitClassList {
  readonly scrollWidth: number
  readonly clientWidth: number
  querySelectorAll(selector: string): ArrayLike<FitClassList>
}

export interface FitResult {
  /// the tight metrics are on
  readonly tight: boolean
  /// how many groups are icon-only, counted from the right
  readonly iconsOnly: number
}

function overflows(ribbon: FitRibbon): boolean {
  // 1px slack: a fractional --ui-scale zoom rounds scrollWidth up
  return ribbon.scrollWidth > ribbon.clientWidth + 1
}

/// Fit one band. Starts from the untouched band, so a wider window gets its
/// metrics and labels back. The class changes and the width reads all happen
/// before the next paint, so nothing flickers.
export function fitRibbon(ribbon: FitRibbon): FitResult {
  const groups = Array.from(ribbon.querySelectorAll(':scope > .ribbon-group'))
  ribbon.classList.remove(TIGHT_CLASS)
  for (const group of groups) group.classList.remove(ICONS_ONLY_CLASS)
  if (!overflows(ribbon)) return { tight: false, iconsOnly: 0 }
  ribbon.classList.add(TIGHT_CLASS)
  let iconsOnly = 0
  for (let i = groups.length - 1; i >= 0 && overflows(ribbon); i--) {
    groups[i]!.classList.add(ICONS_ONLY_CLASS)
    iconsOnly++
  }
  return { tight: true, iconsOnly }
}

/// Keep every band under `header` fitted: on resize (window, interface scale),
/// on tab switches and contextual tabs (subtree changes), on value text that
/// changes a trigger width (number format), on density changes (the header's
/// own class), and once the web fonts have loaded. Two observers, not one:
/// watching class changes across the subtree would catch fitRibbon's own
/// toggles and loop.
export function installRibbonFit(header: HTMLElement): () => void {
  let frame = 0
  const run = (): void => {
    frame = 0
    for (const ribbon of header.querySelectorAll<HTMLElement>('[data-ribbon-body]')) {
      fitRibbon(ribbon)
    }
  }
  const schedule = (): void => {
    if (!frame) frame = requestAnimationFrame(run)
  }
  const resize = new ResizeObserver(schedule)
  resize.observe(header)
  const content = new MutationObserver(schedule)
  content.observe(header, { childList: true, subtree: true, characterData: true })
  const density = new MutationObserver(schedule)
  density.observe(header, { attributes: true, attributeFilter: ['class'] })
  void document.fonts?.ready.then(schedule)
  schedule()
  return () => {
    resize.disconnect()
    content.disconnect()
    density.disconnect()
    if (frame) cancelAnimationFrame(frame)
  }
}
