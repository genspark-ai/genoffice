import type { TitleBarOverlay } from 'electron'

import { tabStripHeight } from '../shared/tab-drag-geometry'

/// Windows/Linux draw the caption buttons over the tab strip (titleBarStyle
/// 'hidden'); the overlay must match the strip's --tabstrip-bg / text tokens
/// in renderer/src/tabbar.css or the buttons sit on a visibly different band.
/// Over the Mica material the strip is see-through, so the band is too; the
/// dark fallback is the strip's --wall color.
///
/// Its height has to follow the interface scale too (genoffice#1913): the strip is
/// `zoom`ed, so a 150% strip is 60px tall and a 40px overlay band would leave
/// the caption buttons floating over the document.

export function tabStripOverlay(dark: boolean, uiScale = 1, material = false): TitleBarOverlay {
  const height = tabStripHeight(uiScale)
  const symbolColor = dark ? '#f4f4f4' : '#454746'
  if (material) return { color: '#00000000', symbolColor, height }
  return { color: dark ? '#21201f' : '#ebebeb', symbolColor, height }
}
