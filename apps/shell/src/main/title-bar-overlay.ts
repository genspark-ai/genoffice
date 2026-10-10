import type { TitleBarOverlay } from 'electron'

import { tabStripHeight } from '../shared/tab-drag-geometry'

/// Windows/Linux draw the caption buttons over the tab strip (titleBarStyle
/// 'hidden'); the overlay must match the strip's --tabstrip-bg / text tokens
/// in renderer/src/tabbar.css or the buttons sit on a visibly different band.
///
/// Its height has to follow the interface scale too (#1913): the strip is
/// `zoom`ed, so a 150% strip is 60px tall and a 40px overlay band would leave
/// the caption buttons floating over the document.

export function tabStripOverlay(dark: boolean, uiScale = 1): TitleBarOverlay {
  const height = tabStripHeight(uiScale)
  return dark
    ? { color: '#2a2a2a', symbolColor: '#e4e4e4', height }
    : { color: '#ebebeb', symbolColor: '#454746', height }
}
