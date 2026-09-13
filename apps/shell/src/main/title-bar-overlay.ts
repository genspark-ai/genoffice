import type { TitleBarOverlay } from 'electron'

/// Windows/Linux draw the caption buttons over the tab strip (titleBarStyle
/// 'hidden'); the overlay must match the strip's --tabstrip-bg / text tokens
/// in renderer/src/tabbar.css or the buttons sit on a visibly different band.
export const TAB_STRIP_HEIGHT = 40

export function tabStripOverlay(dark: boolean): TitleBarOverlay {
  return dark
    ? { color: '#2a2a2a', symbolColor: '#e4e4e4', height: TAB_STRIP_HEIGHT }
    : { color: '#ebebeb', symbolColor: '#454746', height: TAB_STRIP_HEIGHT }
}
