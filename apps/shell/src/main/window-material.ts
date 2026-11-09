import { release } from 'node:os'

/**
 * Windows 11 22H2 (build 22621) is the first release where Electron's
 * `backgroundMaterial` takes effect; earlier builds ignore it and would paint
 * an opaque window behind a transparent page.
 */
const WIN11_22H2_BUILD = 22621

/** True when the shell window is created with a wallpaper-tinted OS material on Windows. */
export function winMaterialSupported(): boolean {
  if (process.platform !== 'win32') return false
  const build = Number(release().split('.')[2])
  return Number.isFinite(build) && build >= WIN11_22H2_BUILD
}

/** Windows caption-button overlay drawn over our own 44px title bar (see tab-bar in tabbar.css). */
export const TITLE_BAR_HEIGHT = 44

export function winTitleBarOverlay(dark: boolean): {
  color: string
  symbolColor: string
  height: number
} {
  return {
    // Mica: transparent so the buttons sit on the same glass as the tab strip;
    // otherwise match the strip's solid fallback color
    color: winMaterialSupported() ? '#00000000' : dark ? '#21201f' : '#ebebeb',
    symbolColor: dark ? '#f4f4f4' : '#454746',
    height: TITLE_BAR_HEIGHT,
  }
}
