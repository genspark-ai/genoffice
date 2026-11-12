/** In-window application menu for Windows (the native menu bar is hidden so the
 *  tab strip, menus and window controls read as one transparent title bar). */
export const CHROME_CHANNELS = {
  menuLabels: 'chrome:menu-labels',
  menuPopup: 'chrome:menu-popup',
  menuChanged: 'chrome:menu-changed',
} as const

export interface ChromeApi {
  /** top-level application menu labels, in order; empty when the OS draws the menu (macOS) */
  menuLabels(): Promise<string[]>
  /** open the submenu at `index` below the given window-relative point */
  popupMenu(index: number, x: number, y: number): Promise<void>
  /** the active tab switched the application menu */
  onMenuChanged(handler: (labels: string[]) => void): () => void
}
