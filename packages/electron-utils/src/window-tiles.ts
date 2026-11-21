export interface TileRect {
  x: number
  y: number
  width: number
  height: number
}

export const TILE_MIN_ROW_HEIGHT = 320

/** Rows at full width while they stay workable, then shared rows; a short last row is stretched */
export function tileRects(
  area: TileRect,
  count: number,
  minRowHeight = TILE_MIN_ROW_HEIGHT,
): TileRect[] {
  if (count <= 0) return []
  const maxRows = Math.max(1, Math.floor(area.height / minRowHeight))
  const cols = Math.ceil(count / maxRows)
  const rows = Math.ceil(count / cols)
  const edge = (size: number, i: number, n: number) => Math.floor((size * i) / n)
  const out: TileRect[] = []
  for (let i = 0; i < count; i++) {
    const row = Math.floor(i / cols)
    const inRow = row === rows - 1 ? count - row * cols : cols
    const col = i - row * cols
    const left = edge(area.width, col, inRow)
    const top = edge(area.height, row, rows)
    out.push({
      x: area.x + left,
      y: area.y + top,
      width: edge(area.width, col + 1, inRow) - left,
      height: edge(area.height, row + 1, rows) - top,
    })
  }
  return out
}

/** the part of Electron's BrowserWindow that tiling needs */
export interface TileWindow {
  isDestroyed(): boolean
  isMinimized(): boolean
  restore(): void
  isMaximized(): boolean
  unmaximize(): void
  isFullScreen(): boolean
  setFullScreen(flag: boolean): void
  once(event: 'leave-full-screen', listener: () => void): unknown
  getMinimumSize(): number[]
  setMinimumSize(width: number, height: number): void
  setBounds(bounds: TileRect): void
}

function placeWindow(win: TileWindow, bounds: TileRect, beforePlace?: () => void): void {
  if (win.isDestroyed()) return
  if (win.isMinimized()) win.restore()
  const place = () => {
    if (win.isDestroyed()) return
    beforePlace?.()
    // leaving fullscreen brings back the maximized state, which ignores setBounds
    if (win.isMaximized()) win.unmaximize()
    const [minW = 0, minH = 0] = win.getMinimumSize()
    if (bounds.width < minW || bounds.height < minH)
      win.setMinimumSize(Math.min(minW, bounds.width), Math.min(minH, bounds.height))
    win.setBounds(bounds)
  }
  // leaving fullscreen is asynchronous and restores the old bounds when it completes
  if (win.isFullScreen()) {
    win.once('leave-full-screen', place)
    win.setFullScreen(false)
  } else {
    place()
  }
}

/** beforePlace runs right before each window moves, also one that had to leave fullscreen first */
export function tileWindows(
  windows: readonly TileWindow[],
  area: TileRect,
  beforePlace?: () => void,
): void {
  const rects = tileRects(area, windows.length)
  windows.forEach((win, i) => {
    const rect = rects[i]
    if (rect) placeWindow(win, rect, beforePlace)
  })
}
