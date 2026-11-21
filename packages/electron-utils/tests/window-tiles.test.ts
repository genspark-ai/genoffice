import { describe, expect, it } from 'vitest'
import { tileRects, tileWindows, type TileWindow } from '../src/window-tiles'

const AREA = { x: 0, y: 25, width: 1440, height: 875 }
const covered = (rects: ReturnType<typeof tileRects>) =>
  rects.reduce((sum, r) => sum + r.width * r.height, 0)
const overlap = (a: typeof AREA, b: typeof AREA) =>
  a.x < b.x + b.width && b.x < a.x + a.width && a.y < b.y + b.height && b.y < a.y + a.height

describe('tileRects', () => {
  it('one window takes the whole work area', () => {
    expect(tileRects(AREA, 1)).toEqual([AREA])
  })

  it('two windows stack top to bottom at full width', () => {
    expect(tileRects(AREA, 2)).toEqual([
      { x: 0, y: 25, width: 1440, height: 437 },
      { x: 0, y: 462, width: 1440, height: 438 },
    ])
  })

  it('more windows than workable rows share rows, a short last row stretched', () => {
    expect(tileRects(AREA, 3)).toEqual([
      { x: 0, y: 25, width: 720, height: 437 },
      { x: 720, y: 25, width: 720, height: 437 },
      { x: 0, y: 462, width: 1440, height: 438 },
    ])
  })

  it('covers the area exactly with no overlap, whatever the count', () => {
    for (const area of [AREA, { x: -1920, y: 0, width: 1919, height: 1055 }]) {
      for (let n = 1; n <= 12; n++) {
        const rects = tileRects(area, n)
        expect(rects).toHaveLength(n)
        expect(covered(rects), `${n}`).toBe(area.width * area.height)
        for (let i = 0; i < n; i++)
          for (let j = i + 1; j < n; j++)
            expect(overlap(rects[i], rects[j]), `${i}/${j}`).toBe(false)
      }
    }
  })

  it('nothing to arrange', () => {
    expect(tileRects(AREA, 0)).toEqual([])
  })

  it('a fullscreen window is tiled once it has left fullscreen, and a small tile lowers the minimum size', () => {
    const calls: string[] = []
    let leave = () => {}
    const win = (fullScreen: boolean): TileWindow => ({
      isDestroyed: () => false,
      isMinimized: () => true,
      restore: () => calls.push('restore'),
      isMaximized: () => true,
      unmaximize: () => calls.push('unmaximize'),
      isFullScreen: () => fullScreen,
      setFullScreen: (flag) => calls.push(`fullscreen:${flag}`),
      once: (_event, listener) => (leave = listener),
      getMinimumSize: () => [720, 550],
      setMinimumSize: (w, h) => calls.push(`min:${w}x${h}`),
      setBounds: (b) => calls.push(`bounds:${b.y}+${b.height}`),
    })
    tileWindows([win(false), win(true)], AREA, () => calls.push('before'))
    expect(calls).toEqual([
      'restore',
      'before',
      'unmaximize',
      'min:720x437',
      'bounds:25+437',
      'restore',
      'fullscreen:false',
    ])
    leave()
    // maximized again once out of fullscreen: un-maximized before the bounds are set
    expect(calls.slice(-4)).toEqual(['before', 'unmaximize', 'min:720x438', 'bounds:462+438'])
  })
})
