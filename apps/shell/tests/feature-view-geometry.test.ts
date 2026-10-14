import { describe, expect, it } from 'vitest'
import { isAllowedViewUrl, viewHost } from '@genoffice/agent-core'
import { MIN_VIEW_SIZE, clampViewRect } from '../src/shared/feature-view-geometry'

describe('native view geometry', () => {
  it('rounds and clamps a rectangle into the content area', () => {
    expect(
      clampViewRect(
        { x: 10.4, y: 20.6, width: 400.5, height: 300.4 },
        { width: 1200, height: 800 },
      ),
    ).toEqual({ x: 10, y: 21, width: 401, height: 300 })
    // the right and bottom edges fold back into the window
    expect(
      clampViewRect({ x: 1000, y: 700, width: 600, height: 400 }, { width: 1200, height: 800 }),
    ).toEqual({ x: 1000, y: 700, width: 200, height: 100 })
    // a negative origin is not a place for a view
    expect(
      clampViewRect({ x: -50, y: -50, width: 400, height: 300 }, { width: 1200, height: 800 }),
    ).toEqual({ x: 0, y: 0, width: 400, height: 300 })
  })

  it('refuses a rectangle that is too small, or not a rectangle', () => {
    const area = { width: 1200, height: 800 }
    expect(clampViewRect({ x: 0, y: 0, width: MIN_VIEW_SIZE - 1, height: 400 }, area)).toBeNull()
    expect(clampViewRect({ x: 0, y: 0, width: 400, height: MIN_VIEW_SIZE - 1 }, area)).toBeNull()
    expect(clampViewRect({ x: 0, y: 0, width: 'wide', height: 400 }, area)).toBeNull()
    expect(clampViewRect(null, area)).toBeNull()
    expect(clampViewRect('rect', area)).toBeNull()
  })

  it('allows https anywhere and http only on the loopback', () => {
    expect(isAllowedViewUrl('https://example.com/a')).toBe(true)
    expect(isAllowedViewUrl('https://localhost:3000')).toBe(true)
    expect(isAllowedViewUrl('http://localhost:5173/app')).toBe(true)
    expect(isAllowedViewUrl('http://127.0.0.1:8080')).toBe(true)
    expect(isAllowedViewUrl('http://example.com')).toBe(false)
    expect(isAllowedViewUrl('file:///etc/passwd')).toBe(false)
    expect(isAllowedViewUrl('data:text/html,<h1>hi</h1>')).toBe(false)
    expect(isAllowedViewUrl('javascript:alert(1)')).toBe(false)
    expect(isAllowedViewUrl('about:blank')).toBe(false)
    expect(isAllowedViewUrl('not a url')).toBe(false)
  })

  it('names the host for the frame', () => {
    expect(viewHost('https://news.ycombinator.com/newest')).toBe('news.ycombinator.com')
    expect(viewHost('nonsense')).toBe('nonsense')
  })
})
