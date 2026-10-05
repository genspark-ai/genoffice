/**
 * The template chrome skeleton (#1821): extraction from a deck's render tree
 * (title box / brand-image slot / accent shapes / background per page role)
 * and its prompt formatting. Pure logic — fixtures are loose render nodes.
 */
// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'

import type { PlacedBox, RenderSlide } from '@genoffice/pptx-render'

import {
  extractLayoutSkeleton,
  formatSkeletonForPrompt,
  skeletonRole,
} from '../src/renderer/ai/layout-skeleton'

const box = (x: number, y: number, w: number, h: number): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})

/** A text shape whose largest run is `sizePx` tall */
const textNode = (x: number, y: number, w: number, h: number, text: string, sizePx: number) => ({
  id: `t${x}-${y}`,
  sourceId: `t${x}-${y}`,
  type: 'text',
  box: box(x, y, w, h),
  fill: { kind: 'none' },
  text: { lines: [{ runs: [{ text, fontSizePx: sizePx, color: '#F8FAFC', isBullet: false }] }] },
})

const rectNode = (x: number, y: number, w: number, h: number, fill: string) => ({
  id: `s${x}-${y}`,
  sourceId: `s${x}-${y}`,
  type: 'shape',
  box: box(x, y, w, h),
  fill: { kind: 'solid', color: fill },
})

const picNode = (x: number, y: number, w: number, h: number) => ({
  id: `p${x}-${y}`,
  sourceId: `p${x}-${y}`,
  type: 'picture',
  box: box(x, y, w, h),
})

const slide = (nodes: unknown[]): RenderSlide =>
  ({
    widthPx: 1280,
    heightPx: 720,
    scale: 1,
    background: { kind: 'solid', color: '#0F172A' },
    nodes,
  }) as unknown as RenderSlide

describe('skeletonRole', () => {
  it('maps first/last/middle pages to cover/content/closing', () => {
    expect(skeletonRole(0, 5)).toBe('cover')
    expect(skeletonRole(2, 5)).toBe('content')
    expect(skeletonRole(4, 5)).toBe('closing')
  })

  it('keeps a two-page deck cover+content (no closing role)', () => {
    expect(skeletonRole(1, 2)).toBe('content')
  })
})

describe('extractLayoutSkeleton', () => {
  it('pulls the recurring chrome out of a generated deck, jitter included', () => {
    const cover = slide([
      textNode(320, 220, 640, 140, 'Q3 Review', 72),
      rectNode(64, 640, 200, 8, '#2563EB'),
    ])
    // Three content pages: same title box, same logo box (±4px), same accent bar
    const content = (i: number) =>
      slide([
        textNode(64, 48, 900, 96, `Section ${i}`, 53),
        picNode(1120 + [4, -4, 0][i]!, 40, 96, 96),
        rectNode(64, 648, 1152, 8, '#2563EB'),
        rectNode(200 + i * 40, 300, 400, 200, '#1E293B'), // unique per page — not chrome
      ])
    const closing = slide([textNode(400, 300, 480, 80, 'Thanks', 48)])
    const skeleton = extractLayoutSkeleton([cover, content(0), content(1), content(2), closing])!

    expect(skeleton.canvas).toEqual({ w: 1280, h: 720 })
    expect(skeleton.content?.title).toMatchObject({ x: 64, y: 48, w: 900, h: 96, sizePt: 40 })
    expect(skeleton.content?.brandImage).toMatchObject({ x: 1120, y: 40, w: 96, h: 96, pages: 3 })
    expect(skeleton.content?.accents).toEqual([{ x: 64, y: 648, w: 1152, h: 8, fill: '#2563EB' }])
    expect(skeleton.content?.background).toBe('#0F172A')
    expect(skeleton.cover?.title).toMatchObject({ x: 320, y: 220, w: 640, h: 140, sizePt: 54 })
    expect(skeleton.closing?.title).toMatchObject({ x: 400, y: 300, w: 480, h: 80 })
    // One-off pages carry no brand slot; the cover's lone accent is not chrome
    expect(skeleton.cover?.brandImage).toBeUndefined()
    expect(skeleton.cover?.accents).toBeUndefined()
  })

  it('resolves group children to absolute boxes', () => {
    const grouped = slide([
      {
        id: 'g1',
        sourceId: 'g1',
        type: 'group',
        box: box(1000, 40, 200, 100),
        children: [picNode(0, 0, 80, 80), picNode(90, 0, 80, 80)],
      },
      {
        id: 'g2',
        sourceId: 'g2',
        type: 'group',
        box: box(1000, 40, 200, 100),
        children: [picNode(0, 0, 80, 80), picNode(90, 0, 80, 80)],
      },
    ])
    const skeleton = extractLayoutSkeleton([grouped, grouped])!
    // Four pictures per slide, two recurring boxes — the most frequent cluster wins
    expect(skeleton.cover?.brandImage).toMatchObject({ x: 1000, y: 40, w: 80, h: 80, pages: 2 })
  })

  it('returns null when nothing recurs', () => {
    expect(extractLayoutSkeleton([])).toBeNull()
    expect(extractLayoutSkeleton([slide([])])).toBeNull()
  })

  it('skips hidden slides', () => {
    const hidden = { ...slide([picNode(10, 10, 50, 50)]), hidden: true } as unknown as RenderSlide
    const keep = slide([picNode(10, 10, 50, 50)])
    const skeleton = extractLayoutSkeleton([hidden, keep])!
    expect(skeleton.cover?.brandImage).toBeUndefined() // 1 visible page, no recurrence
  })
})

describe('formatSkeletonForPrompt', () => {
  it('renders the pinned chrome as explicit lines', () => {
    // Four pages: cover + two content pages + closing, so the content role has recurrence
    const page = (t: string) => slide([textNode(64, 48, 900, 96, t, 53), picNode(1120, 40, 96, 96)])
    const skeleton = extractLayoutSkeleton([page('A'), page('B'), page('C'), page('D')])!
    const text = formatSkeletonForPrompt(skeleton, 'content')
    expect(text).toContain('Template chrome for content pages')
    expect(text).toContain('- title box: x=64 y=48 w=900 h=96, font 40pt, color #F8FAFC')
    expect(text).toContain('- brand image slot: x=1120 y=40 w=96 h=96')
    expect(text).toContain('the same spot on every page')
  })

  it('returns empty text for a role the template does not pin', () => {
    const skeleton = extractLayoutSkeleton([slide([textNode(0, 0, 100, 40, 'A', 40)])])!
    expect(formatSkeletonForPrompt(skeleton, 'closing')).toBe('')
  })
})
