import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'
import { describe, expect, it } from 'vitest'
import { safeExternalUrl } from '@genoffice/electron-utils/safe-external-url'
import type { LinkTargetOp } from '../src/shared/ipc'
import { DECK_LINK_PROTOCOLS, decodeLinkTarget } from '../src/shared/run-link'
import {
  collectExportPdfLinks,
  collectSlideLinkRects,
  exportLinkHref,
} from '../src/renderer/export-links'

const url = (u: string): LinkTargetOp => ({ kind: 'url', url: u })

function shape(sourceId: string, box: { x: number; y: number; w: number; h: number }): RenderNode {
  return { type: 'shape', sourceId, box } as unknown as RenderNode
}

function textShape(
  sourceId: string,
  box: { x: number; y: number; w: number; h: number },
  lines: unknown[],
): RenderNode {
  return {
    type: 'text',
    sourceId,
    box,
    text: { lines, insets: { l: 10, t: 5, r: 10, b: 5 } },
  } as unknown as RenderNode
}

function slide(nodes: RenderNode[], widthPx = 960, heightPx = 540): RenderSlide {
  return { widthPx, heightPx, scale: 1, background: { kind: 'none' }, nodes } as RenderSlide
}

describe('collectSlideLinkRects', () => {
  it('emits a rect over a linked element and skips unlinked ones', () => {
    const s = slide([
      shape('a', { x: 50, y: 60, w: 200, h: 80 }),
      shape('b', { x: 0, y: 0, w: 10, h: 10 }),
    ])
    const rects = collectSlideLinkRects(s, new Map([['a', url('https://x.test/')]]), new Map())
    expect(rects).toEqual([{ x: 50, y: 60, w: 200, h: 80, target: url('https://x.test/') }])
  })

  it('offsets group children by the group box and clips them to it', () => {
    const child = shape('c', { x: 90, y: 10, w: 40, h: 20 }) // group-local, overflows the 100px-wide group
    const group = {
      type: 'group',
      sourceId: 'g',
      box: { x: 100, y: 100, w: 100, h: 100 },
      children: [child],
    } as unknown as RenderNode
    const rects = collectSlideLinkRects(slide([group]), new Map([['c', url('u:c')]]), new Map())
    expect(rects).toEqual([{ x: 190, y: 110, w: 10, h: 20, target: url('u:c') }])
  })

  it('clips element rects to the page', () => {
    const s = slide([shape('a', { x: 900, y: -10, w: 200, h: 40 })])
    const rects = collectSlideLinkRects(s, new Map([['a', url('u')]]), new Map())
    expect(rects).toEqual([{ x: 900, y: 0, w: 60, h: 30, target: url('u') }])
  })

  it('maps run links through the glyph layout with paragraph counting, after the element rect', () => {
    const lines = [
      // paragraph 0: one line, run 0 linked
      {
        top: 0,
        height: 20,
        runs: [
          { x: 0, widthPx: 40, srcRunIdx: 0 },
          { x: 40, widthPx: 30, srcRunIdx: 1 },
        ],
      },
      // wrap continuation of paragraph 0
      { top: 20, height: 20, paraStart: false, runs: [{ x: 0, widthPx: 25, srcRunIdx: 1 }] },
      // paragraph 1: run 0 linked
      { top: 40, height: 20, runs: [{ x: 5, widthPx: 50, srcRunIdx: 0 }] },
    ]
    const s = slide([textShape('t', { x: 100, y: 200, w: 300, h: 100 }, lines)])
    const runLinks = new Map<string, LinkTargetOp>([
      ['t:0:0', url('u:first')],
      ['t:1:0', url('u:second')],
    ])
    const rects = collectSlideLinkRects(s, new Map([['t', url('u:whole')]]), runLinks)
    expect(rects).toEqual([
      { x: 100, y: 200, w: 300, h: 100, target: url('u:whole') },
      { x: 110, y: 205, w: 40, h: 20, target: url('u:first') },
      { x: 115, y: 245, w: 50, h: 20, target: url('u:second') },
    ])
  })

  it('clips run rects to the element box like the show hit test', () => {
    const lines = [
      // overflows the 100px-wide box by 50px on the right
      { top: 0, height: 20, runs: [{ x: 60, widthPx: 80, srcRunIdx: 0 }] },
      // laid out entirely below the 30px-tall box
      { top: 40, height: 20, runs: [{ x: 0, widthPx: 20, srcRunIdx: 1 }] },
    ]
    const s = slide([textShape('t', { x: 100, y: 100, w: 100, h: 30 }, lines)])
    const runLinks = new Map<string, LinkTargetOp>([
      ['t:0:0', url('u:a')],
      ['t:0:1', url('u:b')],
    ])
    expect(collectSlideLinkRects(s, new Map(), runLinks)).toEqual([
      { x: 170, y: 105, w: 30, h: 20, target: url('u:a') },
    ])
  })
})

describe('exportLinkHref', () => {
  const pages = new Map([
    [0, 0],
    [2, 1],
  ]) // model slide 1 hidden
  it('passes URLs through and maps slide jumps to visible page anchors', () => {
    expect(exportLinkHref(url('https://x.test/'), 0, pages, 2)).toBe('https://x.test/')
    expect(exportLinkHref({ kind: 'slide', slideIndex: 2 }, 0, pages, 2)).toBe('#pg2')
    // jump to a hidden page is dropped, matching the show's skip semantics
    expect(exportLinkHref({ kind: 'slide', slideIndex: 1 }, 0, pages, 2)).toBeNull()
  })
  it('maps navigation actions relative to the current page and drops show-only ones', () => {
    expect(exportLinkHref({ kind: 'action', action: 'nextslide' }, 0, pages, 2)).toBe('#pg2')
    expect(exportLinkHref({ kind: 'action', action: 'nextslide' }, 1, pages, 2)).toBeNull()
    expect(exportLinkHref({ kind: 'action', action: 'previousslide' }, 1, pages, 2)).toBe('#pg1')
    expect(exportLinkHref({ kind: 'action', action: 'previousslide' }, 0, pages, 2)).toBeNull()
    expect(exportLinkHref({ kind: 'action', action: 'firstslide' }, 1, pages, 2)).toBe('#pg1')
    expect(exportLinkHref({ kind: 'action', action: 'lastslide' }, 0, pages, 2)).toBe('#pg2')
    expect(exportLinkHref({ kind: 'action', action: 'endshow' }, 0, pages, 2)).toBeNull()
    expect(exportLinkHref({ kind: 'action', action: 'lastslideviewed' }, 0, pages, 2)).toBeNull()
  })

  it('drops dangerous URL schemes from export hrefs', () => {
    expect(exportLinkHref(url('javascript:alert(1)'), 0, pages, 2)).toBeNull()
    expect(exportLinkHref(url('data:text/html,hi'), 0, pages, 2)).toBeNull()
    expect(exportLinkHref(url('file:///etc/passwd'), 0, pages, 2)).toBeNull()
    expect(exportLinkHref(url('mailto:a@b.test'), 0, pages, 2)).toBe('mailto:a@b.test')
    expect(exportLinkHref(url('https://x.test/'), 0, pages, 2)).toBe('https://x.test/')
  })
})

describe('set-link gate (DECK_LINK_PROTOCOLS)', () => {
  it('rejects javascript:/data: and keeps http(s)/mailto', () => {
    const gate = (u: string) => safeExternalUrl(u, { allowedProtocols: DECK_LINK_PROTOCOLS })
    expect(gate('javascript:alert(1)')).toBeNull()
    expect(gate('data:text/html,hi')).toBeNull()
    expect(gate('https://x.test/')).toBe('https://x.test/')
    expect(gate('mailto:a@b.test')).toBe('mailto:a@b.test')
  })
})

describe('decodeLinkTarget', () => {
  it('never drops a file-authored link, whatever its scheme', () => {
    for (const u of ['../report.pdf', 'www.x.com', 'tel:+123', 'ftp://h/f', 'javascript:1']) {
      expect(decodeLinkTarget(u)).toEqual({ kind: 'url', url: u })
    }
    expect(decodeLinkTarget('slide:2')).toEqual({ kind: 'slide', slideIndex: 2 })
  })
})

describe('collectExportPdfLinks', () => {
  it('produces per-page fraction rects with resolved hrefs', () => {
    const s = slide([shape('a', { x: 96, y: 54, w: 480, h: 270 })])
    const out = collectExportPdfLinks(
      [s],
      [[{ sourceId: 'a', target: url('https://x.test/') }]],
      [[]],
      new Map([[0, 0]]),
    )
    expect(out).toEqual([[{ x: 0.1, y: 0.1, w: 0.5, h: 0.5, href: 'https://x.test/' }]])
  })

  it('drops links without a PDF meaning and pages without links stay empty', () => {
    const s = slide([shape('a', { x: 0, y: 0, w: 10, h: 10 })])
    const out = collectExportPdfLinks(
      [s, s],
      [[{ sourceId: 'a', target: { kind: 'action', action: 'endshow' } }], []],
      [[], []],
      new Map([
        [0, 0],
        [1, 1],
      ]),
    )
    expect(out).toEqual([[], []])
  })
})
