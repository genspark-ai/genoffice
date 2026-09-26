import { describe, expect, it } from 'vitest'
import type { PictureRenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import { auditSlideFindings, auditSlideLayout, formatAudit } from '../src/slides/layout-audit'

const W = 1280
const H = 720

function placedBox(x: number, y: number, w: number, h: number) {
  return {
    x,
    y,
    w,
    h,
    rotationDeg: 0,
    flipH: false,
    flipV: false,
    centerX: x + w / 2,
    centerY: y + h / 2,
  }
}

function textNode(
  id: string,
  x: number,
  y: number,
  w: number,
  h: number,
  opts: { text?: string; contentHeight?: number; runWidth?: number; wrap?: boolean } = {},
): ShapeRenderNode {
  const text = opts.text ?? 'Quarterly wins across regions'
  const contentHeight = opts.contentHeight ?? h
  const runWidth = opts.runWidth ?? Math.min(120, w)
  return {
    id,
    type: 'text',
    box: placedBox(x, y, w, h),
    sourceId: id,
    fill: { kind: 'none' },
    text: {
      lines: [
        {
          runs: [
            {
              text,
              x: 0,
              baselineY: 20,
              fontFamily: 'Arial',
              fontSizePx: 16,
              color: '#111111',
              bold: false,
              italic: false,
              underline: false,
              widthPx: runWidth,
            },
          ],
          top: 0,
          height: 20,
        },
      ],
      insets: { l: 0, t: 0, r: 0, b: 0 },
      anchor: 'top',
      fontScale: 1,
      contentHeight,
      wrap: opts.wrap ?? true,
    },
  }
}

function pictureNode(
  id: string,
  x: number,
  y: number,
  w: number,
  h: number,
  extra: Partial<PictureRenderNode> = {},
): PictureRenderNode {
  return { id, type: 'picture', box: placedBox(x, y, w, h), sourceId: id, ...extra }
}

function slide(nodes: RenderSlide['nodes']): RenderSlide {
  return {
    widthPx: W,
    heightPx: H,
    scale: 1,
    background: { kind: 'solid', color: '#FFFFFF' },
    nodes,
  }
}

describe('auditSlideLayout clean pass', () => {
  it('passes widely separated text boxes', () => {
    const s = slide([textNode('a', 80, 80, 400, 100), textNode('b', 700, 400, 400, 100)])
    expect(auditSlideFindings(s)).toEqual([])
    expect(auditSlideLayout(s)).toEqual([])
    expect(formatAudit([])).toContain('Passed')
  })

  it('tolerates edge and overflow noise below the thresholds', () => {
    const s = slide([
      textNode('a', -4, 80, 400, 100, { contentHeight: 102 }),
      textNode('b', 700, 400, 300, 100, { runWidth: 303 }),
    ])
    expect(auditSlideFindings(s)).toEqual([])
  })
})

describe('auditSlideLayout out of bounds', () => {
  it('flags boxes past each canvas edge', () => {
    const s = slide([
      textNode('left', -30, 100, 200, 80),
      textNode('top', 100, -25, 200, 80),
      textNode('right', W - 100, 100, 200, 80),
      textNode('bottom', 100, H - 40, 200, 80),
    ])
    const findings = auditSlideFindings(s)
    expect(findings.map((f) => f.code)).toEqual([
      'out_of_bounds',
      'out_of_bounds',
      'out_of_bounds',
      'out_of_bounds',
    ])
    const messages = findings.map((f) => f.message)
    expect(messages[0]).toContain('past the left edge')
    expect(messages[1]).toContain('past the top edge')
    expect(messages[2]).toContain('past the right edge')
    expect(messages[3]).toContain('past the bottom edge')
    for (const f of findings) {
      expect(f.level).toBe('error')
      expect(f.suggest?.op).toBe('setTransform')
    }
  })
})

describe('auditSlideLayout off slide', () => {
  it('reports a box with no canvas intersection once, as off_slide, with a clamp suggestion', () => {
    const s = slide([
      textNode('gone', W + 50, 100, 200, 80),
      textNode('edge', W - 100, 300, 200, 80),
      textNode('above', 100, -200, 200, 80),
    ])
    const findings = auditSlideFindings(s)
    expect(findings.map((f) => [f.el, f.code])).toEqual([
      ['gone', 'off_slide'],
      ['edge', 'out_of_bounds'],
      ['above', 'off_slide'],
    ])
    const gone = findings[0]!
    expect(gone.level).toBe('error')
    expect(gone.message).toContain('entirely outside the slide')
    expect(gone.message).toContain('past the right edge')
    expect(gone.suggest).toEqual({
      op: 'setTransform',
      target: { el: 'gone' },
      box: { x: (W - 200) * 9525, y: 100 * 9525, cx: 200 * 9525, cy: 80 * 9525 },
    })
    expect(findings[2]!.suggest?.box).toMatchObject({ x: 100 * 9525, y: 0 })
  })

  it('reports a thin line just past the edge as off_slide despite the edge tolerance', () => {
    const s = slide([textNode('rule', 100, H + 2, 400, 4)])
    const findings = auditSlideFindings(s)
    expect(findings.map((f) => f.code)).toEqual(['off_slide'])
    expect(findings[0]!.message).toContain('6px past the bottom edge')
    expect(findings[0]!.suggest?.box).toMatchObject({ y: (H - 4) * 9525 })
  })

  it('keeps a box touching the edge from outside as out_of_bounds', () => {
    const s = slide([textNode('touch', W, 100, 200, 80)])
    expect(auditSlideFindings(s).map((f) => f.code)).toEqual(['out_of_bounds'])
  })
})

describe('auditSlideLayout picture distortion', () => {
  const sizes: Record<string, { w: number; h: number }> = {
    stretched: { w: 800, h: 400 },
    cropped: { w: 800, h: 400 },
    close: { w: 800, h: 400 },
    poster: { w: 800, h: 400 },
  }
  const opts = { pictureSize: (id: string) => sizes[id] }

  it('flags a box whose aspect strays from the source and shrinks the longer side, centred', () => {
    const s = slide([pictureNode('stretched', 100, 100, 400, 400)])
    const findings = auditSlideFindings(s, undefined, opts)
    expect(findings).toHaveLength(1)
    expect(findings[0]).toMatchObject({
      code: 'picture_distorted',
      level: 'warning',
      el: 'stretched',
      expected_ratio: 2,
      actual_ratio: 1,
      distortion_pct: 50,
    })
    expect(findings[0]!.message).toContain('Picture distorted')
    expect(findings[0]!.suggest).toEqual({
      op: 'setTransform',
      target: { el: 'stretched' },
      box: { x: 100 * 9525, y: 200 * 9525, cx: 400 * 9525, cy: 200 * 9525 },
    })
  })

  it('shrinks a square logo in a banner box to the banner height instead of growing it', () => {
    const s = slide([pictureNode('logo', 100, 100, 600, 100)])
    const findings = auditSlideFindings(s, undefined, { pictureSize: () => ({ w: 256, h: 256 }) })
    expect(findings.map((f) => f.code)).toEqual(['picture_distorted'])
    expect(findings[0]!.suggest!.box).toEqual({
      x: 350 * 9525,
      y: 100 * 9525,
      cx: 100 * 9525,
      cy: 100 * 9525,
    })
  })

  it('grows instead when shrinking would leave the picture under 24px', () => {
    const s = slide([pictureNode('strip', 100, 300, 600, 10)])
    const findings = auditSlideFindings(s, undefined, { pictureSize: () => ({ w: 256, h: 256 }) })
    expect(findings[0]!.suggest!.box).toEqual({
      x: 100 * 9525,
      y: 5 * 9525,
      cx: 600 * 9525,
      cy: 600 * 9525,
    })
  })

  it('compares against the cropped source region and tolerates 5%', () => {
    const s = slide([
      pictureNode('cropped', 100, 100, 300, 300, { srcRect: { l: 0.25, t: 0, r: 0.25, b: 0 } }),
      pictureNode('close', 500, 100, 400, 196),
    ])
    expect(auditSlideFindings(s, undefined, opts)).toEqual([])
  })

  it('skips poster frames, unknown sizes and audits without a size lookup', () => {
    const s = slide([
      pictureNode('poster', 100, 100, 400, 400, { media: 'video' }),
      pictureNode('unknown', 600, 100, 400, 400),
    ])
    expect(auditSlideFindings(s, undefined, opts)).toEqual([])
    expect(auditSlideFindings(slide([pictureNode('stretched', 100, 100, 400, 400)]))).toEqual([])
  })

  it('scales a corrected picture down when the true aspect would not fit the canvas', () => {
    const s = slide([pictureNode('square', 0, 0, W, 1400)])
    const findings = auditSlideFindings(s, undefined, {
      pictureSize: () => ({ w: 400, h: 400 }),
    })
    expect(findings.map((f) => f.code)).toEqual(['out_of_bounds', 'picture_distorted'])
    expect(findings[1]!.suggest!.box).toEqual({
      x: ((W - H) / 2) * 9525,
      y: 0,
      cx: H * 9525,
      cy: H * 9525,
    })
    expect(findings[0]!.suggest).toEqual(findings[1]!.suggest)
  })
})

describe('auditSlideLayout text overflow', () => {
  it('flags content taller than the box with a taller suggestion', () => {
    const s = slide([textNode('tall', 100, 100, 400, 40, { contentHeight: 100 })])
    const findings = auditSlideFindings(s)
    expect(findings).toHaveLength(1)
    expect(findings[0]).toMatchObject({ code: 'text_overflow', level: 'error', el: 'tall' })
    expect(findings[0]!.message).toContain('exceeds the box height by 60px')
    expect(findings[0]!.overflowPx).toBe(60)
    expect(findings[0]!.suggest?.op).toBe('setTransform')
  })

  it('flags a line wider than the box as a width warning', () => {
    const s = slide([textNode('wide', 100, 100, 200, 80, { runWidth: 300 })])
    const findings = auditSlideFindings(s)
    expect(findings).toHaveLength(1)
    expect(findings[0]).toMatchObject({ code: 'text_overflow_width', level: 'warning' })
    expect(findings[0]!.message).toContain('exceeds the box width by 100px')
    expect(findings[0]!.suggest?.op).toBe('setTransform')
  })
})

describe('auditSlideLayout overlap', () => {
  it('flags two intersecting text boxes', () => {
    const s = slide([textNode('a', 100, 100, 200, 100), textNode('b', 150, 120, 200, 100)])
    const findings = auditSlideFindings(s)
    expect(findings).toHaveLength(1)
    expect(findings[0]).toMatchObject({ code: 'overlap', level: 'warning', el: 'a' })
    expect(findings[0]!.els).toEqual(['a', 'b'])
    expect(findings[0]!.message).toContain('intersect by 150')
    expect(auditSlideLayout(s)).toHaveLength(1)
    expect(formatAudit(auditSlideLayout(s))).toContain('Found 1 issue')
  })

  it('flags text on image but ignores image on image', () => {
    const textOnImage = slide([
      textNode('caption', 100, 100, 200, 100),
      pictureNode('photo', 150, 120, 200, 100),
    ])
    expect(auditSlideFindings(textOnImage).map((f) => f.code)).toEqual(['overlap'])

    const imageOnImage = slide([
      pictureNode('one', 100, 100, 200, 100),
      pictureNode('two', 150, 120, 200, 100),
    ])
    expect(auditSlideFindings(imageOnImage)).toEqual([])
  })

  it('ignores decoration nodes and full-canvas background blocks', () => {
    const decorated = textNode('deco', 150, 120, 200, 100)
    const s = slide([textNode('main', 100, 100, 200, 100), { ...decorated, decoration: true }])
    expect(auditSlideFindings(s)).toEqual([])

    const backdrop = textNode('backdrop', 0, 0, 1200, 650, { text: 'Backdrop block' })
    const small = textNode('small', 100, 100, 200, 100)
    expect(auditSlideFindings(slide([backdrop, small]))).toEqual([])
  })
})
