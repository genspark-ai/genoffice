/**
 * Insert-gallery preview regression: every gallery entry renders a real
 * geometry/SVG preview, and previews are pairwise distinct within each gallery
 * (guards against the old duplicate/mismatched Unicode glyphs).
 */
import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { isLineDrawKind } from '../src/renderer/draw-shape'
import { CHART_GALLERY, SHAPE_GALLERY, SMARTART_GALLERY } from '../src/renderer/insert-presets'
import {
  shapePreviewBox,
  shapePreviewPath,
  smartArtPreviewShapes,
} from '../src/renderer/components/gallery-previews'
import { ChartKindThumb } from '../src/renderer/components/ChartTypeDialog'
import { ShapeGalleryContent } from '../src/renderer/components/ShapeGalleryPopover'

function expectUnique(ids: Map<string, string>) {
  const seen = new Map<string, string>()
  for (const [key, sig] of ids) {
    const dup = seen.get(sig)
    expect(dup, `${key} and ${dup} share the same preview`).toBeUndefined()
    seen.set(sig, key)
  }
}

describe('shape gallery previews', () => {
  const prsts = SHAPE_GALLERY.flatMap((g) => g.shapes.map((s) => s.prst))

  it('covers every gallery prst with non-empty path data', () => {
    for (const prst of prsts) {
      const { w, h } = shapePreviewBox(prst, 18)
      expect(shapePreviewPath(prst, w, h), prst).toBeTruthy()
    }
  })

  it('renders pairwise distinct previews', () => {
    expectUnique(
      new Map(
        prsts.map((prst) => {
          const { w, h } = shapePreviewBox(prst, 18)
          return [prst, `${w}x${h}:${shapePreviewPath(prst, w, h)}`]
        }),
      ),
    )
  })

  it('renders every supported preset in the shared change-shape gallery', () => {
    const html = renderToStaticMarkup(createElement(ShapeGalleryContent, { onPick: () => {} }))
    for (const prst of prsts) {
      const assertion = expect(html)
      if (isLineDrawKind(prst)) assertion.not.toContain(`data-prst="${prst}"`)
      else assertion.toContain(`data-prst="${prst}"`)
    }
  })
})

describe('chart gallery previews', () => {
  it('renders a non-empty, distinct SVG thumb per kind', () => {
    const ids = new Map(
      CHART_GALLERY.map((def) => {
        const svg = renderToStaticMarkup(createElement(ChartKindThumb, { kind: def.kind }))
        expect(svg, def.kind).toContain('<svg')
        return [def.kind, svg]
      }),
    )
    expect(ids.size).toBe(CHART_GALLERY.length)
    expectUnique(ids)
  })
})

describe('smartart gallery previews', () => {
  it('derives non-empty, distinct previews from the engine layout', () => {
    const ids = new Map(
      SMARTART_GALLERY.map((def) => {
        const shapes = smartArtPreviewShapes(def.layout)
        expect(shapes.length, def.layout).toBeGreaterThan(0)
        for (const s of shapes) expect(s.d, `${def.layout} child path`).toBeTruthy()
        return [def.layout, JSON.stringify(shapes)]
      }),
    )
    expect(ids.size).toBe(SMARTART_GALLERY.length)
    expectUnique(ids)
  })
})
