/**
 * @genoffice/pptx-render — high-fidelity pptx render layer (Phase 2).
 *
 * Layering (approach A: data-driven + thin adapters):
 *   coords      2.1 EMU→px coordinate system + viewport
 *   fill        2.2 fill/stroke resolution
 *   metrics     2.3 font metrics (opentype.js exact / heuristic fallback)
 *   text-layout 2.3 text layout (wrapping/line spacing/autofit/vertical alignment)
 *   build-slide aggregation: Slide → RenderTree (fidelity is locked at this layer; unit-testable without canvas)
 *
 * The Konva adapter (drawing the RenderTree) is added when apps/slides gets a frontend; not in this package.
 */
export * from './coords'
export * from './render-tree'
export * from './fill'
export * from './metrics'
export * from './text-layout'
export * from './build-slide'
export * from './build-chart'
export * from './preset-geometry'
export * from './scene3d'
export { patternGrid } from './pattern-fills'
