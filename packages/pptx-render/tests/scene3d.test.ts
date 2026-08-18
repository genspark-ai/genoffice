import { describe, it, expect } from 'vitest'
import {
  buildExtrusion,
  inPlaneRotationDeg,
  flattenSvgPath,
  ellipseRing,
  roundRectRing,
  type Scene3DProps,
} from '../src/scene3d'

const RECT = [0, 0, 200, 0, 200, 100, 0, 100]

function scene(over: Partial<Scene3DProps> = {}): Scene3DProps {
  return { cameraPreset: 'isometricRightUp', lightRig: 'threePt', lightDir: 't', ...over }
}

function build(over: Partial<Scene3DProps> = {}, extra: Record<string, unknown> = {}) {
  return buildExtrusion({
    rings: [RECT],
    w: 200,
    h: 100,
    depthPx: 40,
    zPx: 0,
    scene: scene(over),
    frontColor: '#C5E0B4',
    sideColor: '#ED7D31',
    ...extra,
  })
}

/** Parse "M x y L …" into [x,y,…] (absolute path emitted by the extrusion builder). */
function pathPoints(d: string): number[] {
  return d
    .split(/[\s,]+/)
    .filter((t) => /^-?[\d.]+$/.test(t))
    .map(Number)
}

function centroid(pts: number[]): { x: number; y: number } {
  let x = 0
  let y = 0
  for (let i = 0; i < pts.length; i += 2) {
    x += pts[i]!
    y += pts[i + 1]!
  }
  return { x: (x * 2) / pts.length, y: (y * 2) / pts.length }
}

describe('flattenSvgPath', () => {
  it('polylines close into rings, curves get sampled', () => {
    const rings = flattenSvgPath('M 0 0 L 10 0 L 10 10 L 0 10 Z')
    expect(rings).toHaveLength(1)
    expect(rings[0]).toEqual([0, 0, 10, 0, 10, 10, 0, 10])
    const curved = flattenSvgPath('M 0 0 C 0 10 10 10 10 0 Z')
    expect(curved[0]!.length).toBeGreaterThan(8)
  })

  it('multiple subpaths become separate rings (donut hole)', () => {
    const rings = flattenSvgPath('M 0 0 L 20 0 L 20 20 L 0 20 Z M 5 5 L 15 5 L 15 15 L 5 15 Z')
    expect(rings).toHaveLength(2)
  })
})

describe('ring helpers', () => {
  it('ellipseRing stays on the ellipse', () => {
    const r = ellipseRing(200, 100, 16)
    for (let i = 0; i < r.length; i += 2) {
      const dx = (r[i]! - 100) / 100
      const dy = (r[i + 1]! - 50) / 50
      expect(dx * dx + dy * dy).toBeCloseTo(1, 6)
    }
  })

  it('roundRectRing clamps the radius and keeps corners inside the box', () => {
    const r = roundRectRing(100, 40, 500)
    for (let i = 0; i < r.length; i += 2) {
      expect(r[i]!).toBeGreaterThanOrEqual(-0.001)
      expect(r[i]!).toBeLessThanOrEqual(100.001)
      expect(r[i + 1]!).toBeGreaterThanOrEqual(-0.001)
      expect(r[i + 1]!).toBeLessThanOrEqual(40.001)
    }
  })
})

describe('inPlaneRotationDeg', () => {
  it('orthographicFront + rev-only rot spins the flat shape (opposite orientation)', () => {
    expect(
      inPlaneRotationDeg(
        scene({ cameraPreset: 'orthographicFront', cameraRot: { lat: 0, lon: 0, rev: 5400000 } }),
      ),
    ).toBe(-90)
  })

  it('extrusion, tilted cameras and unknown presets need the mesh path', () => {
    expect(
      inPlaneRotationDeg(
        scene({
          cameraPreset: 'orthographicFront',
          cameraRot: { lat: 0, lon: 0, rev: 5400000 },
          extrusionEmu: 1,
        }),
      ),
    ).toBeNull()
    expect(inPlaneRotationDeg(scene())).toBeNull()
    expect(inPlaneRotationDeg(scene({ cameraPreset: 'nope' }))).toBeNull()
    // a z shift changes perspective scale / oblique offset even without depth
    expect(inPlaneRotationDeg(scene({ cameraPreset: 'perspectiveFront', zEmu: 914400 }))).toBeNull()
  })
})

describe('buildExtrusion', () => {
  it('depth 0 projects just the tilted front cap; unknown cameras bail out', () => {
    const flat = build({}, { depthPx: 0 })!
    expect(flat.faces).toHaveLength(1)
    expect(flat.faces[0]!.color).not.toBe('transparent')
    expect(build({ cameraPreset: 'bogus' })).toBeNull()
  })

  it('isometric rect: back cap, visible walls, then front cap on top (painter order)', () => {
    const ext = build()!
    expect(ext.wireframe).toBeUndefined()
    const front = ext.faces[ext.faces.length - 1]!
    expect(front.color).not.toBe('transparent')
    // threePt front-face illumination ≈ 1.04 → the cap keeps roughly its base color
    expect(front.color.toUpperCase()).toBe('#CEEABC')
    // rect → 2 walls survive culling (left + top), plus the front cap; back cap is hidden
    expect(ext.faces.length).toBeGreaterThanOrEqual(3)
  })

  it('modern cameras keep the authored cap appearance under every preset', () => {
    const a = build({ cameraPreset: 'isometricRightUp' })!
    const b = build({ cameraPreset: 'perspectiveContrastingLeftFacing' })!
    expect(a.faces[a.faces.length - 1]!.color).toBe(b.faces[b.faces.length - 1]!.color)
  })

  it('legacy oblique: side walls stay visible and the extrusion offsets along the skew', () => {
    const ext = build({ cameraPreset: 'legacyObliqueTopLeft' })!
    expect(ext.faces.length).toBeGreaterThanOrEqual(3) // ≥2 walls + front cap
    // Wireframe emits both caps: 50% skew at -45° puts the back cap 40·0.5·cos45 ≈ 14.1px
    // right/up of the front cap.
    const wf = build({ cameraPreset: 'legacyObliqueTopLeft', material: 'legacyWireframe' })!
    const back = centroid(pathPoints(wf.faces[0]!.path))
    const front = centroid(pathPoints(wf.faces[wf.faces.length - 1]!.path))
    expect(back.x - front.x).toBeCloseTo(14.14, 1)
    expect(back.y - front.y).toBeCloseTo(-14.14, 1)
  })

  it('perspective projection shrinks the far cap', () => {
    const wf = build({ cameraPreset: 'perspectiveFront', material: 'legacyWireframe' })!
    const width = (pts: number[]) =>
      Math.max(...pts.filter((_, i) => i % 2 === 0)) -
      Math.min(...pts.filter((_, i) => i % 2 === 0))
    const backW = width(pathPoints(wf.faces[0]!.path))
    const frontW = width(pathPoints(wf.faces[wf.faces.length - 1]!.path))
    expect(frontW).toBeCloseTo(200, 1)
    expect(backW).toBeLessThan(frontW)
  })

  it('extrusionClr overrides the wall color source', () => {
    const plain = build()!
    const tinted = build({ extrusionColor: '#00FFFF' })!
    const wallOf = (ext: typeof plain) => ext.faces[1]!.color
    expect(wallOf(tinted)).not.toBe(wallOf(plain))
  })

  it('legacyWireframe renders edge strokes only', () => {
    const ext = build(
      { material: 'legacyWireframe' },
      { strokeColor: '#000000', strokeWidthPx: 1 },
    )!
    expect(ext.wireframe).toBe(true)
    for (const f of ext.faces) {
      expect(f.color).toBe('transparent')
      expect(f.stroke).toBe('#000000')
    }
  })

  it('harsh rig shades the visible walls unevenly (left saturated, top dimmer)', () => {
    const ext = build({ cameraPreset: 'isometricOffAxis1Right', lightRig: 'harsh' })!
    const walls = ext.faces.slice(0, -1)
    const lum = (c: string) => parseInt(c.slice(1, 3), 16)
    const lums = walls.map((f) => lum(f.color))
    expect(Math.max(...lums)).toBeGreaterThan(Math.min(...lums) + 20)
  })

  it('holes survive into both caps as subpaths', () => {
    const ext = buildExtrusion({
      rings: [
        [0, 0, 200, 0, 200, 100, 0, 100],
        [50, 25, 150, 25, 150, 75, 50, 75],
      ],
      w: 200,
      h: 100,
      depthPx: 40,
      zPx: 0,
      scene: scene(),
      frontColor: '#C5E0B4',
      sideColor: '#ED7D31',
    })!
    const front = ext.faces[ext.faces.length - 1]!
    expect(front.path.match(/M /g)!.length).toBe(2)
    // 8 wall candidates (4 outer + 4 inner); roughly half survive culling
    expect(ext.faces.length).toBeGreaterThanOrEqual(5)
  })
})
