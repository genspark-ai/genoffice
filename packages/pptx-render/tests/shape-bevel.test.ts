import { describe, it, expect } from 'vitest'
import {
  bevelMaterialFaceColor,
  buildBevelFaces,
  extrusionFrontFace,
  roundRectRing,
} from '../src/scene3d'

const lum = (hex: string) => {
  const n = parseInt(hex.slice(1), 16)
  return 0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)
}
const scene = (over: Record<string, unknown> = {}) => ({
  cameraPreset: 'orthographicFront',
  lightRig: 'chilly',
  lightDir: 't',
  material: 'translucentPowder',
  bevelTop: { wEmu: 127000, hEmu: 25400, preset: 'softRound' },
  ...over,
})

describe('translucentPowder face tint (probe: six fills × seven rigs)', () => {
  it('lifts the fill toward the rig light: 0.69 × fill + rig offset, capped at 227', () => {
    expect(bevelMaterialFaceColor('#5B9BD5', scene())).toBe('#729ece') // PowerPoint 114,159,206
    expect(bevelMaterialFaceColor('#FFC000', scene({ lightRig: 'threePt' }))).toBe('#e3e362')
  })
  it('leaves other materials at the fill color', () => {
    expect(bevelMaterialFaceColor('#5B9BD5', scene({ material: 'matte' }))).toBe('#5B9BD5')
    // glow probed later: gray K = 54
    expect(bevelMaterialFaceColor('#5B9BD5', scene({ lightRig: 'glow' }))).toBe('#75a1c9')
    expect(bevelMaterialFaceColor('#5B9BD5', scene({ lightRig: undefined }))).toBe(
      bevelMaterialFaceColor('#5B9BD5', scene({ lightRig: 'threePt' })),
    )
  })
})

describe('buildBevelFaces', () => {
  const W = 189
  const H = 68
  const build = (over: Record<string, unknown> = {}) =>
    buildBevelFaces({
      rings: [roundRectRing(W, H, 10)],
      w: W,
      h: H,
      scene: scene(over) as never,
      frontColor: '#A5A5A5',
      bevelPx: 13.3,
    })!
  const edgeBands = (r: NonNullable<ReturnType<typeof build>>) => {
    const byEdge: Record<string, number[]> = { T: [], R: [], B: [], L: [] }
    const face = lum(r.faces[r.faces.length - 1]!.color)
    for (const f of r.faces.slice(0, -1)) {
      const nums = (f.path.match(/-?\d+(\.\d+)?/g) ?? []).map(Number)
      const xs = nums.filter((_, i) => i % 2 === 0)
      const ys = nums.filter((_, i) => i % 2 === 1)
      const cx = xs.reduce((a, b) => a + b) / xs.length
      const cy = ys.reduce((a, b) => a + b) / ys.length
      const d = lum(f.color) - face
      if (cy < 20 && cx > 30 && cx < W - 30) byEdge.T!.push(d)
      else if (cy > H - 20 && cx > 30 && cx < W - 30) byEdge.B!.push(d)
      else if (cx < 20 && cy > 20 && cy < H - 20) byEdge.L!.push(d)
      else if (cx > W - 20 && cy > 20 && cy < H - 20) byEdge.R!.push(d)
    }
    return byEdge
  }

  it('chilly dir=t: top and right bands brighter than the face, bottom and left darker', () => {
    const r = build()
    const last = r.faces[r.faces.length - 1]!
    expect(last.color).toBe('#a5a5ad') // tinted face, not the raw fill
    const e = edgeBands(r)
    expect(Math.max(...e.T!)).toBeGreaterThan(10)
    expect(Math.max(...e.R!)).toBeGreaterThan(15)
    expect(Math.min(...e.B!)).toBeLessThan(-15)
    expect(Math.min(...e.L!)).toBeLessThan(-20)
    // bands fade toward the face: the innermost sub-band is the mildest
    expect(Math.abs(e.T![e.T!.length - 1]!)).toBeLessThan(Math.abs(Math.max(...e.T!)))
  })

  it('rotating the rig to dir=b swaps the lit edges', () => {
    const e = edgeBands(build({ lightDir: 'b' }))
    expect(Math.min(...e.T!)).toBeLessThan(-15)
    expect(Math.max(...e.B!)).toBeGreaterThan(5)
  })

  it('matte bands are stronger than translucentPowder bands', () => {
    const p = edgeBands(build())
    const m = edgeBands(build({ material: 'matte' }))
    expect(Math.min(...m.L!)).toBeLessThan(Math.min(...p.L!))
  })

  it('the flat material draws no bevel', () => {
    expect(
      buildBevelFaces({
        rings: [roundRectRing(W, H, 10)],
        w: W,
        h: H,
        scene: scene({ material: 'flat' }) as never,
        frontColor: '#A5A5A5',
        bevelPx: 13.3,
      }),
    ).toBeNull()
  })

  it('a non-solid fill keeps its own paint on the inner face', () => {
    const r = buildBevelFaces({
      rings: [roundRectRing(W, H, 10)],
      w: W,
      h: H,
      scene: scene() as never,
      frontColor: '#A5A5A5',
      frontUsesFill: true,
      bevelPx: 13.3,
    })!
    expect(r.faces[r.faces.length - 1]!.front).toBe(true)
  })
})

it('skips bands for rigs without a calibration (contrasting is specular-only, twoPt unmatched)', () => {
  const rings = [[0, 0, 100, 0, 100, 40, 0, 40]]
  const base = { rings, w: 100, h: 40, frontColor: '#5B9BD5', bevelPx: 6 }
  expect(
    buildBevelFaces({ ...base, scene: scene({ lightRig: 'contrasting', material: 'metal' }) }),
  ).toBeNull()
  expect(
    buildBevelFaces({ ...base, scene: scene({ lightRig: 'twoPt', material: 'matte' }) }),
  ).toBeNull()
  expect(
    buildBevelFaces({ ...base, scene: scene({ lightRig: 'chilly', material: 'matte' }) }),
  ).not.toBeNull()
})

const edgeOf = (path: string, w: number, h: number) => {
  const nums = (path.match(/-?\d+(\.\d+)?/g) ?? []).map(Number)
  const xs = nums.filter((_, i) => i % 2 === 0)
  const ys = nums.filter((_, i) => i % 2 === 1)
  const cx = xs.reduce((a, b) => a + b, 0) / xs.length
  const cy = ys.reduce((a, b) => a + b, 0) / ys.length
  if (cy < h * 0.2 && cx > 20 && cx < w - 20) return 'T'
  if (cy > h * 0.8 && cx > 20 && cx < w - 20) return 'B'
  if (cx < 20) return 'L'
  if (cx > w - 20) return 'R'
  return ''
}
const outerBands = (rig: string, material: string) => {
  const w = 200
  const h = 80
  const r = buildBevelFaces({
    rings: [[0, 0, w, 0, w, h, 0, h]],
    w,
    h,
    scene: scene({ lightRig: rig, material }),
    frontColor: '#A5A5A5',
    bevelPx: 10,
  })!
  const face = lum(r.faces[r.faces.length - 1]!.color)
  const by: Record<string, number> = {}
  for (const f of r.faces.slice(0, 4)) by[edgeOf(f.path, w, h)] = lum(f.color) - face
  return by
}

it('glow lights the left band hardest and the right band less; top and bottom stay flat (probe)', () => {
  const b = outerBands('glow', 'matte')
  expect(b.L).toBeGreaterThan(b.R)
  expect(b.R).toBeGreaterThan(10)
  expect(Math.abs(b.T)).toBeLessThan(1)
  expect(Math.abs(b.B)).toBeLessThan(1)
})

it('morning lights the left band, darkens the top and lifts the bottom (probe: T -66, B +37, L +85)', () => {
  const b = outerBands('morning', 'matte')
  expect(b.L).toBeGreaterThan(50)
  expect(b.T).toBeLessThan(-20)
  expect(b.B).toBeGreaterThan(0)
})

it('colored rigs tint a matte face by the rig light (probe: flood gray 165 -> 140)', () => {
  expect(bevelMaterialFaceColor('#A5A5A5', scene({ lightRig: 'flood', material: 'matte' }))).toBe(
    '#8c8c8c',
  )
  expect(bevelMaterialFaceColor('#A5A5A5', scene({ lightRig: 'chilly', material: 'matte' }))).toBe(
    '#A5A5A5',
  )
  expect(bevelMaterialFaceColor('#A5A5A5', scene({ lightRig: 'flood', material: 'flat' }))).toBe(
    '#A5A5A5',
  )
  // powder under a colored rig: 0.69 x fill + per-channel additive
  expect(bevelMaterialFaceColor('#A5A5A5', scene({ lightRig: 'freezing' }))).toBe('#828b9b')
})

it('hands the renderer the whole outline as the shadow silhouette', () => {
  const rings = [[0, 0, 100, 0, 100, 40, 0, 40]]
  const r = buildBevelFaces({
    rings,
    w: 100,
    h: 40,
    scene: scene(),
    frontColor: '#5B9BD5',
    bevelPx: 6,
  })!
  expect(r.shadowPath).toBe('M 0 0 L 100 0 L 100 40 L 0 40 Z')
  expect(r.flat).toBe(true)
})

it('keeps a translucent fill alpha on the tinted face and every band', () => {
  expect(bevelMaterialFaceColor('#5B9BD580', scene())).toBe('#729ece80')
  const rings = [[0, 0, 100, 0, 100, 40, 0, 40]]
  const r = buildBevelFaces({
    rings,
    w: 100,
    h: 40,
    scene: scene(),
    frontColor: '#5B9BD580',
    bevelPx: 6,
  })!
  for (const f of r.faces) expect(f.color).toMatch(/^#[0-9a-f]{6}80$/)
})

it('shades the bands the same for every non-flat material (probe: matte = plastic = metal = warmMatte)', () => {
  const rings = [[0, 0, 100, 0, 100, 40, 0, 40]]
  const base = { rings, w: 100, h: 40, frontColor: '#A5A5A5', bevelPx: 6 }
  const colors = (material?: string) =>
    buildBevelFaces({ ...base, scene: scene({ material }) })!.faces.map((f) => f.color)
  expect(colors('matte')).toEqual(colors('metal'))
  expect(colors('matte')).toEqual(colors('warmMatte'))
  expect(colors('matte')).toEqual(colors(undefined))
})

it('backs the shadow silhouette with the last filled face, never the transparent stroke face', () => {
  const rings = [[0, 0, 100, 0, 100, 40, 0, 40]]
  const r = buildBevelFaces({
    rings,
    w: 100,
    h: 40,
    scene: scene(),
    frontColor: '#5B9BD5',
    bevelPx: 6,
    strokeColor: '#000000',
    strokeWidthPx: 1,
  })!
  expect(r.faces.at(-1)!.color).toBe('transparent')
  expect(extrusionFrontFace(r.faces)!.color).toBe('#729ece')
  expect(extrusionFrontFace([{ path: 'M 0 0 Z', color: 'transparent' }])).toBeUndefined()
})

it('a lightRig rev spin turns the bevel lights (probe: glow rev=25° → top from the strong light, bottom from the weak one)', () => {
  const w = 200
  const h = 80
  const bands = (lightRot?: { lat: number; lon: number; rev: number }) => {
    const r = buildBevelFaces({
      rings: [[0, 0, w, 0, w, h, 0, h]],
      w,
      h,
      scene: scene({ lightRig: 'glow', material: 'matte', ...(lightRot ? { lightRot } : {}) }),
      frontColor: '#A5A5A5',
      bevelPx: 10,
    })!
    const face = lum(r.faces[r.faces.length - 1]!.color)
    const by: Record<string, number> = {}
    for (const f of r.faces.slice(0, 4)) by[edgeOf(f.path, w, h)] = lum(f.color) - face
    return by
  }
  const flat = bands()
  expect(Math.abs(flat.T)).toBeLessThan(1)
  const spun = bands({ lat: 0, lon: 0, rev: 1500000 })
  expect(spun.T).toBeGreaterThan(spun.B)
  expect(spun.B).toBeGreaterThan(5)
  expect(spun.L).toBeGreaterThan(spun.R)
})
