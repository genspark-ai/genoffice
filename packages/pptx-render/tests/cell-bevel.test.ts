/** cell3D bevel shading, checked against a PowerPoint probe deck (96dpi, default 8px bevel). */
import { describe, it, expect } from 'vitest'
import { bevelFaceColor, buildCellBevel } from '../src/cell-bevel'

const outer = (b: ReturnType<typeof buildCellBevel>, e: 't' | 'r' | 'b' | 'l') =>
  b.edges[e][0]!.color

describe('cell3D bevel', () => {
  it('darkens the face to 85% of the fill (F9F5F4 → ~D4D0D0, 990033 → ~820029)', () => {
    expect(bevelFaceColor('#F9F5F4')).toBe('#D4D0CF')
    expect(bevelFaceColor('#990033')).toBe('#82002B')
  })

  it('flood rig from the top: left specular (white), top mild light, right deep shadow, bottom mild shadow', () => {
    const b = buildCellBevel('#F9F5F4', 8, undefined, 't')
    expect(b.widthPx).toBe(8)
    expect(outer(b, 'l')).toBe('#FFFFFF')
    // probe: top peaks ≈ 249 on a 208 face, right ≈ 85, bottom ≈ 188
    expect(parseInt(outer(b, 't').slice(1, 3), 16)).toBeGreaterThan(240)
    expect(parseInt(outer(b, 'r').slice(1, 3), 16)).toBeCloseTo(85, -1)
    expect(parseInt(outer(b, 'b').slice(1, 3), 16)).toBeCloseTo(189, -1)
    // circle bevel fades into the face over the inner half
    expect(b.edges.r.map((s) => s.pos)).toEqual([0, 0.45, 0.95, 1])
    expect(b.edges.r[3]!.color).toBe(bevelFaceColor('#F9F5F4'))
  })

  it('dir="r" rotates the pattern: specular on top, deep shadow at the bottom', () => {
    const b = buildCellBevel('#F9F5F4', 8, undefined, 'r')
    expect(outer(b, 't')).toBe('#FFFFFF')
    expect(parseInt(outer(b, 'b').slice(1, 3), 16)).toBeCloseTo(85, -1)
  })

  it("keeps a translucent fill's alpha on the face and every band", () => {
    expect(bevelFaceColor('#F9F5F480')).toBe('#D4D0CF80')
    const b = buildCellBevel('#F9F5F480', 8, undefined, 't')
    for (const e of ['t', 'r', 'b', 'l'] as const)
      for (const s of b.edges[e]) expect(s.color).toMatch(/^#[0-9A-F]{6}80$/)
  })

  it('no lightRig behaves like dir="t"; the angle preset keeps flat bands', () => {
    expect(buildCellBevel('#4472C4', 8, undefined, undefined)).toEqual(
      buildCellBevel('#4472C4', 8, undefined, 't'),
    )
    const flat = buildCellBevel('#F9F5F4', 8, 'angle', 't')
    expect(flat.edges.r.map((s) => s.pos)).toEqual([0, 0.95, 1])
    expect(flat.edges.r[0]!.color).toBe(flat.edges.r[1]!.color)
  })
})
