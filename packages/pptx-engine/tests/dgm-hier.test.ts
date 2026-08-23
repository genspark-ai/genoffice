import { describe, expect, it } from 'vitest'
import { layoutHierTree, parseHierConstraints, type HierTreeNode } from '../src/dgm-hier'

const n = (
  text: string,
  children: HierTreeNode[] = [],
  extra: Partial<HierTreeNode> = {},
): HierTreeNode => ({
  texts: text ? [text] : [],
  children,
  ...extra,
})

const CONS = parseHierConstraints(undefined)

// Frame chosen so the unit layout is height-bound with a comfortable margin
const layout = (roots: HierTreeNode[]) => layoutHierTree(roots, CONS, 8000000, 6000000)!

const box = (g: ReturnType<typeof layout>, text: string) =>
  g.boxes.find((b) => b.node.texts[0] === text)!

describe('parseHierConstraints', () => {
  it('reads factors from orgChart1-style constraint lists', () => {
    const xml = `
      <dgm:constr type="h" for="des" forName="rootComposite1" refType="w" refFor="des" refForName="rootComposite1" fact="0.6"/>
      <dgm:constr type="sibSp" refType="w" refFor="des" refForName="rootComposite1" fact="0.3"/>
      <dgm:constr type="sp" for="des" forName="hierRoot1" refType="w" refFor="des" refForName="rootComposite1" fact="0.25"/>
      <dgm:constr type="w" for="ch" forName="rootConnector1" refType="w" refFor="ch" refForName="rootText1" fact="0.2"/>
      <dgm:constr type="primFontSz" val="65"/>`
    const c = parseHierConstraints(xml)
    expect(c.boxAspect).toBe(0.6)
    expect(c.sibSp).toBe(0.3)
    expect(c.sp).toBe(0.25)
    expect(c.trunkOff).toBeCloseTo(0.1)
    expect(c.fontMax).toBe(65)
  })

  it('falls back to orgChart1 defaults without a layout part', () => {
    expect(CONS.boxAspect).toBe(0.5)
    expect(CONS.sibSp).toBe(0.21)
  })

  it('reads primFontSz with attributes between type and val, skipping valueless op=equ', () => {
    const xml = `
      <dgm:constr type="primFontSz" for="des" ptType="node" op="equ"/>
      <dgm:constr type="primFontSz" for="ch" forName="rootText1" val="48"/>`
    expect(parseHierConstraints(xml).fontMax).toBe(48)
  })
})

describe('layoutHierTree', () => {
  // smartart-org-chart2.pptx shape: A → (B1 → C1, C2 → D1, D2; B2 → C3, C4)
  const tree = () =>
    n('A', [n('B1', [n('C1'), n('C2', [n('D1'), n('D2')])]), n('B2', [n('C3'), n('C4')])])

  it('all node boxes share one size with the constraint aspect', () => {
    const g = layout([tree()])
    const w = g.boxes[0]!.w
    for (const b of g.boxes) {
      expect(b.w).toBeCloseTo(w)
      expect(b.h / b.w).toBeCloseTo(CONS.boxAspect)
    }
  })

  it('deep all-leaf branches hang: boxes indent 0.25w from the parent left', () => {
    const g = layout([tree()])
    const b2 = box(g, 'B2')
    const c3 = box(g, 'C3')
    const c4 = box(g, 'C4')
    expect(c3.x - b2.x).toBeCloseTo(0.25 * b2.w, 0)
    expect(c4.x).toBeCloseTo(c3.x)
    // stacked one row apart
    expect(c4.y - c3.y).toBeCloseTo((CONS.boxAspect + CONS.sp) * b2.w, 0)
  })

  it('root-level children never hang even when all leaves', () => {
    const g = layout([n('A', [n('B1'), n('B2')])])
    expect(box(g, 'B1').y).toBeCloseTo(box(g, 'B2').y)
  })

  it('3-row charts stay standard even with deep all-leaf branches (recursion measured)', () => {
    // smartart-recursion.pptx: A → (B1 → C1, C2; B2 → C3) renders all-std in PowerPoint
    const g = layout([n('A', [n('B1', [n('C1'), n('C2')]), n('B2', [n('C3')])])])
    expect(box(g, 'C1').y).toBeCloseTo(box(g, 'C2').y)
    expect(box(g, 'C2').y).toBeCloseTo(box(g, 'C3').y)
  })

  it('contour packing lets a hang overhang slide past a sibling box row', () => {
    // B1's D-hang extends right below B2's box row; B2 must pack against the
    // C row contour, not the D overhang bounding box (PowerPoint measured)
    const g = layout([tree()])
    const b1 = box(g, 'B1')
    const b2 = box(g, 'B2')
    const d1 = box(g, 'D1')
    expect(b2.x).toBeLessThan(d1.x + d1.w)
    // measured center pitch: 1.815 box widths
    expect((b2.x - b1.x) / b1.w).toBeCloseTo(1.815, 1)
  })

  it('parents center on the midpoint of their outer children', () => {
    const g = layout([tree()])
    const a = box(g, 'A')
    const b1 = box(g, 'B1')
    const b2 = box(g, 'B2')
    expect(a.x + a.w / 2).toBeCloseTo((b1.x + b1.w / 2 + b2.x + b2.w / 2) / 2, 0)
  })

  it('leaf assistants get their own row tucked left of the trunk', () => {
    // smartart-org-chart.pptx shape
    const g = layout([
      n('Manager', [n('Assistant', [], { asst: true }), n('Employee'), n('Employee2')]),
      n('Manager2'),
    ])
    const mgr = box(g, 'Manager')
    const asst = box(g, 'Assistant')
    const emp = box(g, 'Employee')
    const trunk = mgr.x + mgr.w / 2
    expect(asst.x + asst.w).toBeCloseTo(trunk - (CONS.sp / 2) * mgr.w, 0)
    expect(asst.y).toBeGreaterThan(mgr.y)
    expect(emp.y).toBeGreaterThan(asst.y)
  })

  it('explicit hierBranch overrides the auto rule', () => {
    const g = layout([n('A', [n('B', [n('C1'), n('C2')], { hierBranch: 'std' })])])
    expect(box(g, 'C1').y).toBeCloseTo(box(g, 'C2').y)
    const g2 = layout([n('A', [n('B', [n('C1'), n('C2')], { hierBranch: 'l' })])])
    const b = box(g2, 'B')
    const c1 = box(g2, 'C1')
    expect(c1.x + c1.w).toBeCloseTo(b.x + b.w - 0.25 * b.w, 0)
  })

  it('scales into the frame and stays inside it', () => {
    const g = layoutHierTree([tree()], CONS, 3000000, 2000000)!
    for (const b of g.boxes) {
      expect(b.x).toBeGreaterThanOrEqual(-1)
      expect(b.y).toBeGreaterThanOrEqual(-1)
      expect(b.x + b.w).toBeLessThanOrEqual(3000001)
      expect(b.y + b.h).toBeLessThanOrEqual(2000001)
    }
  })
})
