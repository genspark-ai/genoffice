import { describe, it, expect } from 'vitest'
import { layoutDiagramFallback } from '../src/parse'

const DATA = `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="d" xmlns:a="a">
<dgm:ptLst>
<dgm:pt modelId="doc" type="doc"><dgm:prSet/></dgm:pt>
<dgm:pt modelId="n1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n2"/>
<dgm:pt modelId="n3"/>
<dgm:pt modelId="n4"/>
<dgm:pt modelId="n5"/>
<dgm:pt modelId="p1" type="pres"/>
</dgm:ptLst>
<dgm:cxnLst>
<dgm:cxn modelId="c1" srcId="doc" destId="n1" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c2" srcId="doc" destId="n2" srcOrd="1" destOrd="0"/>
<dgm:cxn modelId="c3" srcId="doc" destId="n3" srcOrd="2" destOrd="0"/>
<dgm:cxn modelId="c4" srcId="doc" destId="n4" srcOrd="3" destOrd="0"/>
<dgm:cxn modelId="c5" srcId="doc" destId="n5" srcOrd="4" destOrd="0"/>
<dgm:cxn modelId="c6" type="presOf" srcId="n1" destId="p1" srcOrd="0" destOrd="0"/>
</dgm:cxnLst></dgm:dataModel>`

describe('layoutDiagramFallback', () => {
  it('lays out a flat 5-node list as a 2-column snake, last row centered', () => {
    const shapes = layoutDiagramFallback(DATA, {}, 9144000, 6858000)
    expect(shapes).toHaveLength(5)
    const xs = shapes.map((s) => s.transform.offset.x)
    const ys = shapes.map((s) => s.transform.offset.y)
    // Two columns, three rows
    expect(new Set(xs.slice(0, 4)).size).toBe(2)
    expect(new Set(ys).size).toBe(3)
    // Last (odd) block centered between the two columns
    expect(xs[4]).toBeGreaterThan(xs[0]!)
    expect(xs[4]).toBeLessThan(xs[1]!)
    // Aspect 0.6 and node text carried through
    const s0 = shapes[0]!
    expect(s0.transform.offset.cy / s0.transform.offset.cx).toBeCloseTo(0.6, 2)
    expect(JSON.stringify(shapes[0])).toContain('Alpha')
  })

  it('explicit type="node" points are kept', () => {
    const explicit = DATA.replace('<dgm:pt modelId="n2"/>', '<dgm:pt modelId="n2" type="node"/>')
    expect(layoutDiagramFallback(explicit, {}, 9144000, 6858000)).toHaveLength(5)
  })

  it('a height-bound grid stays inside the frame', () => {
    // Wide short frame: height binds; grid bottom must not pass frameCy
    const shapes = layoutDiagramFallback(DATA, {}, 18000000, 2000000)
    const bottom = Math.max(...shapes.map((s) => s.transform.offset.y + s.transform.offset.cy))
    expect(bottom).toBeLessThanOrEqual(2000000)
  })

  const HIER_DATA = `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="d" xmlns:a="a">
<dgm:ptLst>
<dgm:pt modelId="doc" type="doc"/>
<dgm:pt modelId="n1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>A</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n2"><dgm:t><a:bodyPr/><a:p><a:r><a:t>B1</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n3"><dgm:t><a:bodyPr/><a:p><a:r><a:t>B2</a:t></a:r></a:p></dgm:t></dgm:pt>
</dgm:ptLst>
<dgm:cxnLst>
<dgm:cxn modelId="c1" srcId="doc" destId="n1" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c2" srcId="n1" destId="n2" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c3" srcId="n1" destId="n3" srcOrd="1" destOrd="0"/>
</dgm:cxnLst></dgm:dataModel>`

  it('orgChart keeps root children side by side (hang never fires at the root row)', () => {
    const shapes = layoutDiagramFallback(HIER_DATA, {}, 9144000, 6858000, 'orgChart1')
    const boxes = shapes.filter((sp: any) => sp.text)
    expect(boxes).toHaveLength(3)
    const [a, b1, b2] = boxes as any[]
    expect(a.transform.offset.y).toBeLessThan(b1.transform.offset.y)
    // PowerPoint measured (smartart-org-chart): a root's leaf children stay on one row
    expect(b2.transform.offset.y).toBe(b1.transform.offset.y)
    expect(b2.transform.offset.x).toBeGreaterThan(b1.transform.offset.x)
  })

  it('orgChart hangs an all-leaves branch on tall charts (std depth ≥ 4, PPT measured)', () => {
    // R → n1(B1,B2 leaves) + a 4-row sibling chain n4→n5→n6: n1's leaves hang
    const deep = HIER_DATA.replace('srcId="doc" destId="n1"', 'srcId="doc" destId="n0"')
      .replace(
        '<dgm:pt modelId="n1">',
        '<dgm:pt modelId="n0"><dgm:t><a:bodyPr/><a:p><a:r><a:t>R</a:t></a:r></a:p></dgm:t></dgm:pt>' +
          '<dgm:pt modelId="n4"/><dgm:pt modelId="n5"/><dgm:pt modelId="n6"/><dgm:pt modelId="n1">',
      )
      .replace(
        '</dgm:cxnLst>',
        '<dgm:cxn modelId="c0" srcId="n0" destId="n1" srcOrd="0" destOrd="0"/>' +
          '<dgm:cxn modelId="c7" srcId="n0" destId="n4" srcOrd="1" destOrd="0"/>' +
          '<dgm:cxn modelId="c8" srcId="n4" destId="n5" srcOrd="0" destOrd="0"/>' +
          '<dgm:cxn modelId="c9" srcId="n5" destId="n6" srcOrd="0" destOrd="0"/></dgm:cxnLst>',
      )
    const shapes = layoutDiagramFallback(deep, {}, 9144000, 6858000, 'orgChart1')
    const boxes = shapes.filter((sp: any) => sp.text) as any[]
    // n1's leaf children (B1, B2) hang on separate rows
    const texts = boxes.map((b) => JSON.stringify(b))
    const b1 = boxes[texts.findIndex((t) => t.includes('B1'))]!
    const b2 = boxes[texts.findIndex((t) => t.includes('B2'))]!
    expect(b2.transform.offset.y).toBeGreaterThan(b1.transform.offset.y)
  })

  it('orgChart keeps an all-leaves branch side by side on short charts (recursion measured)', () => {
    const deep = HIER_DATA.replace('srcId="doc" destId="n1"', 'srcId="doc" destId="n0"')
      .replace(
        '<dgm:pt modelId="n1">',
        '<dgm:pt modelId="n0"><dgm:t><a:bodyPr/><a:p><a:r><a:t>R</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="n1">',
      )
      .replace(
        '</dgm:cxnLst>',
        '<dgm:cxn modelId="c0" srcId="n0" destId="n1" srcOrd="0" destOrd="0"/></dgm:cxnLst>',
      )
    const shapes = layoutDiagramFallback(deep, {}, 9144000, 6858000, 'orgChart1')
    const boxes = shapes.filter((sp: any) => sp.text) as any[]
    expect(boxes).toHaveLength(4)
    const [, , b1, b2] = boxes
    expect(b2.transform.offset.y).toBe(b1.transform.offset.y)
  })

  it('chevron layout puts all nodes on one horizontal band of chevrons', () => {
    const shapes = layoutDiagramFallback(DATA, {}, 9144000, 6858000, 'chevron1')
    expect(shapes).toHaveLength(5)
    const ys = new Set(shapes.map((sp: any) => Math.round(sp.transform.offset.y)))
    expect(ys.size).toBe(1)
    expect((shapes[0] as any).presetGeometry).toBe('chevron')
  })

  it('columns family (hList1) emits header + tinted body per top node with cycled colors', () => {
    const colors = `<dgm:colorsDef xmlns:dgm="d" xmlns:a="a"><dgm:styleLbl name="node1">
      <dgm:fillClrLst meth="cycle"><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/></dgm:fillClrLst>
    </dgm:styleLbl></dgm:colorsDef>`
    const theme: any = { colors: { accent2: '#C0504D', accent3: '#9BBB59' } }
    const shapes = layoutDiagramFallback(HIER_DATA, { theme }, 9144000, 6858000, 'hList1', colors)
    // one top node -> header + body
    expect(shapes).toHaveLength(2)
    expect((shapes[0] as any).fill).toEqual({ type: 'solid', color: '#C0504D' })
    // body = 20% tint of the header color
    expect((shapes[1] as any).fill.color).toBe('#F2DCDB')
  })

  it('pyramid renders a single node as a full triangle with dark text', () => {
    const one = DATA.replace(/<dgm:cxn modelId="c[2-5][^/]*\/>/g, '')
    const shapes = layoutDiagramFallback(one, {}, 9144000, 6858000, 'pyramid1')
    expect(shapes).toHaveLength(1)
    expect((shapes[0] as any).presetGeometry).toBe('triangle')
    expect((shapes[0] as any).text.paragraphs[0].runs[0].color).toBe('#333333')
  })

  it('multi-level pyramid keeps collinear side edges via trapezoid adj', () => {
    const shapes = layoutDiagramFallback(DATA, {}, 9144000, 6858000, 'pyramid1')
    expect(shapes).toHaveLength(5)
    expect((shapes[0] as any).presetGeometry).toBe('triangle')
    const t1 = shapes[1] as any
    expect(t1.presetGeometry).toBe('trapezoid')
    // adj encodes the per-side inset = w0/(2n) relative to min(w,h)
    expect(t1.adjust?.adj).toBeGreaterThan(0)
    // widths grow row by row
    expect((shapes[2] as any).transform.offset.cx).toBeGreaterThan(t1.transform.offset.cx)
  })

  it('picture strips render outlined rows with dark labels', () => {
    const shapes = layoutDiagramFallback(DATA, {}, 9144000, 6858000, 'PictureStrips')
    expect(shapes).toHaveLength(5)
    const s0 = shapes[0] as any
    expect(s0.fill).toEqual({ type: 'none' })
    expect(s0.stroke).toBeTruthy()
  })

  it('hierarchies render top-node tiles with descendant bullet text', () => {
    const hier = DATA.replace('srcId="doc" destId="n5"', 'srcId="n1" destId="n5"').replace(
      '<dgm:pt modelId="n5"/>',
      '<dgm:pt modelId="n5"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Kid</a:t></a:r></a:p></dgm:t></dgm:pt>',
    )
    const shapes = layoutDiagramFallback(hier, {}, 9144000, 6858000)
    // n5 is now a child of n1: one tile fewer, its text folded into n1's tile as a bullet line
    expect(shapes).toHaveLength(4)
    const n1 = shapes[0] as any
    const paras = n1.text.paragraphs
    expect(paras.length).toBeGreaterThanOrEqual(2)
    expect(paras[1].bullet?.type).toBe('char')
  })
})

const EQ_DATA = `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="d" xmlns:a="a">
<dgm:ptLst>
<dgm:pt modelId="doc" type="doc"/>
<dgm:pt modelId="n1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Higher quality</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n2"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Improved value</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n3"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Better deal</a:t></a:r></a:p></dgm:t></dgm:pt>
</dgm:ptLst>
<dgm:cxnLst>
<dgm:cxn modelId="c1" srcId="doc" destId="n1" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c2" srcId="doc" destId="n2" srcOrd="1" destOrd="0"/>
<dgm:cxn modelId="c3" srcId="doc" destId="n3" srcOrd="2" destOrd="0"/>
</dgm:cxnLst></dgm:dataModel>`

describe('equation1 family (napierone 0005: a + b = c)', () => {
  it('lays out operand circles joined by mathPlus and mathEqual operators', () => {
    const shapes = layoutDiagramFallback(EQ_DATA, {}, 9144000, 3429000, 'equation1') as any[]
    const circles = shapes.filter((s) => s.presetGeometry === 'ellipse')
    expect(circles).toHaveLength(3)
    // circles are equal-sized and vertically centered
    expect(circles[0].transform.offset.cx).toBeCloseTo(circles[0].transform.offset.cy, 3)
    const ops = shapes.filter((s) => /^math/.test(s.presetGeometry ?? ''))
    expect(ops.map((o: any) => o.presetGeometry)).toEqual(['mathPlus', 'mathEqual'])
    // operators sit between the circles
    expect(ops[0].transform.offset.x).toBeGreaterThan(circles[0].transform.offset.x)
    expect(ops[0].transform.offset.x).toBeLessThan(circles[1].transform.offset.x)
    expect(JSON.stringify(circles[0])).toContain('Higher quality')
  })
})

const CHEV_DATA = `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="d" xmlns:a="a">
<dgm:ptLst>
<dgm:pt modelId="doc" type="doc"/>
<dgm:pt modelId="n1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Stage one</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n2"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Stage two</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="k1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Detail A</a:t></a:r></a:p></dgm:t></dgm:pt>
</dgm:ptLst>
<dgm:cxnLst>
<dgm:cxn modelId="c1" srcId="doc" destId="n1" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c2" srcId="doc" destId="n2" srcOrd="1" destOrd="0"/>
<dgm:cxn modelId="c3" srcId="n1" destId="k1" srcOrd="0" destOrd="0"/>
</dgm:cxnLst></dgm:dataModel>`

describe('chevron2 vertical chevron list (napierone 0001)', () => {
  it('renders a down-rotated chevron accent plus an outlined bullet card per row', () => {
    const shapes = layoutDiagramFallback(CHEV_DATA, {}, 8568952, 6048672, 'chevron2') as any[]
    const chevrons = shapes.filter((s) => s.presetGeometry === 'chevron')
    expect(chevrons).toHaveLength(2)
    for (const c of chevrons) expect(c.transform.rot).toBe(5400000)
    const cards = shapes.filter((s) => s.presetGeometry === 'roundRect')
    expect(cards).toHaveLength(2)
    // cards sit to the right of the chevron column
    for (const card of cards) expect(card.transform.offset.x).toBeGreaterThan(0)
    expect(JSON.stringify(cards[0])).toContain('Detail A')
    // the label overlay stays unrotated so its text remains horizontal
    const labels = shapes.filter((s) => s.presetGeometry === 'rect' && s.text)
    expect(labels.length).toBeGreaterThanOrEqual(2)
    expect(JSON.stringify(labels[0])).toContain('Stage one')
  })

  it('chevron1 keeps the horizontal band family', () => {
    const shapes = layoutDiagramFallback(CHEV_DATA, {}, 9144000, 2000000, 'chevron1') as any[]
    for (const s of shapes) expect(s.transform.rot ?? 0).toBe(0)
  })
})

describe('vProcess5 stepped process (napierone 0014)', () => {
  it('staggers boxes to the right as they descend with a down arrow between steps', () => {
    const shapes = layoutDiagramFallback(EQ_DATA, {}, 9144000, 6858000, 'vProcess5') as any[]
    const boxes = shapes.filter((s) => !s.presetGeometry || s.presetGeometry === 'rect')
    expect(boxes.length).toBe(3)
    // strictly increasing x and y
    for (let i = 1; i < boxes.length; i++) {
      expect(boxes[i].transform.offset.x).toBeGreaterThan(boxes[i - 1].transform.offset.x)
      expect(boxes[i].transform.offset.y).toBeGreaterThan(boxes[i - 1].transform.offset.y)
    }
    const arrows = shapes.filter((s) => s.presetGeometry === 'downArrow')
    expect(arrows.length).toBe(2)
  })
})

describe('flatGrid centering and tile text autofit (napierone 0005)', () => {
  it('centers the grid vertically in a width-bound frame', () => {
    // wide short frame relative to a 3x2 grid: rows leave headroom that splits evenly
    const six = DATA.replace(
      '<dgm:cxn modelId="c6" type="presOf" srcId="n1" destId="p1" srcOrd="0" destOrd="0"/>',
      '<dgm:cxn modelId="c6" srcId="doc" destId="p1" srcOrd="5" destOrd="0"/>',
    ).replace('<dgm:pt modelId="p1" type="pres"/>', '<dgm:pt modelId="p1"/>')
    const frameCx = 7680684
    const frameCy = 3888432
    const shapes = layoutDiagramFallback(six, {}, frameCx, frameCy) as any[]
    expect(shapes).toHaveLength(6)
    const top = Math.min(...shapes.map((s) => s.transform.offset.y))
    const bottom = Math.max(...shapes.map((s) => s.transform.offset.y + s.transform.offset.cy))
    expect(top).toBeGreaterThan(frameCy * 0.05)
    expect(Math.abs(top - (frameCy - bottom))).toBeLessThan(frameCy * 0.01)
  })

  it('wrapped tile text shrinks well below the cap', () => {
    const withText = DATA.replace(
      '<dgm:pt modelId="n2"/>',
      '<dgm:pt modelId="n2"><dgm:t><a:bodyPr/><a:p><a:r><a:t>De-layer service delivery</a:t></a:r></a:p></dgm:t></dgm:pt>',
    )
    const shapes = layoutDiagramFallback(withText, {}, 7680684, 3888432) as any[]
    const el = shapes.find((sp: any) =>
      sp.text?.paragraphs?.some((pp: any) =>
        pp.runs?.some((r: any) => r.text?.includes('De-layer')),
      ),
    ) as any
    const runSize = el.text.paragraphs[0].runs[0].fontSize
    expect(runSize).toBeLessThan(20)
    expect(runSize).toBeGreaterThan(10)
  })
})

describe('lProcess column process family (napierone 0005 p8)', () => {
  const LPROC_DATA = `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="d" xmlns:a="a">
<dgm:ptLst>
<dgm:pt modelId="doc" type="doc"/>
<dgm:pt modelId="h1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Engage</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="h2"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Identify</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="k1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Departments</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="k2"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Controls</a:t></a:r></a:p></dgm:t></dgm:pt>
</dgm:ptLst>
<dgm:cxnLst>
<dgm:cxn modelId="c1" srcId="doc" destId="h1" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c2" srcId="doc" destId="h2" srcOrd="1" destOrd="0"/>
<dgm:cxn modelId="c3" srcId="h1" destId="k1" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c4" srcId="h2" destId="k2" srcOrd="0" destOrd="0"/>
</dgm:cxnLst></dgm:dataModel>`

  it('renders one column per top node: header, connector dot, tinted child block', () => {
    const shapes = layoutDiagramFallback(LPROC_DATA, {}, 9144000, 3000000, 'lProcess1')
    // 2 columns × (header + dot + child) = 6 shapes
    expect(shapes).toHaveLength(6)
    const [h1, dot, kid] = shapes as any[]
    expect(JSON.stringify(h1)).toContain('Engage')
    expect(JSON.stringify(kid)).toContain('Departments')
    // dot is a small ellipse centered between header bottom and child top
    expect(JSON.stringify(dot)).toContain('ellipse')
    expect(dot.transform.offset.cx).toBeLessThan(h1.transform.offset.cx / 4)
    expect(dot.transform.offset.y).toBeGreaterThan(h1.transform.offset.y + h1.transform.offset.cy)
    expect(kid.transform.offset.y).toBeGreaterThan(dot.transform.offset.y)
    // both columns share the header y; second column sits right of the first
    const h2 = shapes[3] as any
    expect(h2.transform.offset.y).toBe(h1.transform.offset.y)
    expect(h2.transform.offset.x).toBeGreaterThan(h1.transform.offset.x)
  })

  it('childless lProcess2 nodes render as a large tinted top-anchored panel', () => {
    const single = LPROC_DATA.replace(/<dgm:cxn modelId="c3".*destOrd="0"\/>/, '')
      .replace(/<dgm:cxn modelId="c4".*destOrd="0"\/>/, '')
      .replace(/<dgm:pt modelId="k1">.*?<\/dgm:pt>/, '')
      .replace(/<dgm:pt modelId="k2">.*?<\/dgm:pt>/, '')
    const shapes = layoutDiagramFallback(single, {}, 9144000, 3000000, 'lProcess2')
    expect(shapes).toHaveLength(2)
    const p = shapes[0] as any
    // tall panel, not a slim header bar
    expect(p.transform.offset.cy).toBeGreaterThan(3000000 * 0.7)
    expect(JSON.stringify(p)).toContain('roundRect')
  })
})

describe('arrow5 ring of inward arrows (smartart-autoTxRot / smartart-rotation)', () => {
  it("places downArrows on a ring, first at 12 o'clock pointing down, clockwise", () => {
    const shapes = layoutDiagramFallback(DATA, {}, 9144000, 6858000, 'arrow5')
    expect(shapes).toHaveLength(5)
    const geo = JSON.stringify(shapes[0])
    expect(geo).toContain('downArrow')
    // First node: centered horizontally, at the ring top, unrotated (points at the center)
    const t0 = shapes[0]!.transform.offset
    expect(t0.x + t0.cx / 2).toBeCloseTo(9144000 / 2, -4)
    expect(shapes[0]!.transform.rot).toBe(0)
    // Second node (clockwise) sits right of center and is rotated by 360/5 degrees
    const t1 = shapes[1]!.transform.offset
    expect(t1.x + t1.cx / 2).toBeGreaterThan(9144000 / 2)
    expect(shapes[1]!.transform.rot).toBe(72 * 60000)
  })

  it('modified layout copies (arrow5#1) dispatch to the same family', () => {
    const shapes = layoutDiagramFallback(DATA, {}, 9144000, 6858000, 'arrow5#1')
    expect(JSON.stringify(shapes[0])).toContain('downArrow')
  })
})

describe('orgChart assistants (dgm:pt type="asst")', () => {
  const ASST = HIER_ASST()
  function HIER_ASST() {
    return `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="d" xmlns:a="a">
<dgm:ptLst>
<dgm:pt modelId="doc" type="doc"/>
<dgm:pt modelId="n1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Boss</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="as" type="asst"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Asst</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n2"><dgm:t><a:bodyPr/><a:p><a:r><a:t>E1</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n3"><dgm:t><a:bodyPr/><a:p><a:r><a:t>E2</a:t></a:r></a:p></dgm:t></dgm:pt>
</dgm:ptLst>
<dgm:cxnLst>
<dgm:cxn modelId="c1" srcId="doc" destId="n1" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c2" srcId="n1" destId="n2" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c3" srcId="n1" destId="n3" srcOrd="1" destOrd="0"/>
<dgm:cxn modelId="c4" srcId="n1" destId="as" srcOrd="2" destOrd="0"/>
</dgm:cxnLst></dgm:dataModel>`
  }
  it('a childless assistant gets its own row between parent and children, left of the trunk', () => {
    const shapes = layoutDiagramFallback(ASST, {}, 9144000, 6858000, 'orgChart1')
    const boxes = shapes.filter((sp: any) => sp.text) as any[]
    expect(boxes).toHaveLength(4)
    const texts = boxes.map((b) => JSON.stringify(b))
    const at = (name: string) => boxes[texts.findIndex((t) => t.includes(name))]!
    const boss = at('Boss')
    const asst = at('Asst')
    const e1 = at('E1')
    // Own row: below the boss, above the employees
    expect(asst.transform.offset.y).toBeGreaterThan(boss.transform.offset.y)
    expect(asst.transform.offset.y).toBeLessThan(e1.transform.offset.y)
    // Tucked left of the boss's trunk
    const bossCx = boss.transform.offset.x + boss.transform.offset.cx / 2
    expect(asst.transform.offset.x + asst.transform.offset.cx).toBeLessThan(bossCx)
    // Employees stay side by side (root children never hang)
    expect(at('E2').transform.offset.y).toBe(e1.transform.offset.y)
  })
})

describe('cycleMatrix family (cycle4)', () => {
  const CM_DATA = `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="d" xmlns:a="a">
<dgm:ptLst>
<dgm:pt modelId="doc" type="doc"><dgm:prSet/></dgm:pt>
<dgm:pt modelId="n1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>A1</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n2"><dgm:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></dgm:spPr><dgm:t><a:bodyPr/><a:p><a:r><a:t>B1</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n3"><dgm:t><a:bodyPr/><a:p><a:r><a:t>C1</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="n4"><dgm:t><a:bodyPr/><a:p><a:r><a:t>D1</a:t></a:r></a:p></dgm:t></dgm:pt>
<dgm:pt modelId="k1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>A2</a:t></a:r></a:p></dgm:t></dgm:pt>
</dgm:ptLst>
<dgm:cxnLst>
<dgm:cxn modelId="c1" srcId="doc" destId="n1" srcOrd="0" destOrd="0"/>
<dgm:cxn modelId="c2" srcId="doc" destId="n2" srcOrd="1" destOrd="0"/>
<dgm:cxn modelId="c3" srcId="doc" destId="n3" srcOrd="2" destOrd="0"/>
<dgm:cxn modelId="c4" srcId="doc" destId="n4" srcOrd="3" destOrd="0"/>
<dgm:cxn modelId="c5" srcId="n1" destId="k1" srcOrd="0" destOrd="0"/>
</dgm:cxnLst></dgm:dataModel>`

  it('emits four corner cards behind four pie quadrants with a center hub', () => {
    const shapes = layoutDiagramFallback(CM_DATA, {}, 6096000, 4064000, 'cycle4') as any[]
    const pies = shapes.filter((s) => s.presetGeometry === 'pie')
    const cards = shapes.filter((s) => s.presetGeometry === 'roundRect')
    const hub = shapes.filter((s) => s.presetGeometry === 'donut')
    expect(pies).toHaveLength(4)
    expect(cards).toHaveLength(4)
    expect(hub).toHaveLength(1)
    // Cards precede the wedges (PowerPoint tucks them behind the circle)
    expect(shapes.indexOf(cards[0])).toBeLessThan(shapes.indexOf(pies[0]))
    // Second node's explicit dgm:spPr fill carries through
    expect(JSON.stringify(pies[1])).toContain('ED7D31')
    // The first card carries the child bullet
    expect(JSON.stringify(cards[0])).toContain('A2')
    // The custom-filled node's card stroke follows its wedge color
    expect(JSON.stringify(cards[1])).toContain('ED7D31')
  })
})
