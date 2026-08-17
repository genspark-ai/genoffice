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

  it('orgChart hangs a branch whose children are all leaves (rows below, trunk + stubs)', () => {
    const shapes = layoutDiagramFallback(HIER_DATA, {}, 9144000, 6858000, 'orgChart1')
    // 3 boxes + trunk + 2 stubs
    expect(shapes.length).toBe(6)
    const boxes = shapes.filter((sp: any) => sp.text)
    expect(boxes).toHaveLength(3)
    const [a, b1, b2] = boxes as any[]
    expect(a.transform.offset.y).toBeLessThan(b1.transform.offset.y)
    // hanging: B2 on its own row below B1
    expect(b2.transform.offset.y).toBeGreaterThan(b1.transform.offset.y)
  })

  it('orgChart hangs a branch whose children are all leaves when nested deeper', () => {
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
    // hanging children of n1 (B1, B2 leaves) stack on separate rows
    const [, , b1, b2] = boxes
    expect(b2.transform.offset.y).toBeGreaterThan(b1.transform.offset.y)
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
