import { describe, it, expect } from 'vitest'
import { parseCustGeom, evalGuides } from '../src/custgeom'
import { parseSlide } from '../src/parse'

describe('evalGuides (gd guide formula evaluation)', () => {
  const g = (fmla: string, w = 200, h = 100) => evalGuides([{ name: 'x', fmla }], w, h).x

  it('val / built-in variables', () => {
    expect(g('val 50000')).toBe(50000)
    expect(evalGuides([{ name: 'x', fmla: 'val w' }], 200, 100).x).toBe(200)
    expect(g('val hc')).toBe(100)
    expect(g('val ss')).toBe(100)
    expect(g('val wd2')).toBe(100)
    expect(g('val 3cd4')).toBe(16200000)
  })

  it('*/ (mulDiv) and +- (addSub) and +/', () => {
    expect(g('*/ w 1 2')).toBe(100)
    expect(g('+- w h 50')).toBe(250)
    expect(g('+/ w h 2')).toBe(150)
  })

  it('pin clamps to range', () => {
    expect(evalGuides([{ name: 'x', fmla: 'pin 10 5 20' }], 1, 1).x).toBe(10)
    expect(evalGuides([{ name: 'x', fmla: 'pin 10 15 20' }], 1, 1).x).toBe(15)
    expect(evalGuides([{ name: 'x', fmla: 'pin 10 99 20' }], 1, 1).x).toBe(20)
  })

  it('at2 returns 1/60000 degree units', () => {
    expect(g('at2 1 1')).toBeCloseTo(2700000, 0) // 45°
    expect(g('at2 0 1')).toBeCloseTo(5400000, 0) // 90°
  })

  it('cos/sin/tan take angles in 1/60000 degree units', () => {
    expect(g('cos 100 3600000')).toBeCloseTo(50, 5) // cos60°
    expect(g('sin 100 1800000')).toBeCloseTo(50, 5) // sin30°
    expect(g('tan 100 2700000')).toBeCloseTo(100, 5) // tan45°
  })

  it('cat2/sat2/mod/sqrt/abs/min/max/?:', () => {
    expect(g('cat2 100 1 1')).toBeCloseTo(100 * Math.SQRT1_2, 4)
    expect(g('sat2 100 1 1')).toBeCloseTo(100 * Math.SQRT1_2, 4)
    expect(g('mod 3 4 0')).toBe(5)
    expect(g('sqrt 81')).toBe(9)
    expect(g('abs -7')).toBe(7)
    expect(g('min w h')).toBe(100)
    expect(g('max w h')).toBe(200)
    expect(g('?: 1 11 22')).toBe(11)
    expect(g('?: -1 11 22')).toBe(22)
  })

  it('a guide can reference earlier guides', () => {
    const env = evalGuides(
      [
        { name: 'a', fmla: 'val 40' },
        { name: 'b', fmla: '*/ a 2 1' },
      ],
      1,
      1,
    )
    expect(env.b).toBe(80)
  })
})

const spWrap = (custGeom: string, cx = 914400, cy = 914400) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Freeform 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
  custGeom +
  `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>`

describe('parseCustGeom', () => {
  it('triangle pathLst → normalized M/L/Z', () => {
    const xml = spWrap(
      `<a:custGeom><a:avLst/><a:gdLst/><a:pathLst>` +
        `<a:path w="100" h="100">` +
        `<a:moveTo><a:pt x="0" y="100"/></a:moveTo>` +
        `<a:lnTo><a:pt x="50" y="0"/></a:lnTo>` +
        `<a:lnTo><a:pt x="100" y="100"/></a:lnTo>` +
        `<a:close/></a:path></a:pathLst></a:custGeom>`,
    )
    const geo = parseCustGeom(xml, 914400, 914400)
    expect(geo?.path).toBe('M 0 1 L 0.5 0 L 1 1 Z')
  })

  it('pt references gdLst formulas (element EMU space used when path declares no w/h)', () => {
    const xml = spWrap(
      `<a:custGeom><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst>` +
        `<a:gdLst><a:gd name="gx" fmla="*/ w adj 100000"/></a:gdLst><a:pathLst>` +
        `<a:path>` +
        `<a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
        `<a:lnTo><a:pt x="gx" y="h"/></a:lnTo>` +
        `<a:close/></a:path></a:pathLst></a:custGeom>`,
      200000,
      100000,
    )
    const geo = parseCustGeom(xml, 200000, 100000)
    expect(geo?.path).toBe('M 0 0 L 0.5 1 Z')
  })

  it('quadBezTo/cubicBezTo output in order', () => {
    const xml = spWrap(
      `<a:custGeom><a:pathLst><a:path w="10" h="10">` +
        `<a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
        `<a:cubicBezTo><a:pt x="1" y="0"/><a:pt x="2" y="0"/><a:pt x="5" y="5"/></a:cubicBezTo>` +
        `<a:quadBezTo><a:pt x="8" y="10"/><a:pt x="10" y="10"/></a:quadBezTo>` +
        `</a:path></a:pathLst></a:custGeom>`,
    )
    const geo = parseCustGeom(xml, 914400, 914400)
    expect(geo?.path).toBe('M 0 0 C 0.1 0 0.2 0 0.5 0.5 Q 0.8 1 1 1')
  })

  it('arcTo converts to cubic, endpoint lands on the swept ellipse position', () => {
    // From (100,50), radii (50,50), stAng 0, sweep 90° → endpoint (50,100)
    const xml = spWrap(
      `<a:custGeom><a:pathLst><a:path w="100" h="100">` +
        `<a:moveTo><a:pt x="100" y="50"/></a:moveTo>` +
        `<a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>` +
        `</a:path></a:pathLst></a:custGeom>`,
    )
    const geo = parseCustGeom(xml, 914400, 914400)
    expect(geo?.path).toMatch(/^M 1 0\.5 C /)
    const nums = geo!
      .path!.split(' ')
      .filter((t) => Number.isFinite(Number(t)))
      .map(Number)
    expect(nums[nums.length - 2]).toBeCloseTo(0.5, 4)
    expect(nums[nums.length - 1]).toBeCloseTo(1, 4)
  })

  it('fill="none"/stroke="0" bucketed separately, multiple paths concatenated', () => {
    const xml = spWrap(
      `<a:custGeom><a:pathLst>` +
        `<a:path w="10" h="10" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="10"/></a:lnTo><a:close/></a:path>` +
        `<a:path w="10" h="10" fill="none"><a:moveTo><a:pt x="0" y="10"/></a:moveTo><a:lnTo><a:pt x="10" y="0"/></a:lnTo></a:path>` +
        `<a:path w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="0"/></a:lnTo><a:close/></a:path>` +
        `</a:pathLst></a:custGeom>`,
    )
    const geo = parseCustGeom(xml, 914400, 914400)
    expect(geo?.fillPath).toBe('M 0 0 L 1 1 Z')
    expect(geo?.strokePath).toBe('M 0 1 L 1 0')
    expect(geo?.path).toBe('M 0 0 L 1 0 Z')
  })

  it('close resets the current point to the subpath start, subsequent arcTo center is correct', () => {
    // M(50,0) L(100,0) Z → current point back to (50,0); arc center = (50,0)-50*(cos180,sin180) = (100,0)
    // Sweep -90° (180°→90°) endpoint = (100,0)+50*(cos90,sin90) = (100,50)
    const xml = spWrap(
      `<a:custGeom><a:pathLst><a:path w="100" h="100">` +
        `<a:moveTo><a:pt x="50" y="0"/></a:moveTo>` +
        `<a:lnTo><a:pt x="100" y="0"/></a:lnTo>` +
        `<a:close/>` +
        `<a:arcTo wR="50" hR="50" stAng="10800000" swAng="-5400000"/>` +
        `</a:path></a:pathLst></a:custGeom>`,
    )
    const geo = parseCustGeom(xml, 914400, 914400)
    const nums = geo!
      .path!.split(' ')
      .filter((t) => Number.isFinite(Number(t)))
      .map(Number)
    expect(nums[nums.length - 2]).toBeCloseTo(1, 4)
    expect(nums[nums.length - 1]).toBeCloseTo(0.5, 4)
  })

  it('arcTo wR≠hR: stAng/swAng are ray angles, converted via parametric angles', () => {
    // Ellipse wR=50000 hR=25000, start on the 45° ray (parametric angle 63.4349°) = center(50000,50000)+(22361,22361)
    // Sweep to the 90° ray (parametric angle also 90°) → endpoint (50000, 75000)
    const xml = spWrap(
      `<a:custGeom><a:pathLst><a:path w="100000" h="100000">` +
        `<a:moveTo><a:pt x="72361" y="72361"/></a:moveTo>` +
        `<a:arcTo wR="50000" hR="25000" stAng="2700000" swAng="2700000"/>` +
        `</a:path></a:pathLst></a:custGeom>`,
    )
    const geo = parseCustGeom(xml, 914400, 914400)
    const nums = geo!
      .path!.split(' ')
      .filter((t) => Number.isFinite(Number(t)))
      .map(Number)
    expect(nums[nums.length - 2]).toBeCloseTo(0.5, 3)
    expect(nums[nums.length - 1]).toBeCloseTo(0.75, 3)
  })

  it('multiple paths without w/h and zero element box: shared fallback coordinate space, original order preserved', () => {
    const xml =
      `<p:sp><p:spPr><a:custGeom><a:pathLst>` +
      `<a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="50" y="50"/></a:lnTo></a:path>` +
      `<a:path><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo></a:path>` +
      `</a:pathLst></a:custGeom></p:spPr></p:sp>`
    const geo = parseCustGeom(xml, 0, 0)
    expect(geo?.path).toBe('M 0 0 L 0.5 0.5 M 0 1 L 1 0')
  })

  it('no custGeom / empty pathLst → undefined', () => {
    expect(parseCustGeom('<p:sp><p:spPr/></p:sp>', 100, 100)).toBeUndefined()
    expect(
      parseCustGeom(spWrap('<a:custGeom><a:pathLst/></a:custGeom>'), 914400, 914400),
    ).toBeUndefined()
  })
})

const SLIDE_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

const TRIANGLE_CUSTGEOM =
  `<a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w="100" h="100">` +
  `<a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo>` +
  `<a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>`

describe('parseSlide attaches customGeometry', () => {
  it('top-level sp: customGeometry takes effect and the element is classified as shape', () => {
    const slideXml =
      `<?xml version="1.0"?><p:sld ${SLIDE_NS}><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      spWrap(TRIANGLE_CUSTGEOM) +
      `</p:spTree></p:cSld></p:sld>`
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const el = slide.elements[0] as any
    expect(el.type).toBe('shape')
    expect(el.presetGeometry).toBeUndefined()
    expect(el.customGeometry?.path).toBe('M 0 1 L 0.5 0 L 1 1 Z')
  })

  it('sp inside a group: custGeom parsed from source slices in order', () => {
    const grp =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>` +
      spWrap(TRIANGLE_CUSTGEOM) +
      `</p:grpSp>`
    const slideXml =
      `<?xml version="1.0"?><p:sld ${SLIDE_NS}><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      grp +
      `</p:spTree></p:cSld></p:sld>`
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const grpEl = slide.elements[0] as any
    expect(grpEl.type).toBe('group')
    expect(grpEl.children[0].customGeometry?.path).toBe('M 0 1 L 0.5 0 L 1 1 Z')
  })
})

const PIC_NS =
  SLIDE_NS + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

const picWrap = (spPrExtra: string) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture 2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
  spPrExtra +
  `</p:spPr></p:pic>`

const picSlide = (spPrExtra: string, wrap: (inner: string) => string = (s) => s) =>
  parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml:
      `<?xml version="1.0"?><p:sld ${PIC_NS}><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      wrap(picWrap(spPrExtra)) +
      `</p:spTree></p:cSld></p:sld>`,
    ctx: {},
  })

describe('parseSlide attaches picture custGeom and scene3d', () => {
  it('pic custGeom: the freeform clip path is parsed from raw bytes', () => {
    const el = picSlide(TRIANGLE_CUSTGEOM).elements[0] as any
    expect(el.type).toBe('picture')
    expect(el.customGeometry?.path).toBe('M 0 1 L 0.5 0 L 1 1 Z')
  })

  it('pic scene3d: camera preset and rot are recorded', () => {
    const el = picSlide(
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="10800000" rev="0"/></a:camera>` +
        `<a:lightRig rig="threePt" dir="t"/></a:scene3d>`,
    ).elements[0] as any
    expect(el.type).toBe('picture')
    expect(el.scene3d?.cameraPreset).toBe('orthographicFront')
    expect(el.scene3d?.cameraRot).toEqual({ lat: 0, lon: 10800000, rev: 0 })
  })

  it('pic custGeom inside a group parses from its source slice', () => {
    const grpEl = picSlide(
      TRIANGLE_CUSTGEOM,
      (inner) =>
        `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>` +
        `<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>` +
        inner +
        `</p:grpSp>`,
    ).elements[0] as any
    expect(grpEl.type).toBe('group')
    expect(grpEl.children[0].type).toBe('picture')
    expect(grpEl.children[0].customGeometry?.path).toBe('M 0 1 L 0.5 0 L 1 1 Z')
  })
})
