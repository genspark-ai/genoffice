import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'

const SLIDE_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

function spWith(spPrInner: string): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    spPrInner +
    `</p:spPr></p:sp>`
  )
}

function parseSp(spPrInner: string): any {
  const slideXml =
    `<?xml version="1.0"?><p:sld ${SLIDE_NS}><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    spWith(spPrInner) +
    `</p:spTree></p:cSld></p:sld>`
  return parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} }).elements[0]
}

describe('parseSlide attaches scene3d', () => {
  it('camera preset + light rig + sp3d extrusion', () => {
    const el = parseSp(
      `<a:scene3d><a:camera prst="isometricRightUp"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>` +
        `<a:sp3d extrusionH="457200"/>`,
    )
    expect(el.scene3d).toEqual({
      cameraPreset: 'isometricRightUp',
      lightRig: 'threePt',
      lightDir: 't',
      extrusionEmu: 457200,
    })
  })

  it('camera/light rotation overrides, z shift, material, extrusion color', () => {
    const el = parseSp(
      `<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="13499976" rev="0"/></a:camera>` +
        `<a:lightRig rig="twoPt" dir="t"><a:rot lat="0" lon="0" rev="5400000"/></a:lightRig></a:scene3d>` +
        `<a:sp3d z="228600" extrusionH="914400" prstMaterial="legacyWireframe">` +
        `<a:extrusionClr><a:srgbClr val="00FFFF"/></a:extrusionClr></a:sp3d>`,
    )
    expect(el.scene3d).toEqual({
      cameraPreset: 'orthographicFront',
      cameraRot: { lat: 0, lon: 13499976, rev: 0 },
      lightRig: 'twoPt',
      lightDir: 't',
      lightRot: { lat: 0, lon: 0, rev: 5400000 },
      extrusionEmu: 914400,
      zEmu: 228600,
      extrusionColor: '#00FFFF',
      material: 'legacyWireframe',
    })
  })

  it('no scene3d → field absent; sp3d without a camera is ignored', () => {
    expect(parseSp('').scene3d).toBeUndefined()
    expect(parseSp('<a:sp3d extrusionH="914400"/>').scene3d).toBeUndefined()
  })
})
