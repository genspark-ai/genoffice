/**
 * 3D models stored the way PowerPoint stores them: an mc:AlternateContent whose Choice is a
 * p:graphicFrame holding am3d:model3d ([MS-ODRAWXML] 2.31, the 2017/model3d namespace) and
 * whose Fallback is a p:pic of the model's latest rendering, for applications without 3D.
 * The editor treats the model as that picture (move, resize, thumbnails, export) and keeps
 * the Choice in step: frame, rotation and the rendering it caches (am3d:raster).
 */
import type { GlbBounds } from './glb-bounds'
import type { EmuRect } from './types'
import { creationIdXml, escapeXmlAttr } from './xml-utils'

export const AM3D_NS = 'http://schemas.microsoft.com/office/drawing/2017/model3d'
export const MODEL3D_REL_TYPE = 'http://schemas.microsoft.com/office/2017/06/relationships/model3d'
/** PowerPoint's own spelling, with a dot. */
export const GLB_CONTENT_TYPE = 'model/gltf.binary'
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

/** Denominator of am3d's ratios (meterPerModelUnit, illuminance, intensity, scale). */
export const MODEL3D_RATIO_ONE = 1_000_000
/** am3d lengths are EMUs: 36,000,000 per meter. */
export const MODEL3D_EMU_PER_METER = 36_000_000
/** PowerPoint's default camera: perspective, 45° field of view (60,000ths of a degree). */
export const MODEL3D_FOV = 2_700_000

/** am3d:rot, in 60,000ths of a degree, applied to the model about its x, y and z axes. */
export interface Model3dRotation {
  ax: number
  ay: number
  az: number
}

export const NO_ROTATION: Model3dRotation = { ax: 0, ay: 0, az: 0 }

/** How PowerPoint frames a model: scaled to one meter across, centred, the camera backed off to fit it. */
export function model3dFraming(bounds: GlbBounds | null) {
  const size = bounds?.size ?? [1, 1, 1]
  const center = bounds?.center ?? [0, 0, 0]
  const maxExtent = Math.max(...size) || 1
  const radius = Math.hypot(...size.map((s) => s / 2)) / maxExtent || 1
  const cameraZ = radius / Math.sin(((MODEL3D_FOV / 60000) * Math.PI) / 360)
  return {
    meterPerModelUnit:
      maxExtent >= 1
        ? { n: MODEL3D_RATIO_ONE, d: Math.round(maxExtent * MODEL3D_RATIO_ONE) }
        : { n: Math.round(MODEL3D_RATIO_ONE / maxExtent), d: MODEL3D_RATIO_ONE },
    preTrans: center.map((c) => Math.round((-c / maxExtent) * MODEL3D_EMU_PER_METER)) as [
      number,
      number,
      number,
    ],
    cameraZ: Math.round(cameraZ * MODEL3D_EMU_PER_METER),
  }
}

/**
 * PowerPoint's default lighting for an inserted model: a grey ambient light and three point
 * lights (color in scRGB 100,000ths, intensity n/1,000,000, position in EMUs).
 */
export const MODEL3D_LIGHTS = {
  ambient: { rgb: [50000, 50000, 50000], illuminance: 500000 },
  points: [
    { rgb: [100000, 75000, 50000], intensity: 9765625, pos: [21959998, 70920001, 16344003] },
    { rgb: [40000, 60000, 95000], intensity: 12250000, pos: [-37964106, 51130435, 57631972] },
    { rgb: [86837, 72700, 100000], intensity: 3125000, pos: [-37739122, 58056624, -34769649] },
  ],
} as const

const scrgb = ([r, g, b]: readonly number[]) => `<a:scrgbClr r="${r}" g="${g}" b="${b}"/>`

const rotXml = ({ ax, ay, az }: Model3dRotation) =>
  `<am3d:rot ax="${Math.round(ax)}" ay="${Math.round(ay)}" az="${Math.round(az)}"/>`

export interface Model3dXmlOptions {
  id: number
  name: string
  offset: EmuRect
  modelRId: string
  posterRId: string
  bounds: GlbBounds | null
  rot?: Model3dRotation
}

/** The slide's mc:AlternateContent for a new model. */
export function model3dXml(o: Model3dXmlOptions): string {
  const { x, y, cx, cy } = o.offset
  const framing = model3dFraming(o.bounds)
  const [dx, dy, dz] = framing.preTrans
  const name = escapeXmlAttr(o.name)
  const creationId = creationIdXml()
  const model =
    `<am3d:model3d r:embed="${o.modelRId}">` +
    `<am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></am3d:spPr>' +
    `<am3d:camera><am3d:pos x="0" y="0" z="${framing.cameraZ}"/>` +
    `<am3d:up dx="0" dy="${MODEL3D_EMU_PER_METER}" dz="0"/><am3d:lookAt x="0" y="0" z="0"/>` +
    `<am3d:perspective fov="${MODEL3D_FOV}"/></am3d:camera>` +
    `<am3d:trans><am3d:meterPerModelUnit n="${framing.meterPerModelUnit.n}" d="${framing.meterPerModelUnit.d}"/>` +
    `<am3d:preTrans dx="${dx}" dy="${dy}" dz="${dz}"/>` +
    `<am3d:scale><am3d:sx n="${MODEL3D_RATIO_ONE}" d="${MODEL3D_RATIO_ONE}"/><am3d:sy n="${MODEL3D_RATIO_ONE}" d="${MODEL3D_RATIO_ONE}"/>` +
    `<am3d:sz n="${MODEL3D_RATIO_ONE}" d="${MODEL3D_RATIO_ONE}"/></am3d:scale>` +
    rotXml(o.rot ?? NO_ROTATION) +
    '<am3d:postTrans dx="0" dy="0" dz="0"/></am3d:trans>' +
    `<am3d:raster rName="Office3DRenderer" rVer="16.0.8326"><am3d:blip r:embed="${o.posterRId}"/></am3d:raster>` +
    `<am3d:objViewport viewportSz="${Math.max(cx, cy)}"/>` +
    `<am3d:ambientLight><am3d:clr>${scrgb(MODEL3D_LIGHTS.ambient.rgb)}</am3d:clr>` +
    `<am3d:illuminance n="${MODEL3D_LIGHTS.ambient.illuminance}" d="${MODEL3D_RATIO_ONE}"/></am3d:ambientLight>` +
    MODEL3D_LIGHTS.points
      .map(
        (l) =>
          `<am3d:ptLight rad="0"><am3d:clr>${scrgb(l.rgb)}</am3d:clr>` +
          `<am3d:intensity n="${l.intensity}" d="${MODEL3D_RATIO_ONE}"/>` +
          `<am3d:pos x="${l.pos[0]}" y="${l.pos[1]}" z="${l.pos[2]}"/></am3d:ptLight>`,
      )
      .join('') +
    '</am3d:model3d>'
  return (
    `<mc:AlternateContent xmlns:mc="${MC_NS}">` +
    `<mc:Choice xmlns:am3d="${AM3D_NS}" Requires="am3d">` +
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${o.id}" name="${name}">${creationId}</p:cNvPr>` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${AM3D_NS}">${model}</a:graphicData></a:graphic></p:graphicFrame>` +
    '</mc:Choice><mc:Fallback>' +
    `<p:pic><p:nvPicPr><p:cNvPr id="${o.id}" name="${name}">${creationId}</p:cNvPr>` +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    `<p:blipFill><a:blip r:embed="${o.posterRId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' +
    '</mc:Fallback></mc:AlternateContent>'
  )
}

const CHOICE_FRAME =
  /(<mc:Choice\b[^>]*>\s*<p:graphicFrame\b[\s\S]*?)<p:xfrm\b[^>]*>[\s\S]*?<\/p:xfrm>/
const FALLBACK_XFRM = /<mc:Fallback\b[^>]*>[\s\S]*?(<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>)/

/**
 * Copies the Fallback picture's frame, which the editor moves and resizes, to the Choice:
 * the graphicFrame's p:xfrm, the model's own extent and its viewport.
 */
export function syncModel3dFrame(xml: string): string {
  const fallback = FALLBACK_XFRM.exec(xml)?.[1]
  const choice = CHOICE_FRAME.exec(xml)
  if (!fallback || !choice) return xml
  const pXfrm = fallback.replace(/^<a:xfrm\b/, '<p:xfrm').replace(/<\/a:xfrm>$/, '</p:xfrm>')
  const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(fallback)
  let out = xml.replace(CHOICE_FRAME, (_m, head: string) => head + pXfrm)
  if (ext) {
    const [, cx, cy] = ext
    out = out
      .replace(
        /(<am3d:spPr>\s*<a:xfrm\b[^>]*>[\s\S]*?)<a:ext cx="\d+" cy="\d+"\/>/,
        (_m, head: string) => `${head}<a:ext cx="${cx}" cy="${cy}"/>`,
      )
      .replace(
        /<am3d:objViewport viewportSz="\d+"/,
        `<am3d:objViewport viewportSz="${Math.max(Number(cx), Number(cy))}"`,
      )
  }
  return out
}

/** The model's rotation, as am3d:rot holds it (absent axes are 0). */
export function readModel3dRotation(xml: string): Model3dRotation {
  const tag = /<am3d:rot\b[^>]*\/?>/.exec(xml)?.[0] ?? ''
  const axis = (name: string) => Number(new RegExp(`\\b${name}="(-?\\d+)"`).exec(tag)?.[1] ?? 0)
  return { ax: axis('ax'), ay: axis('ay'), az: axis('az') }
}

export function setModel3dRotationXml(xml: string, rot: Model3dRotation): string {
  if (/<am3d:rot\b[^>]*\/>/.test(xml)) return xml.replace(/<am3d:rot\b[^>]*\/>/, rotXml(rot))
  // am3d:rot sits after scale and before postTrans in CT_Model3DTransform
  return xml.replace(/<am3d:postTrans\b|<\/am3d:trans>/, (next) => rotXml(rot) + next)
}
