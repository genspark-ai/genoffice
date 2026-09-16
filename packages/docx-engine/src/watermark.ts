import {
  vmlColorHex,
  vmlFloatAnchor,
  vmlFraction,
  vmlRotationDeg,
  vmlStyleDimPx,
  vmlStyleProp,
} from './parse-vml'
import type { HfImage, PictureWatermarkInfo } from './types'
import { escapeXmlAttr } from './xml-utils'

/**
 * Text watermark support. Word implements watermarks as a VML shape with a
 * v:textpath inside the page header; the shape floats behind the body text
 * on every page that uses that header.
 */

/** namespaces the header part root needs when it carries a VML watermark */
export const WATERMARK_NS =
  ' xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
  ' xmlns:w10="urn:schemas-microsoft-com:office:word"'

/** read the watermark text from a header part; null when it has none */
export function readWatermarkText(headerXml: string): string | null {
  if (!headerXml.includes('<v:textpath')) return null
  const m = /<v:textpath[^>]*\bstring="([^"]*)"/.exec(headerXml)
  if (!m) return null
  const text = m[1]
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
  return text || null
}

/**
 * Display geometry of a WordArt watermark pict: box size, rotation, anchor,
 * fill color/opacity and the textpath font. Word stretches the glyph ink to
 * the box (fitshape), so the declared font-size is irrelevant here.
 */
export function readWatermarkShape(pictXml: string): HfImage | null {
  const text = readWatermarkText(pictXml)
  if (!text) return null
  const body = pictXml.replace(/<v:shapetype\b[\s\S]*?<\/v:shapetype>/g, '')
  const shapeTag = /<v:shape\b[^>]*>/.exec(body)?.[0]
  if (!shapeTag) return null
  const attr = (tag: string, key: string): string | undefined =>
    new RegExp(`\\s${key}="([^"]*)"`).exec(tag)?.[1]
  const style = attr(shapeTag, 'style') ?? ''
  const widthPx = vmlStyleDimPx(style, 'width')
  const heightPx = vmlStyleDimPx(style, 'height')
  if (!widthPx || !heightPx) return null
  const fillTag = /<v:fill\b[^>]*>/.exec(body)?.[0]
  const tpTag = /<v:textpath\b[^>]*\bstring="[^>]*>/.exec(body)?.[0] ?? ''
  const tpStyle = (attr(tpTag, 'style') ?? '').replace(/&quot;/g, '"').replace(/&amp;/g, '&')
  const family = vmlStyleProp(tpStyle, 'font-family')?.replace(/^"|"$/g, '')
  const img: HfImage = {
    dataUrl: '',
    widthPx,
    heightPx,
    floating: true,
    wordArt: {
      text,
      colorHex:
        attr(shapeTag, 'filled') === 'f'
          ? 'FFFFFF'
          : (vmlColorHex(attr(shapeTag, 'fillcolor')) ?? '000000'),
      opacity: Math.max(
        0,
        Math.min(1, (fillTag ? vmlFraction(attr(fillTag, 'opacity')) : undefined) ?? 1),
      ),
      ...(family ? { fontFamily: family } : {}),
      ...(/font-weight:\s*bold/.test(tpStyle) ? { bold: true } : {}),
      ...(/font-style:\s*italic/.test(tpStyle) ? { italic: true } : {}),
    },
  }
  if (/z-index:\s*-/.test(style)) img.behind = true
  const rot = vmlRotationDeg(style)
  if (rot != null) img.rotationDeg = rot
  vmlFloatAnchor(style, img)
  return img
}

/** Text watermark as written by set_watermark / the Design tab; a bare string is the text alone. */
export interface WatermarkSpec {
  text: string
  fontFamily?: string
  /** hex without '#' (default silver) */
  colorHex?: string
  /** 0..1 fill opacity (default 0.5) */
  opacity?: number
  /** rotated 315 degrees like Word's diagonal layout (default true) */
  diagonal?: boolean
  bold?: boolean
  italic?: boolean
}

/** Picture watermark as written by set_watermark / Word's Design tab. */
export interface PictureWatermarkSpec {
  image: {
    /** raw image bytes, base64 encoded */
    base64: string
    mime: 'image/png' | 'image/jpeg' | 'image/gif'
    widthPx: number
    heightPx: number
  }
  /** percent of the image's natural size; absent = fit inside the page margin box */
  scale?: number
  /** Word's washout (faded) look (default true) */
  washout?: boolean
}

export type Watermark = string | WatermarkSpec | PictureWatermarkSpec

export function isPictureWatermark(wm: Watermark): wm is PictureWatermarkSpec {
  return typeof wm === 'object' && 'image' in wm
}

const PICTURE_WATERMARK_ID =
  /<v:shape\b[^>]*\bid="(?:WordPictureWatermark|PowerPlusWaterMarkObject)/

/** header part child that carries the watermark (Word wraps the paragraph in a Watermarks-gallery sdt) */
export function isWatermarkChild(child: { name: string; xml: string }): boolean {
  if (child.name !== 'w:p' && child.name !== 'w:sdt') return false
  if (child.xml.includes('<v:textpath')) return true
  return (
    child.xml.includes('<v:imagedata') &&
    (PICTURE_WATERMARK_ID.test(child.xml) ||
      child.xml.includes('<w:docPartGallery w:val="Watermarks"/>'))
  )
}

export type { PictureWatermarkInfo }

/** a VML shape written as a picture watermark (Word's or ours) */
export function isPictureWatermarkShape(xml: string): boolean {
  return PICTURE_WATERMARK_ID.test(xml)
}

/**
 * The header image a pending (unsaved) picture watermark will parse back as,
 * so the page preview can draw it exactly where the save path will put it.
 */
export function pictureWatermarkPreviewImage(
  spec: PictureWatermarkSpec,
  marginBox: { widthPt: number; heightPt: number } | null,
): HfImage {
  const box = pictureWatermarkBoxPt(spec, marginBox)
  return {
    dataUrl: `data:${spec.image.mime};base64,${spec.image.base64}`,
    widthPx: Math.round((box.widthPt * 96) / 72),
    heightPx: Math.round((box.heightPt * 96) / 72),
    floating: true,
    behind: true,
    watermark: true,
    posH: 'center',
    posV: 'center',
    posHRel: 'margin',
    posVRel: 'margin',
    ...(spec.washout !== false ? { washout: { gain: 0.3, blackLevel: 0.35 } } : {}),
  }
}

const PT_PER_UNIT: Record<string, number> = {
  pt: 1,
  in: 72,
  cm: 72 / 2.54,
  mm: 72 / 25.4,
  px: 0.75,
}

function vmlStyleDimPt(style: string, key: 'width' | 'height'): number {
  const m = new RegExp(`(?:^|;)\\s*${key}:\\s*(-?[\\d.]+)(pt|in|cm|mm|px)?`).exec(style)
  if (!m) return 0
  return Math.round(parseFloat(m[1]!) * (PT_PER_UNIT[m[2] ?? 'pt'] ?? 1) * 100) / 100
}

/** the picture watermark of a header part; null when it has none */
export function readPictureWatermark(headerXml: string): PictureWatermarkInfo | null {
  for (const m of headerXml.matchAll(/<v:shape\b[^>]*>[\s\S]*?<\/v:shape>/g)) {
    const shape = m[0]
    if (!PICTURE_WATERMARK_ID.test(shape) && !isWatermarkChild({ name: 'w:p', xml: shape }))
      continue
    const rId = /<v:imagedata[^>]*\br:id="([^"]+)"/.exec(shape)?.[1]
    if (!rId) continue
    const style = /\sstyle="([^"]*)"/.exec(shape)?.[1] ?? ''
    return {
      rId,
      widthPt: vmlStyleDimPt(style, 'width'),
      heightPt: vmlStyleDimPt(style, 'height'),
      washout: /\sgain="/.test(shape),
    }
  }
  return null
}

/** Box the picture watermark occupies, in points: explicit scale, or fitted inside the margin box. */
export function pictureWatermarkBoxPt(
  spec: PictureWatermarkSpec,
  marginBox: { widthPt: number; heightPt: number } | null,
): { widthPt: number; heightPt: number } {
  const naturalW = (spec.image.widthPx * 72) / 96
  const naturalH = (spec.image.heightPx * 72) / 96
  let factor: number
  if (spec.scale !== undefined) factor = spec.scale / 100
  else {
    // Letter with 1in margins when the section has no usable page box
    const box = marginBox ?? { widthPt: 468, heightPt: 648 }
    factor = Math.min(box.widthPt / naturalW, box.heightPt / naturalH)
  }
  return {
    widthPt: Math.max(1, Math.round(naturalW * factor * 100) / 100),
    heightPt: Math.max(1, Math.round(naturalH * factor * 100) / 100),
  }
}

/**
 * The picture watermark Word generates (shapetype 75 = picture frame) in the
 * Watermarks building-block sdt, centered on the margin box behind the body.
 * `rId` is the image relationship in the header part's own rels.
 */
export function pictureWatermarkParagraphXml(
  spec: PictureWatermarkSpec,
  rId: string,
  marginBox: { widthPt: number; heightPt: number } | null,
): string {
  const box = pictureWatermarkBoxPt(spec, marginBox)
  const washout = spec.washout !== false ? ' gain="19661f" blacklevel="22938f"' : ''
  const shapetype =
    '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t"' +
    ' path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">' +
    '<v:stroke joinstyle="miter"/>' +
    '<v:formulas>' +
    '<v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/>' +
    '<v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/>' +
    '<v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/>' +
    '<v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/>' +
    '</v:formulas>' +
    '<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>' +
    '<o:lock v:ext="edit" aspectratio="t"/>' +
    '</v:shapetype>'
  const shape =
    '<v:shape id="WordPictureWatermark1" o:spid="_x0000_s2050" type="#_x0000_t75"' +
    ` style="position:absolute;margin-left:0;margin-top:0;width:${box.widthPt}pt;height:${box.heightPt}pt;` +
    'z-index:-251657216;' +
    'mso-position-horizontal:center;mso-position-horizontal-relative:margin;' +
    'mso-position-vertical:center;mso-position-vertical-relative:margin"' +
    ' o:allowincell="f">' +
    `<v:imagedata r:id="${escapeXmlAttr(rId)}" o:title="watermark"${washout}/>` +
    '</v:shape>'
  return (
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Watermarks"/><w:docPartUnique/></w:docPartObj></w:sdtPr>' +
    `<w:sdtContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:pict>${shapetype}${shape}</w:pict></w:r></w:p></w:sdtContent></w:sdt>`
  )
}

/**
 * The gray text watermark Word generates (shapetype 136 = text-on-path),
 * wrapped in the Watermarks building-block sdt so Word's Design tab
 * recognizes and replaces it. Lives as the first child of the header part.
 */
export function watermarkParagraphXml(spec: string | WatermarkSpec): string {
  const wm = typeof spec === 'string' ? { text: spec } : spec
  const diagonal = wm.diagonal !== false
  const box = diagonal
    ? 'width:412.4pt;height:247.45pt;rotation:315;'
    : 'width:527.85pt;height:131.95pt;'
  const opacity = Math.max(0, Math.min(1, wm.opacity ?? 0.5))
  const fill = wm.colorHex ? `#${wm.colorHex.replace(/^#/, '')}` : 'silver'
  const font = escapeXmlAttr(wm.fontFamily ?? 'DengXian')
  const tpStyle =
    `font-family:&quot;${font}&quot;;font-size:1pt` +
    (wm.bold ? ';font-weight:bold' : '') +
    (wm.italic ? ';font-style:italic' : '')
  const shape =
    '<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t136"' +
    ` style="position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;${box}` +
    'z-index:-251656192;' +
    'mso-position-horizontal:center;mso-position-horizontal-relative:margin;' +
    'mso-position-vertical:center;mso-position-vertical-relative:margin"' +
    ` o:allowincell="f" fillcolor="${escapeXmlAttr(fill)}" stroked="f">` +
    `<v:fill opacity="${opacity}"/>` +
    `<v:textpath style="${tpStyle}" string="${escapeXmlAttr(wm.text)}"/>` +
    '</v:shape>'
  const shapetype =
    '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800"' +
    ' path="m@7,l@8,m@5,21600l@6,21600e">' +
    '<v:formulas>' +
    '<v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>' +
    '<v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>' +
    '<v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>' +
    '<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>' +
    '<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/>' +
    '</v:formulas>' +
    '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"' +
    ' o:connectangles="270,180,90,0"/>' +
    '<v:textpath on="t" fitshape="t"/>' +
    '<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>' +
    '<o:lock v:ext="edit" text="t" shapetype="t"/>' +
    '</v:shapetype>'
  return (
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Watermarks"/><w:docPartUnique/></w:docPartObj></w:sdtPr>' +
    `<w:sdtContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:pict>${shapetype}${shape}</w:pict></w:r></w:p></w:sdtContent></w:sdt>`
  )
}
