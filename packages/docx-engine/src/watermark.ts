import {
  vmlColorHex,
  vmlFloatAnchor,
  vmlFraction,
  vmlRotationDeg,
  vmlStyleDimPx,
  vmlStyleProp,
} from './parse-vml'
import type { HfImage } from './types'
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

/**
 * The diagonal gray text watermark paragraph Word generates (shapetype 136 =
 * text-on-path). Lives as the first paragraph of the header part.
 */
export function watermarkParagraphXml(text: string): string {
  const shape =
    '<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t136"' +
    ' style="position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;' +
    'width:412.4pt;height:247.45pt;rotation:315;z-index:-251656192;' +
    'mso-position-horizontal:center;mso-position-horizontal-relative:margin;' +
    'mso-position-vertical:center;mso-position-vertical-relative:margin"' +
    ' o:allowincell="f" fillcolor="silver" stroked="f">' +
    '<v:fill opacity=".5"/>' +
    `<v:textpath style="font-family:&quot;DengXian&quot;;font-size:1pt" string="${escapeXmlAttr(text)}"/>` +
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
  return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:pict>${shapetype}${shape}</w:pict></w:r></w:p>`
}
