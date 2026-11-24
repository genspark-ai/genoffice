import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { NodeSelection } from '@tiptap/pm/state'
import { NEW_PICTURE_SVG_SUFFIX } from '@genoffice/docx-engine'
import { getActiveSubEditor } from './active-editor'
import { heldPictureToken, heldPictureXml } from './convert'
import {
  pictureVectorAltOf,
  pictureVectorAttr,
  rememberPictureVector,
  type PictureVector,
} from './picture-vectors'
import { rasteriseSvg, svgBase64Of, svgDataUrl } from './svg-raster'

const UNPAINTED = /^(?:none|transparent)$/i

/**
 * The SVG with its shapes filled in one colour (hex without '#'), or with no
 * fill for null: Word's Graphics Fill. A shape that says it has no fill keeps
 * none, and strokes are left alone.
 */
export function recolorSvg(svg: string, hex: string | null): string | null {
  const parsed = new DOMParser().parseFromString(svg, 'image/svg+xml')
  const root = parsed.documentElement
  if (root.localName !== 'svg' || parsed.getElementsByTagName('parsererror').length > 0) return null
  const fill = hex ? `#${hex}` : 'none'
  root.setAttribute('fill', fill)
  for (const el of Array.from(root.querySelectorAll('*'))) {
    const own = el.getAttribute('fill')
    if (own !== null && !UNPAINTED.test(own.trim())) el.setAttribute('fill', fill)
    const style = el.getAttribute('style')
    if (style && /(?:^|;)\s*fill\s*:/i.test(style))
      el.setAttribute(
        'style',
        style.replace(/((?:^|;)\s*fill\s*:\s*)([^;]+)/gi, (all, head: string, value: string) =>
          UNPAINTED.test(value.trim()) ? all : head + fill,
        ),
      )
  }
  return new XMLSerializer().serializeToString(root)
}

const HEX = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i

function hexOf(value: string): string | null {
  const match = HEX.exec(value.trim())
  if (!match) return /^black$/i.test(value.trim()) ? '000000' : null
  const digits = match[1].length === 3 ? match[1].replace(/./g, '$&$&') : match[1]
  return digits.toUpperCase()
}

/**
 * The colour an SVG's shapes are filled with (hex without '#'), null for no
 * fill: the root's fill, else the first shape that names one, else SVG's
 * black. Undefined for a paint that is not a plain colour.
 */
export function svgFillOf(svg: string): string | null | undefined {
  const parsed = new DOMParser().parseFromString(svg, 'image/svg+xml')
  const root = parsed.documentElement
  if (root.localName !== 'svg' || parsed.getElementsByTagName('parsererror').length > 0)
    return undefined
  for (const el of [root, ...Array.from(root.querySelectorAll('*'))]) {
    const own =
      el.getAttribute('fill') ??
      /(?:^|;)\s*fill\s*:\s*([^;]+)/i.exec(el.getAttribute('style') ?? '')?.[1] ??
      null
    if (own === null) continue
    if (UNPAINTED.test(own.trim())) {
      if (el === root) return null
      continue
    }
    return hexOf(own) ?? undefined
  }
  return '000000'
}

const BLIP_EMBED = /(<a:blip\b[^>]*\br:embed=")[^"]*(")/
const SVG_BLIP_EMBED = /(<\w+:svgBlip\b[^>]*\br:embed=")[^"]*(")/

/**
 * A picture's drawing pointed at bytes the save has yet to write: everything
 * else in it (anchor, wrap, crop, effects) stays as the file has it. Null for
 * a drawing that is not one picture with an SVG.
 */
export function retargetGraphicXml(xml: string, token: string): string | null {
  if ((xml.match(/<pic:pic[\s>]/g) ?? []).length !== 1) return null
  if (!BLIP_EMBED.test(xml) || !SVG_BLIP_EMBED.test(xml)) return null
  return xml
    .replace(BLIP_EMBED, `$1${token}$2`)
    .replace(SVG_BLIP_EMBED, `$1${token}${NEW_PICTURE_SVG_SUFFIX}$2`)
}

interface SelectedGraphic {
  editor: Editor
  pos: number
  node: PmNode
  bitmapUrl: string
  vector: PictureVector
  widthPx: number
  heightPx: number
}

/**
 * The selected picture when the page draws it from an SVG. It may be in a
 * header, footer or text box: the editor that has the focus holds it.
 */
function selectedGraphic(main: Editor): SelectedGraphic | null {
  const editor = getActiveSubEditor() ?? main
  const sel = editor.state.selection
  if (!(sel instanceof NodeSelection)) return null
  const { node } = sel
  const run = node.type.name === 'docInlineImage'
  if (!run && !(node.type.name === 'docProtected' && node.attrs.blockType === 'image')) return null
  if (run && node.attrs.rule) return null
  const bitmapUrl = String((run ? node.attrs.dataUrl : node.attrs.imageDataUrl) ?? '')
  const vector = (run ? node.attrs.vector : node.attrs.imageVector) as PictureVector | null
  const widthPx = Number(run ? node.attrs.widthPx : node.attrs.imageWidthPx)
  const heightPx = Number(run ? node.attrs.heightPx : node.attrs.imageHeightPx)
  if (!vector?.url || vector.for !== bitmapUrl || !(widthPx > 0) || !(heightPx > 0)) return null
  return { editor, pos: sel.from, node, bitmapUrl, vector, widthPx, heightPx }
}

/** the selected picture is a graphic: an icon or another picture kept as SVG */
export const graphicSelected = (editor: Editor): boolean => selectedGraphic(editor) !== null

const unescapeAttr = (value: string): string =>
  value
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')

/** the alternative text the picture is saved with, which a redrawn picture has to keep */
function altTextOf(graphic: SelectedGraphic): string | undefined {
  if (graphic.node.attrs.altText != null) return String(graphic.node.attrs.altText)
  const base64 = /^data:[^;]+;base64,(.*)$/s.exec(graphic.bitmapUrl)?.[1]
  const kept = base64 ? pictureVectorAltOf(base64) : undefined
  if (kept) return kept
  if (graphic.node.attrs.imageAlt) return String(graphic.node.attrs.imageAlt)
  const descr = /<wp:docPr\b[^>]*\sdescr="([^"]*)"/.exec(String(graphic.node.attrs.xml ?? ''))
  return descr ? unescapeAttr(descr[1]) : undefined
}

/** the SVG behind a URL: a data URL is decoded here (the page may not fetch one), a part of the open file is fetched */
async function svgTextOf(url: string): Promise<string> {
  const data = /^data:image\/svg\+xml(;base64)?,(.*)$/s.exec(url)
  if (!data) return (await fetch(url)).text()
  if (!data[1]) return decodeURIComponent(data[2])
  const bytes = Uint8Array.from(atob(data[2]), (ch) => ch.charCodeAt(0))
  return new TextDecoder().decode(bytes)
}

/**
 * What a picture run saves once its bytes are others: a drawing that came with
 * the file is pointed at them, and one the save draws itself is drawn anew. In
 * a header, footer or text box the run always carries its drawing. Null when
 * the drawing cannot take other bytes.
 */
function runDrawingFor(graphic: SelectedGraphic, dataUrl: string): string | null {
  const { node, editor } = graphic
  const xml = String(node.attrs.xml ?? '')
  const retargeted = xml ? retargetGraphicXml(xml, heldPictureToken()) : null
  if (retargeted) return retargeted
  const held = !editor.schema.nodes.docProtected
  // an anchored drawing of the file holds its place on the page: nothing here redraws that
  if (xml && (held || node.attrs.wrap)) return null
  if (!held) return ''
  return (
    heldPictureXml({
      dataUrl,
      widthPx: graphic.widthPx,
      heightPx: graphic.heightPx,
      altText: altTextOf(graphic),
    }) || null
  )
}

/** the colour the selected graphic is filled with: hex without '#', null for no fill, undefined when unknown */
export async function selectedGraphicFill(editor: Editor): Promise<string | null | undefined> {
  const graphic = selectedGraphic(editor)
  if (!graphic) return undefined
  try {
    return svgFillOf(await svgTextOf(graphic.vector.url))
  } catch {
    return undefined
  }
}

/**
 * Graphics Fill: the selected graphic is redrawn in the colour (hex without
 * '#', null for no fill). The SVG is what changes; its fallback bitmap is
 * drawn again from it, and the save writes both.
 */
export async function setSelectedGraphicFill(main: Editor, hex: string | null): Promise<boolean> {
  const graphic = selectedGraphic(main)
  // a click on the ribbon ends a text box's editing; its picture is still the one selected
  if (!graphic || !main.isEditable) return false
  const { editor } = graphic
  let svg: string | null
  try {
    svg = recolorSvg(await svgTextOf(graphic.vector.url), hex)
    if (!svg) return false
    const dataUrl = await rasteriseSvg(svg, graphic.widthPx, graphic.heightPx)
    const base64 = dataUrl.slice(dataUrl.indexOf(',') + 1)
    rememberPictureVector(base64, svgBase64Of(svg), altTextOf(graphic))
    // the picture may have moved or gone while it was drawn
    const { state } = editor
    if (state.doc.nodeAt(graphic.pos) !== graphic.node) return false
    const vector = pictureVectorAttr(dataUrl, svgDataUrl(svg))
    const { attrs } = graphic.node
    const run = graphic.node.type.name === 'docInlineImage'
    const xml = run ? runDrawingFor(graphic, dataUrl) : ''
    if (xml === null) return false
    const next = run
      ? { dataUrl, vector, xml }
      : {
          imageDataUrl: dataUrl,
          imageVector: vector,
          imageLink: null,
          ...(attrs.docxIndex != null
            ? { imageReplace: { base64, mime: 'image/png' } }
            : { genImage: { ...attrs.genImage, base64, mime: 'image/png' } }),
        }
    const tr = state.tr.setNodeMarkup(graphic.pos, undefined, { ...attrs, ...next })
    editor.view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, graphic.pos)))
    // a text box hands its content to the document
    if (editor !== main && !editor.schema.nodes.hfField)
      window.dispatchEvent(new Event('ai-docs-commit-tables'))
    return true
  } catch {
    return false
  }
}
