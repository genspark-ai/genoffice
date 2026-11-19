import type { NewInkImage } from './types'
import { escapeXmlAttr } from './xml-utils'

/**
 * Ink annotations (freehand strokes). Each saved annotation is a floating picture —
 * a run-level <w:drawing><wp:anchor> injected into its anchor paragraph:
 *
 * - wp:positionV relativeFrom="paragraph": the drawing moves with the
 *   paragraph when content above it reflows.
 * - wp:positionH relativeFrom="column": offsets are measured from the text
 *   column edge, independent of the paragraph's own indentation.
 * - behindDoc="0" + wp:wrapNone: floats in front of the text without
 *   affecting layout.
 * - wp:docPr name starts with "aidocs-ink" so we can recognize our own
 *   annotations when the file is reopened; descr carries the editor's
 *   stroke vectors (opaque payload) so the layer stays re-editable.
 *
 * Word renders these as ordinary floating pictures; only our editor restores
 * them into live strokes.
 */

export const INK_NAME_PREFIX = 'aidocs-ink'

/** media filename prefix for ink PNGs (distinct from inline-image aidocsN) */
export const INK_MEDIA_PREFIX = 'aidocsink'

/** matches a document.xml.rels entry that targets an ink media part */
export const INK_REL_RE = new RegExp(
  `<Relationship [^>]*Target="media/${INK_MEDIA_PREFIX}\\d+\\.png"[^>]*/>`,
  'g',
)

/** matches an ink media part's zip path */
export const INK_MEDIA_PATH_RE = new RegExp(`^word/media/${INK_MEDIA_PREFIX}\\d+\\.png$`)

const EMU_PER_PX = 9525

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/**
 * The anchored-picture run for one ink annotation. Namespaces are declared
 * inline so the fragment is valid regardless of what the document root
 * declares (w: is always present).
 */
export function anchoredInkRunXml(
  ink: Pick<NewInkImage, 'widthPx' | 'heightPx' | 'offsetXPx' | 'offsetYPx' | 'payload'>,
  rId: string,
  docPrId: number,
): string {
  // Math.max/min propagate NaN, so sanitize first: a non-finite measurement
  // would otherwise serialize as cx="NaN" corrupt OOXML.
  const safePx = (v: number, fallback: number): number =>
    Number.isFinite(v) ? Math.round(v * EMU_PER_PX) : fallback
  const cx = Math.max(1, safePx(ink.widthPx, 1))
  const cy = Math.max(1, safePx(ink.heightPx, 1))
  const x = safePx(ink.offsetXPx, 0)
  const y = safePx(ink.offsetYPx, 0)
  const name = `${INK_NAME_PREFIX} ${docPrId}`
  const descr = ink.payload ? ` descr="${escapeXmlAttr(ink.payload)}"` : ''
  return (
    '<w:r><w:drawing>' +
    `<wp:anchor xmlns:wp="${WP_NS}" distT="0" distB="0" distL="0" distR="0" simplePos="0"` +
    ` relativeHeight="${251658240 + docPrId}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    '<wp:wrapNone/>' +
    `<wp:docPr id="${docPrId}" name="${name}"${descr}/>` +
    '<wp:cNvGraphicFramePr/>' +
    `<a:graphic xmlns:a="${A_NS}">` +
    `<a:graphicData uri="${PIC_NS}">` +
    `<pic:pic xmlns:pic="${PIC_NS}">` +
    `<pic:nvPicPr><pic:cNvPr id="${docPrId}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="${R_NS}" r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
  )
}

const ANCHOR_RUN_RE = /<w:r><w:drawing><wp:anchor[\s\S]*?<\/wp:anchor><\/w:drawing><\/w:r>/g

const isInkRun = (run: string) => run.includes(`name="${INK_NAME_PREFIX}`)

/** Remove every aidocs-ink run from a body-element XML slice. */
export function stripInkRuns(xml: string): string {
  if (!xml.includes(INK_NAME_PREFIX)) return xml
  return xml.replace(ANCHOR_RUN_RE, (run) => (isInkRun(run) ? '' : run))
}

/** One raw ink run found in a paragraph's XML. */
export interface InkRunMatch {
  /** the exact <w:r>...</w:r> slice, for stripping */
  xml: string
  offsetXPx: number
  offsetYPx: number
  widthPx: number
  heightPx: number
  payload: string | null
  /** relationship id of the embedded PNG */
  embedRId: string | null
}

/** Find aidocs-ink anchored runs inside one paragraph's XML. */
export function findInkRuns(paragraphXml: string): InkRunMatch[] {
  if (!paragraphXml.includes(INK_NAME_PREFIX)) return []
  const out: InkRunMatch[] = []
  for (const run of paragraphXml.match(ANCHOR_RUN_RE) ?? []) {
    if (!isInkRun(run)) continue
    const emu = (value: string | undefined) => parseInt(value ?? '0', 10) || 0
    const descr = /<wp:docPr [^>]*descr="([^"]*)"/.exec(run)?.[1]
    // the extent attributes come in any order and need not start the tag, so read
    // cx and cy off the matched element instead of pinning one to the other
    const extent = /<wp:extent[^>]*\/?>/.exec(run)?.[0] ?? ''
    out.push({
      xml: run,
      offsetXPx: emu(/<wp:positionH[^>]*><wp:posOffset>(-?\d+)/.exec(run)?.[1]) / EMU_PER_PX,
      offsetYPx: emu(/<wp:positionV[^>]*><wp:posOffset>(-?\d+)/.exec(run)?.[1]) / EMU_PER_PX,
      widthPx: emu(/\bcx="(\d+)"/.exec(extent)?.[1]) / EMU_PER_PX,
      heightPx: emu(/\bcy="(\d+)"/.exec(extent)?.[1]) / EMU_PER_PX,
      payload: descr ? decodeXmlEntities(descr) : null,
      embedRId: /r:embed="([^"]+)"/.exec(run)?.[1] ?? null,
    })
  }
  return out
}

/**
 * Insert ink runs at the end of a paragraph fragment (after all content
 * runs). Returns null when the fragment's root is not a w:p — floating ink
 * can only anchor to paragraphs.
 */
export function injectInkRunsIntoParagraph(xml: string, runsXml: string): string | null {
  if (!/^<w:p[\s/>]/.test(xml)) return null
  if (xml.endsWith('/>')) return `${xml.slice(0, -2)}>${runsXml}</w:p>`
  const close = xml.lastIndexOf('</w:p>')
  if (close === -1) return null
  return xml.slice(0, close) + runsXml + xml.slice(close)
}

function decodeXmlEntities(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}
