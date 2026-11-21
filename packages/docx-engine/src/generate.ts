import { sdtCheckboxGlyphs, sdtCheckboxIsChecked, syncSdtCheckbox } from './checkbox-control'
import type {
  CellMargins,
  CharIndents,
  FieldTextRun,
  GeneratedBlock,
  ImageWrap,
  ParaFormat,
  ParaFrame,
  PropChangeInfo,
  RevisionInfo,
  Run,
  TextGlow,
  TextOutline,
  TextReflection,
  TextShadow,
  TextboxVert,
  TabStop,
  TableCell,
  TableModel,
} from './types'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'
import { MERGE_RULE_INSTR_RE, NO_FIELD_RESULT } from './parse-props'
import { resolvePropChangeXml } from './revision-xml'

export interface GenerateContext {
  /** heading level -> styleId existing in the original styles.xml */
  headingStyleIds: Map<number, string>
  /** styleId for list paragraphs, if present in the original doc */
  listParagraphStyleId?: string
  /** allocate a new relationship id for a hyperlink target; returns rId */
  allocateHyperlinkRel: (href: string) => string
  /**
   * Mint the w:bookmarkStart/@w:id for a bookmark name. w:id must be unique in
   * the part, so a caller writing into an existing document supplies an
   * allocator seeded above that part's highest id. Without one (a fragment
   * rendered outside any document) ids come from a process counter, which
   * keeps them unique among themselves but is not coordinated with a part.
   */
  allocateBookmarkId?: (name: string) => number
  /** Mint w:id for a tracked change that has none; document-wide so two edited paragraphs never share ids */
  allocateRevisionId?: () => string
}

const EMU_PER_PX = 9525
const EMU_PER_PT = 12700

export interface ImagePatch {
  /** new display size in CSS px; rewrites wp:extent and pic a:ext */
  widthPx?: number
  heightPx?: number
  /** paragraph alignment; null/'left' removes w:jc, undefined keeps as-is */
  align?: 'left' | 'center' | 'right' | null
  /**
   * New horizontal/vertical posOffset (in EMU) for floating images. Rewrites
   * the <wp:posOffset> inside wp:positionH / wp:positionV.
   * Ignored when the anchor uses <wp:align> instead of <wp:posOffset>.
   */
  posOffsetX?: number
  posOffsetY?: number
  /** rotation in degrees clockwise; 0 removes the rot attribute; undefined keeps */
  rotDeg?: number
  /** mirror flips; false removes the attribute; undefined keeps */
  flipH?: boolean
  flipV?: boolean
  /** alternative text (wp:docPr descr); '' removes it; undefined keeps */
  altText?: string
}

/** Sets the alternative text of the first drawing in `xml` (wp:docPr descr); '' removes it. */
export function setDrawingAltText(xml: string, alt: string): string {
  // the first docPr belongs to the drawing's own anchor or inline
  return xml.replace(/<wp:docPr\b([^>]*?)(\/?)>/, (_whole, attrs: string, close: string) => {
    const rest = attrs.replace(/\s*\bdescr\s*=\s*(?:"[^"]*"|'[^']*')/, '')
    return `<wp:docPr${rest}${alt ? ` descr="${escapeXmlAttr(alt)}"` : ''}${close}>`
  })
}

/**
 * Rewrite the display size (wp:extent + a:xfrm a:ext) and/or paragraph
 * alignment (w:jc) of an image paragraph, leaving everything else untouched.
 */
export function patchImageParagraphXml(xml: string, patch: ImagePatch): string {
  let out = xml
  if (patch.altText !== undefined) out = setDrawingAltText(out, patch.altText)
  if (patch.rotDeg !== undefined || patch.flipH !== undefined || patch.flipV !== undefined) {
    // rotation/flip live on the pic's own xfrm, not an anchored textbox sibling's
    out = out.replace(/(<pic:spPr[^>]*>[\s\S]*?)<a:xfrm([^>]*)>/, (_whole, prefix, attrs) => {
      let a = attrs as string
      const setAttr = (name: string, value: string | null) => {
        a = a.replace(new RegExp(`\\s*\\b${name}\\s*=\\s*(?:"[^"]*"|'[^']*')`), '')
        if (value != null) a += ` ${name}="${value}"`
      }
      if (patch.rotDeg !== undefined && Number.isFinite(patch.rotDeg)) {
        const norm = ((Math.round(patch.rotDeg) % 360) + 360) % 360
        setAttr('rot', norm ? String(norm * 60000) : null)
      }
      if (patch.flipH !== undefined) setAttr('flipH', patch.flipH ? '1' : null)
      if (patch.flipV !== undefined) setAttr('flipV', patch.flipV ? '1' : null)
      return `${prefix}<a:xfrm${a}>`
    })
  }
  if (patch.widthPx && patch.heightPx) {
    // Non-finite dimensions must not land in the XML verbatim (cx="Infinity"
    // is schema-invalid); skip the resize and keep the original extents.
    const toEmu = (px: number): number | null =>
      Number.isFinite(px) ? Math.max(1, Math.round(px * EMU_PER_PX)) : null
    const cx = toEmu(patch.widthPx)
    const cy = toEmu(patch.heightPx)
    if (cx !== null && cy !== null) {
      const resize = (tag: string) =>
        tag.replace(/cx="\d+"/, `cx="${cx}"`).replace(/cy="\d+"/, `cy="${cy}"`)
      out = out.replace(/<wp:extent[^>]*\/?>/, resize)
      out = out.replace(/<a:ext[^>]*\/>/, resize)
    }
  }
  // Word lays the drawing out against the unrotated wp:extent plus
  // wp:effectExtent: a 90°/270° turn of a non-square picture needs the extra
  // bounding-box space recorded there or Word crops / misplaces it. Recompute
  // only when rotation changes, or the size changes while a rotation is in
  // effect — flips leave the bounding box alone, and untouched drawings keep
  // their original effectExtent (possibly Word-authored shadow/glow padding).
  {
    const rotM = /<pic:spPr[^>]*>[\s\S]*?<a:xfrm[^>]*?\brot="(-?\d+)"/.exec(out)
    const extM = /<wp:extent[^>]*?\bcx="(\d+)"[^>]*?\bcy="(\d+)"/.exec(out)
    const touchRot = patch.rotDeg !== undefined
    const touchSize = !!(patch.widthPx && patch.heightPx)
    if (extM && (touchRot || (touchSize && rotM))) {
      const rad = (((rotM ? Number(rotM[1]) : 0) / 60000) * Math.PI) / 180
      const cx = Number(extM[1])
      const cy = Number(extM[2])
      const bw = Math.abs(cx * Math.cos(rad)) + Math.abs(cy * Math.sin(rad))
      const bh = Math.abs(cx * Math.sin(rad)) + Math.abs(cy * Math.cos(rad))
      const dx = Math.max(0, Math.round((bw - cx) / 2))
      const dy = Math.max(0, Math.round((bh - cy) / 2))
      const ee = `<wp:effectExtent l="${dx}" t="${dy}" r="${dx}" b="${dy}"/>`
      if (/<wp:effectExtent\b[^>]*\/>/.test(out))
        out = out.replace(/<wp:effectExtent\b[^>]*\/>/, ee)
      else out = out.replace(/(<wp:extent\b[^>]*\/?>)/, `$1${ee}`)
    }
  }
  if (patch.align !== undefined) {
    out = out.replace(/<w:jc w:val="[^"]*"\/>/, '')
    out = out.replace(/<w:pPr\s*\/>/, '<w:pPr></w:pPr>')
    const align = patch.align === 'left' ? null : patch.align
    if (align) {
      const jc = `<w:jc w:val="${align}"/>`
      const pPr = /(<w:pPr[^>]*>)([\s\S]*?)<\/w:pPr>/.exec(out)
      if (pPr) {
        // w:jc must precede the paragraph-mark w:rPr inside pPr
        const [whole, open, inner] = pPr
        const rPrIdx = inner.indexOf('<w:rPr>')
        const patched =
          rPrIdx === -1 ? inner + jc : inner.slice(0, rPrIdx) + jc + inner.slice(rPrIdx)
        out =
          out.slice(0, pPr.index) +
          open +
          patched +
          '</w:pPr>' +
          out.slice(pPr.index + whole.length)
      } else {
        out = out.replace(/(<w:p(?: [^>]*)?>)/, `$1<w:pPr>${jc}</w:pPr>`)
      }
    }
  }
  // Rewrite the position inside positionH / positionV (surgical); an alignment keyword
  // gives way to the offset. Non-finite offsets would land verbatim (posOffset>NaN<) —
  // skip and keep the original.
  const setOffset = (axis: 'H' | 'V', value: number | undefined): void => {
    if (value === undefined || !Number.isFinite(value)) return
    out = out.replace(
      new RegExp(
        `(<wp:position${axis}[^>]*>[\\s\\S]*?)<wp:(posOffset|align)>[^<]*<\\/wp:\\2>([\\s\\S]*?<\\/wp:position${axis}>)`,
      ),
      `$1<wp:posOffset>${Math.round(value)}</wp:posOffset>$3`,
    )
  }
  setOffset('H', patch.posOffsetX)
  setOffset('V', patch.posOffsetY)
  return out
}

const WRAP_ELEMENT_RE =
  /<wp:wrapNone\s*\/>|<wp:wrapSquare[^>]*\/>|<wp:wrapTight[^>]*\/>|<wp:wrapThrough[^>]*\/>|<wp:wrapTopAndBottom\s*\/>|<wp:wrapSquare[\s\S]*?<\/wp:wrapSquare>|<wp:wrapTight[\s\S]*?<\/wp:wrapTight>|<wp:wrapThrough[\s\S]*?<\/wp:wrapThrough>|<wp:wrapTopAndBottom[\s\S]*?<\/wp:wrapTopAndBottom>/g

/**
 * Re-encode ONLY the stacking rank of an existing wp:anchor as Word's
 * base + rank relativeHeight. Everything else — position basis (relativeFrom),
 * wrap element, distances — keeps its original bytes, so a picture never
 * shifts from a pure reorder (or from the save-time harmonization of wild
 * producer values). No-op on inline images (no wp:anchor).
 */
export function applyImageZOrder(xml: string, zOrder?: number): string {
  return xml.replace(/(<wp:anchor[^>]*?relativeHeight=")\d+(")/, `$1${251658240 + (zOrder ?? 0)}$2`)
}

/**
 * Switch an image paragraph between inline (in line with text, wrap = null) and floating
 * (wp:anchor with the given wrap mode). Position/wrap elements are rebuilt;
 * extent, docPr and the pic graphic stay untouched.
 *
 * When `posOffset` is provided, the anchor's positionH/V use numeric
 * `<wp:posOffset>` instead of `<wp:align>`, enabling free-position drag.
 */
export function applyImageWrap(
  xml: string,
  wrap: ImageWrap | null,
  posOffset?: {
    x: number
    y: number
    relativeTo?: 'page' | 'margin'
    /** what each axis measures from, when the two differ; wins over `relativeTo` */
    relH?: 'page' | 'margin' | 'column'
    relV?: 'page' | 'margin' | 'paragraph'
  },
  marginAlign?: { h: 'left' | 'center' | 'right'; v: 'top' | 'center' | 'bottom' },
  zOrder?: number,
): string {
  const hasAnchor = /<wp:anchor[\s>]/.test(xml)
  // Original wrapTight/wrapThrough bytes (including wp:wrapPolygon) — reused as a whole
  // when the target is still the same wrap kind
  const existingWrap = xml.match(WRAP_ELEMENT_RE)?.[0]
  let out = xml
    .replace(/<wp:simplePos[^>]*\/>/, '')
    .replace(/<wp:positionH[\s\S]*?<\/wp:positionH>/, '')
    .replace(/<wp:positionV[\s\S]*?<\/wp:positionV>/, '')
    .replace(WRAP_ELEMENT_RE, '')

  if (!wrap) {
    if (!hasAnchor) return xml
    out = out.replace(/<wp:anchor[^>]*>/, '<wp:inline distT="0" distB="0" distL="0" distR="0">')
    return out.replace(/<\/wp:anchor>/, '</wp:inline>')
  }

  const behind = wrap === 'behind' ? '1' : '0'
  const isSide =
    wrap === 'square-left' ||
    wrap === 'square-right' ||
    wrap === 'tight-left' ||
    wrap === 'tight-right' ||
    wrap === 'through-left' ||
    wrap === 'through-right'
  let position: string
  if (posOffset !== undefined) {
    const relH = posOffset.relH ?? posOffset.relativeTo ?? 'column'
    const relV = posOffset.relV ?? posOffset.relativeTo ?? 'paragraph'
    position =
      '<wp:simplePos x="0" y="0"/>' +
      `<wp:positionH relativeFrom="${relH}"><wp:posOffset>${Math.round(posOffset.x)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="${relV}"><wp:posOffset>${Math.round(posOffset.y)}</wp:posOffset></wp:positionV>`
  } else if (marginAlign !== undefined) {
    position =
      '<wp:simplePos x="0" y="0"/>' +
      `<wp:positionH relativeFrom="margin"><wp:align>${marginAlign.h}</wp:align></wp:positionH>` +
      `<wp:positionV relativeFrom="margin"><wp:align>${marginAlign.v}</wp:align></wp:positionV>`
  } else {
    const hAlign = wrap.endsWith('-right') ? 'right' : wrap === 'topBottom' ? 'center' : 'left'
    position =
      '<wp:simplePos x="0" y="0"/>' +
      `<wp:positionH relativeFrom="column"><wp:align>${hAlign}</wp:align></wp:positionH>` +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
  }
  const keepKind = wrap.startsWith('tight-')
    ? 'wp:wrapTight'
    : wrap.startsWith('through-')
      ? 'wp:wrapThrough'
      : null
  const wrapElement =
    keepKind && existingWrap?.startsWith(`<${keepKind}`)
      ? existingWrap
      : isSide
        ? '<wp:wrapSquare wrapText="bothSides"/>'
        : wrap === 'topBottom'
          ? '<wp:wrapTopAndBottom/>'
          : '<wp:wrapNone/>'
  const anchorOpen =
    `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"` +
    ` relativeHeight="${251658240 + (zOrder ?? 0)}" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1">`
  if (hasAnchor) {
    out = out.replace(/<wp:anchor[^>]*>/, anchorOpen)
  } else {
    out = out.replace(/<wp:inline[^>]*>/, anchorOpen).replace(/<\/wp:inline>/, '</wp:anchor>')
  }
  out = out.replace(/(<wp:anchor[^>]*>)/, `$1${position}`)
  // wrap element sits between extent/effectExtent and docPr in CT_Anchor order
  if (/<wp:docPr/.test(out)) return out.replace(/<wp:docPr/, `${wrapElement}<wp:docPr`)
  return out.replace(/<a:graphic[\s>]/, (m) => `${wrapElement}${m}`)
}

/** The box drawings (wps shapes/textboxes) in a paragraph, in document order — same walk patchShapeStyles uses. */
function boxDrawingSegments(paragraphXml: string): Array<{ start: number; end: number }> {
  return xmlSegments(paragraphXml, 'w:drawing', 0, paragraphXml.length).filter((seg) =>
    isBoxDrawing(paragraphXml.slice(seg.start, seg.end)),
  )
}

/** How a caller addresses one shape drawing: by its cNvPr id when known, else by box ordinal. */
export interface ShapeDrawingLocation {
  /** wps:cNvPr (or pic:cNvPr for a photo box) id of the owning shape (parsed boxes carry it) */
  shapeId?: string
  /** fallback ordinal among box drawings (generated shapes have no id) */
  boxIndex: number
}

function shapeDrawingSegment(
  paragraphXml: string,
  location: ShapeDrawingLocation,
): { start: number; end: number } | null {
  if (location.shapeId) {
    // an id was given: never fall back to a different drawing
    const found = xmlSegments(paragraphXml, 'w:drawing', 0, paragraphXml.length).find(
      (seg) =>
        /<(?:wps|pic):cNvPr\b[^>]*\bid="([^"]+)"/.exec(
          paragraphXml.slice(seg.start, seg.end),
        )?.[1] === location.shapeId,
    )
    return found ?? null
  }
  return boxDrawingSegments(paragraphXml)[location.boxIndex] ?? null
}

/** Re-encode the rank of the drawing's own wp:anchor tag, never a nested drawing's. */
function setAnchorRank(drawingXml: string, zOrder: number): string {
  return drawingXml.replace(/<wp:anchor[^>]*>/, (tag) =>
    tag.replace(/relativeHeight="\d+"/, `relativeHeight="${251658240 + zOrder}"`),
  )
}

/**
 * Re-encode ONLY the stacking rank of one shape drawing as Word's base + rank
 * relativeHeight. Everything else keeps its bytes, like applyImageZOrder.
 */
export function applyShapeZOrderAt(
  paragraphXml: string,
  location: ShapeDrawingLocation,
  zOrder: number,
): string {
  const seg = shapeDrawingSegment(paragraphXml, location)
  if (!seg) return paragraphXml
  const next = setAnchorRank(paragraphXml.slice(seg.start, seg.end), zOrder)
  return paragraphXml.slice(0, seg.start) + next + paragraphXml.slice(seg.end)
}

/**
 * Switch one shape drawing's wrap mode in place, preserving its positionH/V
 * bytes: only the anchor attributes, the wrap element and (when given)
 * relativeHeight change. `wrap === null` converts the anchor to inline.
 */
export function applyShapeWrapAt(
  paragraphXml: string,
  location: ShapeDrawingLocation,
  wrap: ImageWrap | null,
  zOrder?: number,
): string {
  const seg = shapeDrawingSegment(paragraphXml, location)
  if (!seg) return paragraphXml
  const drawing = paragraphXml.slice(seg.start, seg.end)
  const hasAnchor = /<wp:anchor[\s>]/.test(drawing)
  let next: string
  if (wrap === null) {
    if (!hasAnchor) return paragraphXml
    next = drawing
      .replace(/<wp:simplePos[^>]*\/>/, '')
      .replace(/<wp:positionH[\s\S]*?<\/wp:positionH>/, '')
      .replace(/<wp:positionV[\s\S]*?<\/wp:positionV>/, '')
      .replace(WRAP_ELEMENT_RE, '')
      .replace(/<wp:anchor[^>]*>/, '<wp:inline distT="0" distB="0" distL="0" distR="0">')
      .replace(/<\/wp:anchor>/, '</wp:inline>')
  } else if (!hasAnchor) {
    next = applyImageWrap(drawing, wrap, undefined, undefined, zOrder)
  } else {
    next = swapAnchorWrap(drawing, wrap, zOrder)
  }
  return paragraphXml.slice(0, seg.start) + next + paragraphXml.slice(seg.end)
}

/**
 * Switch the wrap mode of an existing wp:anchor drawing in place: only the
 * wrap element, behindDoc and (when given) relativeHeight change, so the
 * positionH/V bytes keep the object where it was.
 */
export function swapAnchorWrap(drawingXml: string, wrap: ImageWrap, zOrder?: number): string {
  const behind = wrap === 'behind' ? '1' : '0'
  const wrapElement =
    wrap === 'front' || wrap === 'behind'
      ? '<wp:wrapNone/>'
      : wrap === 'topBottom'
        ? '<wp:wrapTopAndBottom/>'
        : '<wp:wrapSquare wrapText="bothSides"/>'
  let next = drawingXml
    .replace(WRAP_ELEMENT_RE, '')
    .replace(/<wp:anchor[^>]*>/, (tag) =>
      tag.includes('behindDoc=')
        ? tag.replace(/behindDoc="[^"]*"/, `behindDoc="${behind}"`)
        : tag.replace(/<wp:anchor/, `<wp:anchor behindDoc="${behind}"`),
    )
  if (zOrder !== undefined) next = setAnchorRank(next, zOrder)
  return /<wp:docPr/.test(next)
    ? next.replace(/<wp:docPr/, `${wrapElement}<wp:docPr`)
    : next.replace(/<a:graphic[\s>]/, (m) => `${wrapElement}${m}`)
}

// ---- protected field / formula token patching ----

interface XmlTextNode {
  start: number
  end: number
  open: string
  close: string
  text: string
}

function decodeXmlText(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_m, dec: string) => {
      const cp = Number(dec)
      if (!Number.isFinite(cp) || cp < 0 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff))
        return _m
      return String.fromCodePoint(cp)
    })
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, hex: string) => {
      const cp = parseInt(hex, 16)
      if (!Number.isFinite(cp) || cp < 0 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff))
        return _m
      return String.fromCodePoint(cp)
    })
    .replace(/&amp;/g, '&')
}

function textNodes(xml: string, tag: 'w:t' | 'm:t'): XmlTextNode[] {
  const nodes: XmlTextNode[] = []
  // Paired text nodes: <w:t>text</w:t> (may carry xml:space)
  const re = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`, 'g')
  let match: RegExpExecArray | null
  while ((match = re.exec(xml)) !== null) {
    const openEnd = match[0].indexOf('>') + 1
    const closeStart = match[0].lastIndexOf(`</${tag}>`)
    nodes.push({
      start: match.index,
      end: match.index + match[0].length,
      open: match[0].slice(0, openEnd),
      close: match[0].slice(closeStart),
      text: decodeXmlText(match[1]),
    })
  }
  // Self-closing empty <w:t/> (Word emits these for empty runs). Only for w:t:
  // mathTokensOf tokenizes paired <m:t> only, so counting <m:t/> here would
  // desync patchMathTokens' length check against the tokens it was given.
  if (tag === 'w:t') {
    const selfRe = /<w:t(?:\s[^>]*)?\/>/g
    while ((match = selfRe.exec(xml)) !== null) {
      nodes.push({
        start: match.index,
        end: match.index + match[0].length,
        open: match[0],
        close: '',
        text: '',
      })
    }
    nodes.sort((a, b) => a.start - b.start)
  }
  return nodes
}

function distributeText(value: string, nodes: XmlTextNode[]): string[] {
  if (nodes.length === 0) return []
  const parts: string[] = []
  let cursor = 0
  for (let i = 0; i < nodes.length; i++) {
    if (i === nodes.length - 1) {
      parts.push(value.slice(cursor))
      break
    }
    const length = nodes[i].text.length
    parts.push(value.slice(cursor, cursor + length))
    cursor += length
  }
  return parts
}

function replaceTextNodes(
  xml: string,
  replacements: ReadonlyArray<{ node: XmlTextNode; text: string }>,
): string {
  let out = xml
  for (const { node, text } of [...replacements].sort((a, b) => b.node.start - a.node.start)) {
    // wtText drops edge whitespace without xml:space="preserve": pin it when
    // the replacement introduces leading/trailing whitespace. Word trims more
    // than just tab/space (NBSP, zero-width, line breaks, CJK spaces).
    const hasEdgeWs =
      text.length > 0 &&
      (/^\s|\s$/.test(text) || text.startsWith('\u200B') || text.endsWith('\u200B'))
    const open =
      /xml:space\s*=/.test(node.open) || !hasEdgeWs
        ? node.open
        : node.open.replace(/<w:t(?=[\s>/])/, '<w:t xml:space="preserve"')
    // Self-closing empty run: expand to paired form so the replacement lands
    const close = node.close || `</w:t>`
    const openPaired = node.close ? open : open.replace(/\/>$/, '>')
    const replacement = openPaired + escapeXmlText(text) + close
    out = out.slice(0, node.start) + replacement + out.slice(node.end)
  }
  return out
}

export interface FieldTextPatch {
  left?: string
  right?: string
}

/**
 * Patch only cached visible w:t text in a field paragraph. Field instructions,
 * hyperlinks, tabs, run styling, and field boundaries remain byte-identical.
 */
export function patchFieldParagraphXml(xml: string, patch: FieldTextPatch): string {
  const hyphens =
    /<w:(?:noBreakHyphen|softHyphen)\s*\/>/.test(xml) || /[\u2011\u00ad]/.test(patch.left ?? '')
  if (!hyphens) return patchFieldTextNodes(xml, patch)
  // fieldDisplayOf reads the hyphen elements as U+2011 / U+00AD, so the edited
  // text carries them as characters: lift the elements into text nodes for the
  // distribution and fold every hyphen character back into its element after
  // (a tracked deletion's hyphens lift as w:delText so they stay out of the live text)
  const lift = (seg: string, tag: string) =>
    seg
      .replace(/<w:noBreakHyphen\s*\/>/g, `<${tag}>\u2011</${tag}>`)
      .replace(/<w:softHyphen\s*\/>/g, `<${tag}>\u00ad</${tag}>`)
  const lifted = lift(
    xml.replace(/<w:del(?:\s[^>]*)?(?<!\/)>[\s\S]*?<\/w:del>/g, (del) => lift(del, 'w:delText')),
    'w:t',
  )
  return patchFieldTextNodes(lifted, patch).replace(
    /<(w:t|w:delText)(\s[^>]*)?>([^<]*[\u2011\u00ad][^<]*)<\/\1>/g,
    (_m, tag: string, attrs: string | undefined, text: string) =>
      text
        .split(/([\u2011\u00ad])/)
        .filter(Boolean)
        .map((seg) =>
          seg === '\u2011'
            ? '<w:noBreakHyphen/>'
            : seg === '\u00ad'
              ? '<w:softHyphen/>'
              : `<${tag}${attrs ?? ''}>${seg}</${tag}>`,
        )
        .join(''),
  )
}

function patchFieldTextNodes(xml: string, patch: FieldTextPatch): string {
  const nodes = textNodes(xml, 'w:t')
  if (nodes.length === 0) return xml
  // run-level tab only (attribute-less CT_Empty) — `<w:tab w:val=…/>` inside
  // w:tabs is a tab-stop definition and must not split the left/right texts.
  // Mirrors fieldDisplayOf: the page number follows the LAST tab, and a short
  // space-free first segment before ≥2 tabs is the outline-number cell (not
  // part of the editable title).
  const tabStarts: number[] = []
  // attribute-less run tab only: spaced (<w:tab />) and paired
  // (<w:tab></w:tab>) variants from non-Word producers count as well.
  const tabRe = /<w:tab\s*\/>|<w:tab>\s*<\/w:tab>/g
  let tabMatch: RegExpExecArray | null
  while ((tabMatch = tabRe.exec(xml)) !== null) tabStarts.push(tabMatch.index)
  const lastTab = tabStarts.length > 0 ? tabStarts[tabStarts.length - 1] : -1
  if (patch.right !== undefined && lastTab === -1) return xml
  let leftFrom = -1
  if (tabStarts.length >= 2) {
    const firstSeg = nodes
      .filter((node) => node.start < tabStarts[0])
      .map((node) => node.text)
      .join('')
      .trim()
    if (/^\S{1,15}$/.test(firstSeg)) leftFrom = tabStarts[0]
  }
  const leftNodes =
    lastTab === -1 ? nodes : nodes.filter((node) => node.start > leftFrom && node.start < lastTab)
  const rightNodes = lastTab === -1 ? [] : nodes.filter((node) => node.start > lastTab)
  if (patch.left !== undefined && leftNodes.length === 0) return xml
  if (patch.right !== undefined && rightNodes.length === 0) return xml

  const replacements: Array<{ node: XmlTextNode; text: string }> = []
  if (patch.left !== undefined) {
    distributeText(patch.left, leftNodes).forEach((text, i) => {
      replacements.push({ node: leftNodes[i], text })
    })
  }
  if (patch.right !== undefined) {
    distributeText(patch.right, rightNodes).forEach((text, i) => {
      replacements.push({ node: rightNodes[i], text })
    })
  }
  return replaceTextNodes(xml, replacements)
}

/**
 * Patch OMML leaf token text without rebuilding its structural math nodes.
 * Token count must stay unchanged so fractions, scripts, matrices, and styling
 * remain exactly as authored by Word.
 */
export function patchMathTokens(xml: string, tokens: readonly string[]): string {
  const nodes = textNodes(xml, 'm:t')
  if (nodes.length === 0 || nodes.length !== tokens.length) return xml
  return replaceTextNodes(
    xml,
    nodes.map((node, i) => ({ node, text: tokens[i] })),
  )
}

// ---- table cell text patching ----

/** top-level segments of `tag` between [from, to), depth-aware for nesting */
export function xmlSegments(
  xml: string,
  tag: string,
  from: number,
  to: number,
): Array<{ start: number; end: number }> {
  const openPrefix = '<' + tag
  const closeTag = '</' + tag + '>'
  const segs: Array<{ start: number; end: number }> = []
  let depth = 0
  let segStart = -1
  let i = from
  while (i < to) {
    const o = xml.indexOf(openPrefix, i)
    const c = xml.indexOf(closeTag, i)
    const hasOpen = o !== -1 && o < to
    const hasClose = c !== -1 && c < to
    // a trailing self-closing element (Word's empty <w:p/>) has no close tag after it
    if (!hasOpen && !hasClose) break
    if (hasOpen && (!hasClose || o < c)) {
      const after = xml.charAt(o + openPrefix.length)
      if (after !== '>' && after !== ' ' && after !== '/') {
        i = o + openPrefix.length // prefix of a longer tag (w:tr vs w:trPr)
        continue
      }
      const gt = xml.indexOf('>', o)
      if (gt !== -1 && xml.charAt(gt - 1) === '/') {
        if (depth === 0) segs.push({ start: o, end: gt + 1 }) // self-closing
        i = gt + 1
        continue
      }
      if (depth === 0) segStart = o
      depth++
      i = o + openPrefix.length
    } else if (depth === 0) {
      break
    } else {
      depth--
      if (depth === 0) segs.push({ start: segStart, end: c + closeTag.length })
      if (depth < 0) break
      i = c + closeTag.length
    }
  }
  return segs
}

/**
 * Rebuild one cell's paragraphs, keeping tcPr. String entries keep the legacy
 * behavior (first paragraph's pPr / first run's rPr for every paragraph); rich
 * entries regenerate full runs against their own original paragraph's pPr, and
 * null entries keep the original paragraph bytes untouched.
 * Returns null when the cell is too complex to patch safely (nested table).
 */
function patchCellXml(tcXml: string, paras: readonly CellParaPatch[]): string | null {
  if (tcXml.indexOf('<w:tbl', 1) !== -1) return null
  const openTag = /^<w:tc(?: [^>]*)?>/.exec(tcXml)?.[0]
  if (!openTag) return null
  const tcPr = /<w:tcPr[\s\S]*?<\/w:tcPr>|<w:tcPr[^>]*\/>/.exec(tcXml)?.[0] ?? ''
  const originals = xmlSegments(tcXml, 'w:p', openTag.length, tcXml.length).map((seg) =>
    tcXml.slice(seg.start, seg.end),
  )
  const pPrOf = (pXml: string) => /<w:pPr[\s\S]*?<\/w:pPr>|<w:pPr[^>]*\/>/.exec(pXml)?.[0] ?? ''
  const firstPPr = pPrOf(originals[0] ?? '')
  const firstRun = /<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/.exec(originals[0] ?? '')?.[0] ?? ''
  const rPr = /<w:rPr[\s\S]*?<\/w:rPr>/.exec(firstRun)?.[0] ?? ''
  // picture runs are not part of the text model; carry them over verbatim so a
  // text edit in a cell with an inline image doesn't drop the image
  const drawingRuns = tcXml.includes('<w:drawing')
    ? xmlSegments(tcXml, 'w:r', 0, tcXml.length)
        .map((seg) => tcXml.slice(seg.start, seg.end))
        .filter((r) => r.includes('<w:drawing'))
        .join('')
    : ''
  // The run model re-emits every drawing it resolved media for (run.image.xml),
  // including Word's mc:Choice drawing whose VML w:pict fallback shares the same
  // media relationship. Only drawings the model cannot see (shapes, textboxes,
  // charts, blind pict/object) must ride over. Carried copies keep only the
  // drawing payload children — run-level text belongs to the rich runs, while
  // text nested inside the drawing (textbox content) stays intact.
  const mediaRIds = (xml: string) =>
    [...xml.matchAll(/<(?:a:blip|v:imagedata)\b[^>]*\br:(?:embed|link|id)="([^"]+)"/g)].map(
      (m) => m[1],
    )
  const DRAWING_TAGS = new Set(['mc:AlternateContent', 'w:drawing', 'w:pict', 'w:object'])
  const carriedRunsOf = (pXml: string, emittedMedia: Set<string>) => {
    // a payload is a twin of an emitted image when all its media is already
    // re-emitted and it carries no textbox text of its own
    const isTwin = (payload: string) => {
      const ids = mediaRIds(payload)
      return (
        ids.length > 0 &&
        ids.every((id) => emittedMedia.has(id)) &&
        !/<w:txbxContent[\s>]|<v:textbox[\s>]/.test(payload)
      )
    }
    return xmlSegments(pXml, 'w:r', 0, pXml.length)
      .map((seg) => pXml.slice(seg.start, seg.end))
      .filter((r) => /<wp:anchor[\s>]|<w:pict[\s>]|<w:object[\s>]/.test(r))
      .map((r) => {
        const open = /^<w:r(?:\s[^>]*)?>/.exec(r)?.[0]
        if (!open) return ''
        const kept = splitXmlChildren(r.slice(open.length, r.length - '</w:r>'.length)).filter(
          (c) => DRAWING_TAGS.has(c.name) && !isTwin(c.xml),
        )
        return kept.length > 0 ? `<w:r>${kept.map((c) => c.xml).join('')}</w:r>` : ''
      })
      .join('')
  }
  const aligned = paras.length === originals.length
  const body = paras
    .map((entry, i) => {
      if (entry === null) return originals[i] ?? ''
      if (typeof entry === 'string') {
        const keep = i === 0 ? drawingRuns : ''
        return entry === ''
          ? `<w:p>${firstPPr}${keep}</w:p>`
          : `<w:p>${firstPPr}<w:r>${rPr}${cellRunContentXml(entry)}</w:r>${keep}</w:p>`
      }
      const template = originals[Math.min(i, originals.length - 1)] ?? ''
      const pPr = entry.format
        ? mergeCellParaFormat(pPrOf(template), entry.format, entry.original)
        : pPrOf(template)
      const emitted = entry.xml ?? runsXml(entry.runs, null)
      const emittedMedia = new Set(mediaRIds(emitted))
      const keep = aligned
        ? carriedRunsOf(template, emittedMedia)
        : i === 0
          ? originals.map((p) => carriedRunsOf(p, emittedMedia)).join('')
          : ''
      return `<w:p>${pPr}${emitted}${keep}</w:p>`
    })
    .join('')
  return openTag + tcPr + body + '</w:tc>'
}

/** cell paragraph texts carry tabs/breaks as control chars (extractRuns encoding);
 *  they must go back as real elements, not literal characters inside w:t */
function cellRunContentXml(text: string): string {
  const CONTROL: Record<string, string> = {
    '\t': '<w:tab/>',
    '\n': '<w:br/>',
    '\f': '<w:br w:type="page"/>',
    '\v': '<w:br w:type="column"/>',
    '\u001e': '<w:br w:type="textWrapping" w:clear="all"/>',
    '\u2011': '<w:noBreakHyphen/>',
    '\u00ad': '<w:softHyphen/>',
  }
  // eslint-disable-next-line no-control-regex
  const parts = text.split(/([\t\n\f\v\u001e\u2011\u00ad])/)
  return parts
    .map((seg) =>
      seg === '' ? '' : (CONTROL[seg] ?? `<w:t xml:space="preserve">${escapeXmlText(seg)}</w:t>`),
    )
    .join('')
}

/** one cell paragraph: control-char text (legacy first-run-rPr rebuild), rich runs, or null = keep the original paragraph bytes */
export type CellParaPatch =
  | string
  | {
      runs: Run[]
      /** pre-rendered run XML (fields etc.) used instead of runsXml(runs) */
      xml?: string
      /** paragraph format: only the pPr groups that differ from `original` are rebuilt (mergeCellParaFormat) */
      format?: ParaFormat
      original?: ParaFormat
    }
  | null

/** per-cell patch: paragraph patches, or nested-table cell texts by nested index */
export type CellTextsPatch =
  | readonly CellParaPatch[]
  | {
      /** this cell's own text (optional; rewriting the outer text of a cell containing a nested table is not supported yet) */
      paras?: readonly string[] | null
      /** one cell grid per direct nested table (null = leave that nested table untouched) */
      nested?: ReadonlyArray<ReadonlyArray<
        ReadonlyArray<readonly CellParaPatch[] | null | undefined> | null | undefined
      > | null>
    }

/**
 * Replace cell texts inside a <w:tbl> fragment.
 * `texts[row][cell]` = new paragraph strings for that cell (or a CellTextsPatch
 * carrying nested-table cell texts), or null/undefined to leave the cell
 * untouched. Indexes follow document order of w:tr / w:tc (matching
 * TableModel.rows, which includes vMerge-continue cells).
 */
export function patchTableCellTexts(
  tableXml: string,
  texts: ReadonlyArray<ReadonlyArray<CellTextsPatch | null | undefined> | null | undefined>,
): string {
  const trSegs = xmlSegments(tableXml, 'w:tr', 0, tableXml.length)
  let out = ''
  let cursor = 0
  trSegs.forEach((tr, r) => {
    const rowTexts = texts[r]
    if (!rowTexts) return
    const tcSegs = xmlSegments(tableXml, 'w:tc', tr.start, tr.end)
    tcSegs.forEach((tc, c) => {
      const entry = rowTexts[c]
      if (entry == null) return
      const tcXml = tableXml.slice(tc.start, tc.end)
      const patched = Array.isArray(entry)
        ? patchCellXml(tcXml, entry as CellParaPatch[])
        : patchNestedInCell(tcXml, (entry as { nested?: unknown }).nested as never)
      if (patched === null) return
      out += tableXml.slice(cursor, tc.start) + patched
      cursor = tc.end
    })
  })
  return out + tableXml.slice(cursor)
}

/** patch the cell's direct nested tables' cell texts (recursive surgical patch) */
function patchNestedInCell(
  tcXml: string,
  nested:
    | ReadonlyArray<ReadonlyArray<
        ReadonlyArray<readonly CellParaPatch[] | null | undefined> | null | undefined
      > | null>
    | undefined,
): string | null {
  if (!nested || nested.length === 0) return null
  const openTag = /^<w:tc(?: [^>]*)?>/.exec(tcXml)?.[0]
  if (!openTag) return null
  const tblSegs = xmlSegments(tcXml, 'w:tbl', openTag.length, tcXml.length)
  if (tblSegs.length === 0) return null
  let out = ''
  let cursor = 0
  tblSegs.forEach((seg, i) => {
    const grid = nested[i]
    if (!grid) return
    out +=
      tcXml.slice(cursor, seg.start) + patchTableCellTexts(tcXml.slice(seg.start, seg.end), grid)
    cursor = seg.end
  })
  return out + tcXml.slice(cursor)
}

// ---- textbox paragraph patching ----

/**
 * Replacement content for one textbox paragraph. `runs` carry the full rich
 * style (bold, color, size, ...) and are regenerated as fresh OOXML runs.
 * `align`: undefined keeps the original pPr untouched, null removes w:jc,
 * a value rewrites it.
 */
export interface TextboxParaPatch {
  runs: Run[]
  align?: 'left' | 'center' | 'right' | 'justify' | 'distribute' | null
  /** list membership write-back: sets or clears the paragraph's w:numPr.
   * Always explicit (null clears) so leaving a list lands in the file. */
  numPr?: { numId: string; ilvl: number } | null
  /** left indent in twips; null clears w:ind/@w:left, undefined keeps the original */
  indentLeft?: number | null
  /**
   * Tracked paragraph-property revision (w:pPrChange) resolution: 'accept'
   * keeps the current properties and strips the record; 'reject' swaps the
   * whole pPr for the record's old one — Word's reject restores the recorded
   * pre-change state, so unmodeled properties revert faithfully (#2892).
   */
  pPrResolution?: 'accept' | 'reject'
}

/** plain text of one w:p fragment (w:t + tabs/breaks), for change detection */
function paraPlainText(pXml: string): string {
  let out = ''
  const re =
    /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:tab>\s*<\/w:tab>|<w:br\/>|<w:cr\/>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(pXml)) !== null) {
    if (m[0].startsWith('<w:tab')) out += '\t'
    else if (m[0] === '<w:br/>' || m[0] === '<w:cr/>') out += '\n'
    else {
      out += m[1]
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&')
    }
  }
  return out
}

/** Set or clear w:numPr in a paragraph's pPr. CT_PPr order places numPr after
 * pStyle and the keep-together flags, before pBdr, shd, tabs, spacing, ind and
 * jc — so it is inserted after pStyle when present (the common textbox shape)
 * and otherwise directly after the pPr opening tag. */
function pPrWithNumPr(pPr: string, numPr: TextboxParaPatch['numPr']): string {
  const out = pPr.replace(/<w:pPr([^>]*)\/>/, '<w:pPr$1></w:pPr>')
  const stripped = out.replace(/<w:numPr>[\s\S]*?<\/w:numPr>/, '').replace(/<w:numPr\/>/, '')
  if (!numPr) return stripped === '<w:pPr></w:pPr>' ? '' : stripped
  const num = `<w:numPr><w:ilvl w:val="${numPr.ilvl}"/><w:numId w:val="${numPr.numId}"/></w:numPr>`
  const styleEnd = stripped.indexOf('<w:pStyle ')
  const afterStyle = styleEnd === -1 ? -1 : stripped.indexOf('/>', styleEnd)
  const pPrOpen = stripped.indexOf('>')
  const insertAt = afterStyle !== -1 ? afterStyle + 2 : pPrOpen !== -1 ? pPrOpen + 1 : -1
  if (insertAt === -1) return `<w:pPr>${num}</w:pPr>`
  return stripped.slice(0, insertAt) + num + stripped.slice(insertAt)
}

/** Set or clear the left indent on the paragraph's w:ind (other w:ind
 * attributes survive). Word prefers a character-unit left indent, own or from
 * the style chain, over the twips twin, so a twips write also pins
 * w:leftChars="0" the way Word itself writes pt indents. CT_PPr places ind
 * before jc and the paragraph-mark rPr. */
function pPrWithIndLeft(pPr: string, left: number | null): string {
  let out = pPr.replace(/<w:pPr([^>]*)\/>/, '<w:pPr$1></w:pPr>')
  const value = left == null ? '' : ` w:left="${left}" w:leftChars="0"`
  const ind = /<w:ind\b([^>]*)\/>/.exec(out)
  if (ind) {
    const attrs = ind[1].replace(/\s+w:(left|start|leftChars|startChars)="[^"]*"/g, '')
    const next = `${attrs}${value}`
    out = out.replace(ind[0], next.trim() ? `<w:ind${next}/>` : '')
  } else if (left != null) {
    const tag = `<w:ind${value}/>`
    if (!out) return `<w:pPr>${tag}</w:pPr>`
    const at = [out.indexOf('<w:jc '), out.indexOf('<w:rPr')].filter((i) => i !== -1)
    const insertAt = at.length ? Math.min(...at) : out.lastIndexOf('</w:pPr>')
    out = out.slice(0, insertAt) + tag + out.slice(insertAt)
  }
  return out === '<w:pPr></w:pPr>' ? '' : out
}

function pPrWithJc(pPr: string, align: TextboxParaPatch['align']): string {
  const out = pPr.replace(/<w:pPr([^>]*)\/>/, '<w:pPr$1></w:pPr>').replace(/<w:jc [^>]*\/>/, '')
  const jcVal = align === 'justify' ? 'both' : align
  if (!jcVal) return out === '<w:pPr></w:pPr>' ? '' : out
  const jc = `<w:jc w:val="${jcVal}"/>`
  if (!out) return `<w:pPr>${jc}</w:pPr>`
  // w:jc must precede the paragraph-mark rPr inside pPr
  const rPrIdx = out.indexOf('<w:rPr')
  const insertAt = rPrIdx === -1 ? out.lastIndexOf('</w:pPr>') : rPrIdx
  return out.slice(0, insertAt) + jc + out.slice(insertAt)
}

/**
 * Rebuild the paragraphs of one w:txbxContent. `paras[i]` = null keeps the
 * original paragraph bytes; a patch regenerates the paragraph from its rich
 * runs, reusing the nearest original paragraph's pPr (extra new paragraphs
 * inherit the last original's pPr). Original paragraphs beyond paras.length
 * are dropped (paragraph deleted in the editor).
 */
function patchTxbxContent(
  segXml: string,
  paras: ReadonlyArray<TextboxParaPatch | null | undefined>,
): string {
  const open = /^<w:txbxContent(?: [^>]*)?>/.exec(segXml)?.[0]
  if (!open) return segXml
  const pSegs = xmlSegments(segXml, 'w:p', open.length, segXml.length - '</w:txbxContent>'.length)
  if (pSegs.length === 0) return segXml
  const originals = pSegs.map((s) => segXml.slice(s.start, s.end))
  const parts: string[] = []
  for (let i = 0; i < paras.length; i++) {
    const patch = paras[i]
    if (patch == null) {
      if (i < originals.length) parts.push(originals[i])
      continue
    }
    const template = originals[Math.min(i, originals.length - 1)]
    // depth-aware: a w:pPrChange child nests another <w:pPr>, so a plain
    // non-greedy match truncates the element mid-record
    let pPr = rawPprOfParagraph(template)
    if (patch.pPrResolution === 'reject') {
      // the recorded old state wins wholesale; a record without a usable old
      // pPr falls back to stripping it, the modeled old props ride the fields
      const old = pPrChangeOldXml(pPr)
      pPr = old ?? stripPPrChange(pPr)
    } else {
      // the record goes first, so the field patches below meet a clean pPr
      // (the record's own old pPr would otherwise answer their lookups)
      if (patch.pPrResolution === 'accept') pPr = stripPPrChange(pPr)
      if (patch.numPr !== undefined) pPr = pPrWithNumPr(pPr, patch.numPr)
      if (patch.indentLeft !== undefined) pPr = pPrWithIndLeft(pPr, patch.indentLeft)
      if (patch.align !== undefined) pPr = pPrWithJc(pPr, patch.align)
    }
    // no rel allocation inside textboxes: only links that already have an rId
    // survive re-generation, new links degrade to plain text
    parts.push(`<w:p>${pPr}${runsXml(patch.runs, null)}</w:p>`)
  }
  return open + parts.join('') + '</w:txbxContent>'
}

/**
 * Text patches for the anchored boxes of one paragraph fragment. `byIndex` is
 * sparse, indexed by TextboxDisplay.txbxIndex — the ordinal of the box's
 * w:txbxContent among all non-fallback segments, empty ones included, so the
 * mapping survives paragraphs that mix text-bearing boxes with ink-only
 * shapes. `inject` carries the first text of shapes that have no
 * w:txbxContent at all, addressed by their wps:cNvPr id.
 */
export interface TextboxParasPatchSet {
  byIndex?: ReadonlyArray<ReadonlyArray<TextboxParaPatch | null | undefined> | null | undefined>
  inject?: ReadonlyArray<{ shapeId: string; paras: TextboxParaPatch[] }>
}

/**
 * Replace textbox paragraphs inside a paragraph fragment that carries
 * anchored DrawingML textboxes. Word pairs every DrawingML shape with a VML
 * twin inside mc:Fallback whose w:txbxContent duplicates the content —
 * fallback copies are patched with the same paragraphs as the preceding
 * visible box so both renderings stay in sync.
 */
export function patchTextboxParas(paragraphXml: string, patches: TextboxParasPatchSet): string {
  const byIndex = patches.byIndex ?? []
  const segs = xmlSegments(paragraphXml, 'w:txbxContent', 0, paragraphXml.length)
  let xml = paragraphXml
  if (segs.length > 0 && byIndex.length > 0) {
    const fallbacks: Array<{ start: number; end: number }> = []
    const fbRe = /<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g
    let fb: RegExpExecArray | null
    while ((fb = fbRe.exec(paragraphXml)) !== null) {
      fallbacks.push({ start: fb.index, end: fb.index + fb[0].length })
    }
    const inFallback = (pos: number) => fallbacks.some((f) => pos >= f.start && pos < f.end)

    let out = ''
    let cursor = 0
    let ordinal = -1
    let current: ReadonlyArray<TextboxParaPatch | null | undefined> | null = null
    for (const seg of segs) {
      // a fallback twin repeats the preceding visible segment's patch
      if (!inFallback(seg.start)) current = byIndex[++ordinal] ?? null
      if (current == null) continue
      out +=
        paragraphXml.slice(cursor, seg.start) +
        patchTxbxContent(paragraphXml.slice(seg.start, seg.end), current)
      cursor = seg.end
    }
    xml = out + paragraphXml.slice(cursor)
  }
  for (const entry of patches.inject ?? []) xml = injectShapeText(xml, entry)
  return xml
}

/**
 * First text on a shape that has no <w:txbxContent>: inject a fresh wps:txbx
 * into the shape carrying the cNvPr id, before wps:bodyPr per the
 * CT_WordprocessingShape sequence. Word centers shape text, so a bodyPr
 * without an anchor gains anchor="ctr" — matching the editable box's live
 * preview. A shape that already has a txbx (stale inject) has its content
 * rewritten instead. The VML fallback twin is left alone: only pre-2010 Word
 * builds read it, and pairing it up is not worth corrupting on a miss.
 */
function injectShapeText(
  paragraphXml: string,
  entry: { shapeId: string; paras: TextboxParaPatch[] },
): string {
  for (const seg of xmlSegments(paragraphXml, 'wps:wsp', 0, paragraphXml.length)) {
    let shapeXml = paragraphXml.slice(seg.start, seg.end)
    const id = /<wps:cNvPr\b[^>]*\bid="([^"]+)"/.exec(shapeXml)?.[1]
    if (id !== entry.shapeId) continue
    if (shapeXml.includes('<wps:txbx')) {
      const content = xmlSegments(shapeXml, 'w:txbxContent', 0, shapeXml.length)[0]
      if (!content) return paragraphXml
      shapeXml =
        shapeXml.slice(0, content.start) +
        patchTxbxContent(shapeXml.slice(content.start, content.end), entry.paras) +
        shapeXml.slice(content.end)
    } else {
      const body = entry.paras
        .map(
          (p) =>
            `<w:p>${pPrWithJc(pPrWithIndLeft('', p.indentLeft ?? null), p.align)}${runsXml(p.runs, null)}</w:p>`,
        )
        .join('')
      const txbx = `<wps:txbx><w:txbxContent>${body}</w:txbxContent></wps:txbx>`
      const bodyPrAt = shapeXml.search(/<wps:bodyPr[\s/>]/)
      shapeXml =
        bodyPrAt === -1
          ? shapeXml.replace('</wps:wsp>', `${txbx}<wps:bodyPr anchor="ctr"/></wps:wsp>`)
          : shapeXml.slice(0, bodyPrAt) +
            txbx +
            shapeXml
              .slice(bodyPrAt)
              .replace(/<wps:bodyPr\b(?![^>]*\banchor=")/, '<wps:bodyPr anchor="ctr"')
    }
    return paragraphXml.slice(0, seg.start) + shapeXml + paragraphXml.slice(seg.end)
  }
  return paragraphXml
}

/**
 * Resize the shape's own a:ext, the one in its wps:spPr/a:xfrm, keeping the dimension
 * that was not patched. The first a:ext in a drawing is not necessarily the shape's:
 * a textbox that holds a picture carries that picture's a:ext too, so writing the new
 * size there resized the picture and left the box at its old size. A shape without an
 * explicit transform is left alone rather than resized through someone else's a:ext.
 */
function setShapeExt(drawingXml: string, size: { cx?: number; cy?: number }): string {
  const spPr = /<wps:spPr(?:\s[^>]*)?>[\s\S]*?<\/wps:spPr>|<wps:spPr(?:\s[^>]*)?\/>/.exec(
    drawingXml,
  )
  if (!spPr) return drawingXml
  const tag = /<a:ext\b[^>]*?(?:\/>|><\/a:ext>)/.exec(spPr[0])
  if (!tag) return drawingXml
  const cx = size.cx ?? /\bcx="(\d+)"/.exec(tag[0])?.[1]
  const cy = size.cy ?? /\bcy="(\d+)"/.exec(tag[0])?.[1]
  if (cx == null || cy == null) return drawingXml
  const resized = `<a:ext cx="${cx}" cy="${cy}"/>`
  return (
    drawingXml.slice(0, spPr.index) +
    spPr[0].slice(0, tag.index) +
    resized +
    spPr[0].slice(tag.index + tag[0].length) +
    drawingXml.slice(spPr.index + spPr[0].length)
  )
}

/** Resize fixed DrawingML textboxes while preserving their anchors and styling. */
export function patchTextboxHeights(
  paragraphXml: string,
  heightsPx: ReadonlyArray<number | null | undefined>,
): string {
  if (heightsPx.every((height) => height == null)) return paragraphXml
  const drawings = xmlSegments(paragraphXml, 'w:drawing', 0, paragraphXml.length)
  let out = ''
  let cursor = 0
  let boxIndex = -1
  for (const drawing of drawings) {
    const drawingXml = paragraphXml.slice(drawing.start, drawing.end)
    if (!drawingXml.includes('<w:txbxContent') || paraPlainText(drawingXml) === '') continue
    boxIndex++
    const heightPx = heightsPx[boxIndex]
    if (!heightPx) continue
    const cy = Math.max(1, Math.round(heightPx * EMU_PER_PX))
    const resized = setShapeExt(
      drawingXml.replace(/(<wp:extent\b[^>]*\bcy=")\d+(")/, `$1${cy}$2`),
      { cy },
    )
    out += paragraphXml.slice(cursor, drawing.start) + resized
    cursor = drawing.end
  }
  let resizedXml = out + paragraphXml.slice(cursor)

  let fallbackIndex = -1
  resizedXml = resizedXml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, (fallback) => {
    if (!fallback.includes('<w:txbxContent') || paraPlainText(fallback) === '') return fallback
    fallbackIndex++
    const heightPx = heightsPx[fallbackIndex]
    if (!heightPx) return fallback
    return fallback.replace(
      /(style="[^"]*\bheight:)\s*[\d.]+(pt|px)/,
      (_whole, prefix: string, unit: string) => {
        const value =
          unit === 'pt' ? Math.round(heightPx * 75) / 100 : Math.round(heightPx * 100) / 100
        return `${prefix}${value}${unit}`
      },
    )
  })
  return resizedXml
}

/** Connector/line prsts rendered as display boxes despite having no text body */
const LINE_SHAPE_RE =
  /<a:prstGeom[^>]*prst="(?:line|straightConnector1|bentConnector[234]|curvedConnector[234])"/

/** Same drawing set extractTextboxes yields boxes for (keeps patch indexes aligned) */
function isBoxDrawing(drawingXml: string): boolean {
  if (drawingXml.includes('<w:txbxContent') && paraPlainText(drawingXml) !== '') return true
  return LINE_SHAPE_RE.test(drawingXml) && drawingXml.includes('<wp:wrapSquare')
}

export interface TextboxSizePatch {
  wPx?: number | null
  hPx?: number | null
}

/** Resize fixed DrawingML textboxes/shapes while preserving anchors and styling. */
export function patchTextboxSizes(
  paragraphXml: string,
  sizes: ReadonlyArray<TextboxSizePatch | null | undefined>,
): string {
  if (sizes.every((size) => size == null || (size.wPx == null && size.hPx == null)))
    return paragraphXml
  const drawings = xmlSegments(paragraphXml, 'w:drawing', 0, paragraphXml.length)
  let out = ''
  let cursor = 0
  let boxIndex = -1
  for (const drawing of drawings) {
    const drawingXml = paragraphXml.slice(drawing.start, drawing.end)
    if (!isBoxDrawing(drawingXml)) continue
    boxIndex++
    const size = sizes[boxIndex]
    if (!size || (size.wPx == null && size.hPx == null)) continue
    let resized = drawingXml
    if (size.wPx != null) {
      const cx = Math.max(1, Math.round(size.wPx * EMU_PER_PX))
      resized = setShapeExt(resized.replace(/(<wp:extent\b[^>]*\bcx=")\d+(")/, `$1${cx}$2`), {
        cx,
      })
    }
    if (size.hPx != null) {
      const cy = Math.max(1, Math.round(size.hPx * EMU_PER_PX))
      resized = setShapeExt(resized.replace(/(<wp:extent\b[^>]*\bcy=")\d+(")/, `$1${cy}$2`), {
        cy,
      })
      // a fixed height only sticks if Word stops auto-fitting the shape
      resized = resized.replace(
        /<a:spAutoFit\s*\/>|<a:spAutoFit\s*>\s*<\/a:spAutoFit>/,
        '<a:noAutofit/>',
      )
    }
    out += paragraphXml.slice(cursor, drawing.start) + resized
    cursor = drawing.end
  }
  let resizedXml = out + paragraphXml.slice(cursor)

  let fallbackIndex = -1
  resizedXml = resizedXml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, (fallback) => {
    if (!isBoxDrawing(fallback) && !fallback.includes('<w:txbxContent')) return fallback
    if (fallback.includes('<w:txbxContent') && paraPlainText(fallback) === '') return fallback
    fallbackIndex++
    const size = sizes[fallbackIndex]
    if (!size) return fallback
    let next = fallback
    if (size.hPx != null) {
      next = next.replace(
        /(style="[^"]*\bheight:)\s*[\d.]+(pt|px)/,
        (_whole, prefix: string, unit: string) => {
          const value =
            unit === 'pt' ? Math.round(size.hPx! * 75) / 100 : Math.round(size.hPx! * 100) / 100
          return `${prefix}${value}${unit}`
        },
      )
    }
    if (size.wPx != null) {
      next = next.replace(
        /(style="[^"]*\bwidth:)\s*[\d.]+(pt|px)/,
        (_whole, prefix: string, unit: string) => {
          const value =
            unit === 'pt' ? Math.round(size.wPx! * 75) / 100 : Math.round(size.wPx! * 100) / 100
          return `${prefix}${value}${unit}`
        },
      )
    }
    return next
  })
  return resizedXml
}

export interface ShapeStylePatch {
  /** solid fill hex without '#'; null = a:noFill; undefined = keep */
  fillHex?: string | null
  /** outline color hex without '#'; null = no outline; undefined = keep */
  borderHex?: string | null
  /** rotation in degrees clockwise; null or 0 clears a:xfrm@rot; undefined = keep */
  rotDeg?: number | null
  /** outline width in CSS px (a:ln@w = px * 9525 EMU); null clears the attribute; undefined = keep */
  borderWidthPx?: number | null
  /** wps:bodyPr@vert; null is horizontal (the attribute goes); undefined = keep */
  vert?: TextboxVert | null
}

function patchBodyVert(drawingXml: string, vert: TextboxVert | null): string {
  const bodyPr = /<wps:bodyPr\b[^>]*>/.exec(drawingXml)
  if (!bodyPr) {
    return vert
      ? drawingXml.replace('</wps:wsp>', `<wps:bodyPr vert="${vert}"/></wps:wsp>`)
      : drawingXml
  }
  const bare = bodyPr[0].replace(/\s+vert="[^"]*"/, '')
  const tag = vert ? bare.replace(/^<wps:bodyPr\b/, `<wps:bodyPr vert="${vert}"`) : bare
  return drawingXml.replace(bodyPr[0], tag)
}

/**
 * Write a:xfrm@rot (60000ths of a degree clockwise) on a wps:spPr slice.
 * null / 0 removes the attribute so Word saves a clean shape; a slice with no
 * xfrm gets a bare one inserted in its schema position (xfrm comes first).
 */
function patchShapeRotation(spPr: string, rotDeg: number | null): string {
  const norm = rotDeg == null ? 0 : ((rotDeg % 360) + 360) % 360
  // 360deg wraps to 0: never write a full turn
  const emu = Math.round(norm * 60000) % 21600000
  const xfrm = /<a:xfrm\b[^>]*>/.exec(spPr)
  if (!xfrm) {
    if (emu === 0) return spPr
    return spPr.replace(/^(<wps:spPr\b[^>]*>)/, `$1<a:xfrm rot="${emu}"/>`)
  }
  const tag =
    emu === 0
      ? xfrm[0].replace(/\s+rot="[^"]*"/, '')
      : /\brot="/.test(xfrm[0])
        ? xfrm[0].replace(/\brot="[^"]*"/, `rot="${emu}"`)
        : xfrm[0].replace(/^<a:xfrm\b/, `<a:xfrm rot="${emu}"`)
  return spPr.replace(xfrm[0], tag)
}

/**
 * Write a:ln@w (EMU) on a wps:spPr slice. null removes the attribute (Word's
 * own default hairline) and a slice with no a:ln gets one appended in its
 * schema position, after the fill and before any effects.
 */
function patchShapeLineWidth(spPr: string, borderWidthPx: number | null): string {
  const emu = borderWidthPx == null ? null : Math.max(1, Math.round(borderWidthPx * EMU_PER_PX))
  const ln = /<a:ln\b[^>]*>/.exec(spPr)
  if (!ln) {
    if (emu === null) return spPr
    return spPr.replace('</wps:spPr>', `<a:ln w="${emu}"/></wps:spPr>`)
  }
  const tag =
    emu === null
      ? ln[0].replace(/\s+w="[^"]*"/, '')
      : /\bw="/.test(ln[0])
        ? ln[0].replace(/\bw="[^"]*"/, `w="${emu}"`)
        : ln[0].replace(/^<a:ln\b/, `<a:ln w="${emu}"`)
  return spPr.replace(ln[0], tag)
}

/** Replace the first fill slot (a:solidFill/a:noFill) inside an XML slice. */
function replaceFillSlot(xml: string, hex: string | null): string {
  const markup = hex ? `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>` : '<a:noFill/>'
  if (/<a:solidFill>[\s\S]*?<\/a:solidFill>/.test(xml)) {
    return xml.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, markup)
  }
  if (xml.includes('<a:noFill/>')) return xml.replace('<a:noFill/>', markup)
  return xml
}

/** Recolor floating shapes/textboxes/lines (same drawing set as extractTextboxes). */
export function patchShapeStyles(
  paragraphXml: string,
  styles: ReadonlyArray<ShapeStylePatch | null | undefined>,
): string {
  if (
    styles.every(
      (s) =>
        s == null ||
        (s.fillHex === undefined &&
          s.borderHex === undefined &&
          s.rotDeg === undefined &&
          s.borderWidthPx === undefined &&
          s.vert === undefined),
    )
  )
    return paragraphXml
  const drawings = xmlSegments(paragraphXml, 'w:drawing', 0, paragraphXml.length)
  let out = ''
  let cursor = 0
  let boxIndex = -1
  for (const drawing of drawings) {
    let drawingXml = paragraphXml.slice(drawing.start, drawing.end)
    if (!isBoxDrawing(drawingXml)) continue
    boxIndex++
    const style = styles[boxIndex]
    if (
      !style ||
      (style.fillHex === undefined &&
        style.borderHex === undefined &&
        style.rotDeg === undefined &&
        style.borderWidthPx === undefined &&
        style.vert === undefined)
    )
      continue
    const spPrMatch = /<wps:spPr>[\s\S]*?<\/wps:spPr>/.exec(drawingXml)
    if (spPrMatch) {
      let spPr = spPrMatch[0]
      const lnStart = spPr.search(/<a:ln[\s>]/)
      if (style.fillHex !== undefined) {
        // the shape fill slot lives before a:ln; outline colors stay untouched
        const head = lnStart === -1 ? spPr : spPr.slice(0, lnStart)
        const tail = lnStart === -1 ? '' : spPr.slice(lnStart)
        spPr = replaceFillSlot(head, style.fillHex) + tail
      }
      if (style.borderHex !== undefined) {
        const lnMatch = /<a:ln[\s>][\s\S]*?<\/a:ln>/.exec(spPr)
        if (lnMatch) {
          spPr = spPr.replace(lnMatch[0], replaceFillSlot(lnMatch[0], style.borderHex))
        } else if (style.borderHex) {
          spPr = spPr.replace(
            '</wps:spPr>',
            `<a:ln><a:solidFill><a:srgbClr val="${style.borderHex}"/></a:solidFill></a:ln></wps:spPr>`,
          )
        }
      }
      if (style.rotDeg !== undefined) spPr = patchShapeRotation(spPr, style.rotDeg)
      if (style.borderWidthPx !== undefined) spPr = patchShapeLineWidth(spPr, style.borderWidthPx)
      drawingXml = drawingXml.replace(spPrMatch[0], spPr)
    }
    if (style.vert !== undefined) drawingXml = patchBodyVert(drawingXml, style.vert)
    out += paragraphXml.slice(cursor, drawing.start) + drawingXml
    cursor = drawing.end
  }
  return out + paragraphXml.slice(cursor)
}

/** Rewrite the first drawing's extent (chart/SmartArt graphicFrame paragraphs). */
export function patchDrawingExtent(paragraphXml: string, wPx: number, hPx: number): string {
  const cx = Math.max(1, Math.round(wPx * EMU_PER_PX))
  const cy = Math.max(1, Math.round(hPx * EMU_PER_PX))
  return paragraphXml
    .replace(/(<wp:extent\b[^>]*\bcx=")\d+(")/, `$1${cx}$2`)
    .replace(/(<wp:extent\b[^>]*\bcy=")\d+(")/, `$1${cy}$2`)
    .replace(/(<a:ext\b[^>]*\bcx=")\d+(")/, `$1${cx}$2`)
    .replace(/(<a:ext\b[^>]*\bcy=")\d+(")/, `$1${cy}$2`)
}

/** Insertable line/connector kinds: stroke-only wps:wsp with optional arrow ends */
export const LINE_KINDS: Record<string, { prst: string; head?: boolean; tail?: boolean }> = {
  line: { prst: 'line' },
  lineArrow: { prst: 'straightConnector1', tail: true },
  lineArrowDouble: { prst: 'straightConnector1', head: true, tail: true },
  lineBent: { prst: 'bentConnector3' },
  lineCurved: { prst: 'curvedConnector3' },
}

/** Insert a floating stroke-only line/connector paragraph (wp:anchor + wps:wsp). */
export function buildLineParagraphXml(opts: {
  kind: string
  widthEmu?: number
  heightEmu?: number
  id?: number
  colorHex?: string
}): string {
  const def = LINE_KINDS[opts.kind] ?? LINE_KINDS.line
  const widthEmu = opts.widthEmu ?? 1800000
  const heightEmu = opts.heightEmu ?? 114300
  const id = opts.id ?? 1
  const colorHex = opts.colorHex ?? '000000'

  const ln =
    `<a:ln w="12700"><a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill>` +
    (def.head ? '<a:headEnd type="triangle"/>' : '') +
    (def.tail ? '<a:tailEnd type="triangle"/>' : '') +
    `</a:ln>`

  const spPr =
    `<wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    `<a:prstGeom prst="${def.prst}"><a:avLst/></a:prstGeom>` +
    `<a:noFill/>` +
    ln +
    `</wps:spPr>`

  const wsp =
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:cNvSpPr/>` +
    spPr +
    `<wps:bodyPr/>` +
    `</wps:wsp>`

  const graphicData =
    `<a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${wsp}</a:graphicData>`

  const graphic = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${graphicData}</a:graphic>`

  const anchor =
    `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapSquare wrapText="bothSides"/>` +
    `<wp:docPr id="${id}" name="${def.prst} ${id}"/>` +
    graphic +
    `</wp:anchor>`

  const mcNs =
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
  const mcChoice = `<mc:Choice Requires="wps"><w:drawing>${anchor}</w:drawing></mc:Choice>`
  const vmlLine =
    `<v:line xmlns:v="urn:schemas-microsoft-com:vml" ` +
    `from="0,0" to="${Math.round(widthEmu / EMU_PER_PT)}pt,${Math.round(heightEmu / EMU_PER_PT)}pt" ` +
    `strokecolor="#${colorHex}"/>`
  const mcFallback = `<mc:Fallback><w:pict>${vmlLine}</w:pict></mc:Fallback>`

  return `<w:p><w:r><mc:AlternateContent ${mcNs}>${mcChoice}${mcFallback}</mc:AlternateContent></w:r></w:p>`
}

/** CT_PPr child sequence (subset), for schema-ordered assembly and merging */
export const PPR_CHILD_ORDER = [
  'w:pStyle',
  'w:keepNext',
  'w:keepLines',
  'w:pageBreakBefore',
  'w:framePr',
  'w:widowControl',
  'w:numPr',
  'w:suppressLineNumbers',
  'w:pBdr',
  'w:shd',
  'w:tabs',
  'w:suppressAutoHyphens',
  'w:kinsoku',
  'w:wordWrap',
  'w:overflowPunct',
  'w:topLinePunct',
  'w:autoSpaceDE',
  'w:autoSpaceDN',
  'w:bidi',
  'w:adjustRightInd',
  'w:snapToGrid',
  'w:spacing',
  'w:ind',
  'w:contextualSpacing',
  'w:mirrorIndents',
  'w:suppressOverlap',
  'w:jc',
  'w:textDirection',
  'w:textAlignment',
  'w:textboxTightWrap',
  'w:outlineLvl',
  'w:divId',
  'w:cnfStyle',
  'w:rPr',
  'w:sectPr',
  'w:pPrChange',
]

interface PPrChild {
  name: string
  xml: string
}

/** w:spacing element for a ParaFormat ('' when nothing is set) — shared by body and table-cell paragraphs */
function paraSpacingXml(format: ParaFormat): string {
  const attrs: string[] = []
  if (format.spaceBefore && format.spaceBefore > 0) {
    attrs.push(`w:before="${Math.round(format.spaceBefore)}"`)
  }
  // explicit "0" must be written back: a bare literal would not override a
  // style-chain autospacing (Word resolves the attributes independently)
  if (format.spaceBeforeAuto !== undefined)
    attrs.push(`w:beforeAutospacing="${format.spaceBeforeAuto ? 1 : 0}"`)
  // explicit 0 overrides the style's space-after in Word, so it must be written back
  if (format.spaceAfter !== undefined && format.spaceAfter >= 0) {
    attrs.push(`w:after="${Math.round(format.spaceAfter)}"`)
  }
  if (format.spaceAfterAuto !== undefined)
    attrs.push(`w:afterAutospacing="${format.spaceAfterAuto ? 1 : 0}"`)
  if ((format.lineRule === 'exact' || format.lineRule === 'atLeast') && format.lineRawTwips) {
    // Exact/at-least line height: write the twips back verbatim (previously only auto was
    // recognized, so edited paragraphs lost their line spacing)
    attrs.push(`w:line="${Math.round(format.lineRawTwips)}"`, `w:lineRule="${format.lineRule}"`)
  } else if (format.lineSpacing && format.lineSpacing > 0) {
    attrs.push(`w:line="${Math.round(format.lineSpacing * 240)}"`, 'w:lineRule="auto"')
  }
  return attrs.length > 0 ? `<w:spacing ${attrs.join(' ')}/>` : ''
}

/** pPr children the ParaFormat model owns (rebuilt on format edits) */
function formatPPrChildren(format: ParaFormat | undefined): PPrChild[] {
  if (!format) return []
  const out: PPrChild[] = []
  if (format.pageBreakBefore) out.push({ name: 'w:pageBreakBefore', xml: '<w:pageBreakBefore/>' })
  for (const tag of PPR_FLAG_TAGS) {
    const v = format[PPR_FLAG_FIELDS[tag]]
    if (v !== undefined) out.push({ name: tag, xml: v ? `<${tag}/>` : `<${tag} w:val="0"/>` })
  }
  if (format.borders) {
    const style = format.borderStyle
    const defaultSz = Math.max(2, Math.round(style?.szEighths ?? 4))
    // ECMA-376 17.3.4: an omitted w:space is 0; the renderer pads an undeclared side by the same 0
    const space = Math.min(31, Math.max(0, Math.round(style?.spacePt ?? 0)))
    const defaultColor = style?.color ? escapeXmlAttr(style.color) : 'auto'
    const line = (side: string, ch: 't' | 'b' | 'l' | 'r') => {
      const declared = format.borderLines?.[ch]
      const sz = declared?.szPt ? Math.max(1, Math.round(declared.szPt * 8)) : defaultSz
      const color = declared?.color ? escapeXmlAttr(declared.color) : defaultColor
      const sp =
        declared?.spacePt !== undefined
          ? Math.min(31, Math.max(0, Math.round(declared.spacePt)))
          : space
      const val = declared?.val ? escapeXmlAttr(declared.val) : 'single'
      const fx = `${declared?.shadow ? ' w:shadow="1"' : ''}${declared?.frame ? ' w:frame="1"' : ''}`
      return `<w:${side} w:val="${val}" w:sz="${sz}" w:space="${sp}" w:color="${color}"${fx}/>`
    }
    const sides: string[] = []
    // schema order inside pBdr: top, left, bottom, right
    if (format.borders.includes('t')) sides.push(line('top', 't'))
    if (format.borders.includes('l')) sides.push(line('left', 'l'))
    if (format.borders.includes('b')) sides.push(line('bottom', 'b'))
    if (format.borders.includes('r')) sides.push(line('right', 'r'))
    if (sides.length > 0) out.push({ name: 'w:pBdr', xml: `<w:pBdr>${sides.join('')}</w:pBdr>` })
  }
  if (format.shadingFill || format.shadingPattern) {
    const pattern = format.shadingPattern
    const val = pattern ? escapeXmlAttr(pattern.val) : 'clear'
    const color = pattern?.color ? escapeXmlAttr(pattern.color) : 'auto'
    const fill = format.shadingFill ? escapeXmlAttr(format.shadingFill) : 'auto'
    out.push({ name: 'w:shd', xml: `<w:shd w:val="${val}" w:color="${color}" w:fill="${fill}"/>` })
  }
  if (format.bidi) out.push({ name: 'w:bidi', xml: '<w:bidi/>' })
  const spacing = paraSpacingXml(format)
  if (spacing) out.push({ name: 'w:spacing', xml: spacing })
  const indAttrs: string[] = []
  // w:left/w:right are signed in OOXML — negative indents (text extending
  // into the margin) are valid and must survive a paragraph rebuild; the
  // old > 0 guard silently dropped them, shifting rebuilt paragraphs
  // rightward on save.
  // explicit w:left="0" must be written back: it cancels a numbering-level indent
  if (format.indentLeft !== undefined) indAttrs.push(`w:left="${Math.round(format.indentLeft)}"`)
  if (format.indentRight !== undefined) indAttrs.push(`w:right="${Math.round(format.indentRight)}"`)
  // a character-unit special indent rides along its twips twin (Word writes
  // both; the *Chars attribute wins and rescales with the first run's size —
  // CJK "first line: 2 characters", issue #1892)
  if (format.charIndents?.firstLine)
    indAttrs.push(`w:firstLineChars="${Math.round(format.charIndents.firstLine)}"`)
  if (format.indentFirstLine !== undefined) {
    if (format.indentFirstLine >= 0)
      indAttrs.push(`w:firstLine="${Math.round(format.indentFirstLine)}"`)
    else indAttrs.push(`w:hanging="${Math.round(-format.indentFirstLine)}"`)
  }
  if (indAttrs.length > 0) out.push({ name: 'w:ind', xml: `<w:ind ${indAttrs.join(' ')}/>` })
  if (format.align) {
    let jc: string = format.align === 'justify' ? 'both' : format.align
    // convert visual value back to Word's logical value (inverse of the parse side)
    if ((format.bidi || format.bidiInherited) && (jc === 'left' || jc === 'right'))
      jc = jc === 'left' ? 'right' : 'left'
    out.push({ name: 'w:jc', xml: `<w:jc w:val="${jc}"/>` })
  }
  // rel stops (w:ptab mirrors) and style-inherited stops are display-only: never written into w:tabs
  const realStops = (format.tabStops ?? []).filter((ts) => !ts.rel && !ts.inherited)
  if (realStops.length > 0) {
    const tabXml = realStops
      .map((ts) => {
        let xml = `<w:tab w:val="${escapeXmlAttr(ts.val)}" w:pos="${ts.pos}"`
        if (ts.leader && ts.leader !== 'none') xml += ` w:leader="${escapeXmlAttr(ts.leader)}"`
        return xml + '/>'
      })
      .join('')
    out.push({ name: 'w:tabs', xml: `<w:tabs>${tabXml}</w:tabs>` })
  }
  if (format.textDirection) {
    out.push({ name: 'w:textDirection', xml: `<w:textDirection w:val="${format.textDirection}"/>` })
  }
  if (format.textAlignment) {
    out.push({
      name: 'w:textAlignment',
      xml: `<w:textAlignment w:val="${escapeXmlAttr(format.textAlignment)}"/>`,
    })
  }
  if (format.frame) {
    out.push({ name: 'w:framePr', xml: framePrXml(format.frame) })
  } else if (format.dropCap) {
    const { type, lines, hSpace } = format.dropCap
    const gap = hSpace && hSpace > 0 ? ` w:hSpace="${Math.round(hSpace)}"` : ''
    out.push({
      name: 'w:framePr',
      xml: `<w:framePr w:dropCap="${type}" w:lines="${lines}"${gap} w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>`,
    })
  }
  if (format.emptyRunSizeHalfPoints) {
    const sz = Math.round(format.emptyRunSizeHalfPoints)
    out.push({ name: 'w:rPr', xml: `<w:rPr><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>` })
  }
  return out
}

/** positioned paragraph frame (P19): w:framePr with absolute x/y, Word's attribute order */
function framePrXml(frame: ParaFrame): string {
  const attrs = [`w:w="${Math.round(frame.wTwips)}"`]
  if (frame.hTwips !== undefined) {
    attrs.push(`w:h="${Math.round(frame.hTwips)}"`, `w:hRule="${frame.hRule ?? 'atLeast'}"`)
  }
  if (frame.vSpaceTwips) attrs.push(`w:vSpace="${Math.round(frame.vSpaceTwips)}"`)
  if (frame.hSpaceTwips) attrs.push(`w:hSpace="${Math.round(frame.hSpaceTwips)}"`)
  attrs.push(
    `w:wrap="${frame.wrap ?? 'none'}"`,
    `w:vAnchor="${frame.vAnchor ?? 'page'}"`,
    `w:hAnchor="${frame.hAnchor ?? 'page'}"`,
    `w:x="${Math.round(frame.xTwips)}"`,
  )
  if (frame.xAlign) attrs.push(`w:xAlign="${frame.xAlign}"`)
  attrs.push(`w:y="${Math.round(frame.yTwips)}"`)
  if (frame.yAlign) attrs.push(`w:yAlign="${frame.yAlign}"`)
  if (frame.anchorLock) attrs.push('w:anchorLock="1"')
  return `<w:framePr ${attrs.join(' ')}/>`
}

/** tri-state pPr flags: undefined = not edited (raw bytes kept), false = explicit w:val="0" overriding the style */
const PPR_FLAG_FIELDS = {
  'w:keepNext': 'keepNext',
  'w:keepLines': 'keepLines',
  'w:widowControl': 'widowControl',
  'w:suppressLineNumbers': 'suppressLineNumbers',
  'w:contextualSpacing': 'contextualSpacing',
  'w:kinsoku': 'kinsoku',
  'w:wordWrap': 'wordWrapDirect',
  'w:topLinePunct': 'topLinePunct',
  'w:overflowPunct': 'overflowPunctDirect',
  'w:autoSpaceDE': 'autoSpaceDE',
  'w:autoSpaceDN': 'autoSpaceDN',
  'w:suppressAutoHyphens': 'suppressAutoHyphens',
  'w:snapToGrid': 'snapToGrid',
  'w:mirrorIndents': 'mirrorIndents',
} as const
const PPR_FLAG_TAGS = Object.keys(PPR_FLAG_FIELDS) as Array<keyof typeof PPR_FLAG_FIELDS>

const FORMAT_MANAGED_TAGS = new Set([
  'w:pageBreakBefore',
  'w:pBdr',
  'w:shd',
  'w:bidi',
  'w:spacing',
  'w:ind',
  'w:jc',
])

/** top-level child elements of an XML fragment (depth-aware, attrs kept) */
export function splitXmlChildren(xml: string): PPrChild[] {
  const out: PPrChild[] = []
  const tagRe = /<(\/?)([A-Za-z0-9:._-]+)((?:"[^"]*"|'[^']*'|[^"'>])*)>/g
  let depth = 0
  let start = -1
  let name = ''
  let match: RegExpExecArray | null
  while ((match = tagRe.exec(xml)) !== null) {
    const closing = match[1] === '/'
    const selfClosing = match[3].endsWith('/')
    if (closing) {
      depth--
      if (depth === 0) out.push({ name, xml: xml.slice(start, match.index + match[0].length) })
    } else if (selfClosing) {
      if (depth === 0) out.push({ name: match[2], xml: match[0] })
    } else {
      if (depth === 0) {
        start = match.index
        name = match[2]
      }
      depth++
    }
  }
  return out
}

/** mirrors JC_ALIGN on the parse side (kept local: parse.ts imports from this module) */
const JC_TO_ALIGN: Record<string, ParaFormat['align']> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  both: 'justify',
  lowKashida: 'justify',
  mediumKashida: 'justify',
  highKashida: 'justify',
  thaiDistribute: 'justify',
  distribute: 'distribute',
}

function rawSpacingUnchanged(raw: string | undefined, f: ParaFormat): boolean {
  const rule = rawAttr(raw, 'w:lineRule') ?? 'auto'
  const line = parseInt(rawAttr(raw, 'w:line') ?? '', 10)
  const before = parseInt(rawAttr(raw, 'w:before') ?? '', 10)
  const afterStr = rawAttr(raw, 'w:after')
  const after = parseInt(afterStr ?? '', 10)
  const rawBefore = before > 0 ? before : undefined
  const rawAfter = afterStr !== undefined && after >= 0 ? after : undefined
  const fBefore = f.spaceBefore && f.spaceBefore > 0 ? Math.round(f.spaceBefore) : undefined
  const fAfter =
    f.spaceAfter !== undefined && f.spaceAfter >= 0 ? Math.round(f.spaceAfter) : undefined
  if (rawBefore !== fBefore || rawAfter !== fAfter) return false
  const rawAuto = (v: string | undefined): boolean | undefined =>
    v === undefined ? undefined : v === '1' || v === 'true'
  if (rawAuto(rawAttr(raw, 'w:beforeAutospacing')) !== f.spaceBeforeAuto) return false
  if (rawAuto(rawAttr(raw, 'w:afterAutospacing')) !== f.spaceAfterAuto) return false
  if ((f.lineRule === 'exact' || f.lineRule === 'atLeast') && f.lineRawTwips) {
    return line > 0 && rule === f.lineRule && line === Math.round(f.lineRawTwips)
  }
  if (f.lineSpacing && f.lineSpacing > 0) {
    return line > 0 && rule === 'auto' && Math.round((line / 240) * 100) / 100 === f.lineSpacing
  }
  return !(line > 0)
}

function rawIndUnchanged(raw: string | undefined, f: ParaFormat): boolean {
  const left = parseInt(rawAttr(raw, 'w:left') ?? rawAttr(raw, 'w:start') ?? '', 10)
  const right = parseInt(rawAttr(raw, 'w:right') ?? rawAttr(raw, 'w:end') ?? '', 10)
  const firstLine = parseInt(rawAttr(raw, 'w:firstLine') ?? '', 10)
  const hanging = parseInt(rawAttr(raw, 'w:hanging') ?? '', 10)
  // mirrors the parse side: an explicit 0 stays distinct from absent
  const rawLeft = Number.isFinite(left) ? left : undefined
  const rawRight = Number.isFinite(right) ? right : undefined
  const rawFirst =
    hanging > 0
      ? -hanging
      : firstLine > 0
        ? firstLine
        : Number.isFinite(firstLine) || Number.isFinite(hanging)
          ? 0
          : undefined
  const norm = (v: number | undefined) => (v !== undefined ? Math.round(v) : undefined)
  return (
    rawLeft === norm(f.indentLeft) &&
    rawRight === norm(f.indentRight) &&
    rawFirst === norm(f.indentFirstLine)
  )
}

const rawOn = (v: string | undefined): boolean => v === '1' || v === 'true' || v === 'on'

function rawPBdrUnchanged(raw: string | undefined, f: ParaFormat): boolean {
  let rawBorders = ''
  const rawLines: NonNullable<ParaFormat['borderLines']> = {}
  if (raw) {
    const inner = raw.replace(/^<w:pBdr[^>]*>/, '').replace(/<\/w:pBdr>$/, '')
    const kids = splitXmlChildren(inner)
    for (const [side, ch] of [
      ['top', 't'],
      ['bottom', 'b'],
      ['left', 'l'],
      ['right', 'r'],
    ] as const) {
      const el = kids.find((k) => k.name === `w:${side}`)
      // mirror extractParaFormat: nil is a reset, not a border, or the raw/model
      // comparison never matches and the pBdr always gets rebuilt from the model
      const val = el ? rawAttr(el.xml, 'w:val') : undefined
      if (!el || val === 'none' || val === 'nil') continue
      rawBorders += ch
      const color = rawAttr(el.xml, 'w:color')
      const sz = parseInt(rawAttr(el.xml, 'w:sz') ?? '', 10)
      const line: NonNullable<ParaFormat['borderLines']>[typeof ch] = {}
      if (color && color !== 'auto') line.color = color
      if (Number.isFinite(sz) && sz > 0) line.szPt = sz / 8
      const space = parseInt(rawAttr(el.xml, 'w:space') ?? '', 10)
      if (Number.isFinite(space) && space > 0) line.spacePt = space
      if (val && val !== 'single') line.val = val
      if (rawOn(rawAttr(el.xml, 'w:shadow'))) line.shadow = true
      if (rawOn(rawAttr(el.xml, 'w:frame'))) line.frame = true
      if (Object.keys(line).length > 0) rawLines[ch] = line
    }
  }
  const norm = (s: string | undefined) => (s ? [...new Set(s)].sort().join('') : '')
  if (norm(rawBorders) !== norm(f.borders)) return false
  const normLines = (lines: ParaFormat['borderLines']) =>
    JSON.stringify(
      (['t', 'b', 'l', 'r'] as const).map((ch) => [
        lines?.[ch]?.color ?? null,
        lines?.[ch]?.szPt ?? null,
        lines?.[ch]?.spacePt ?? null,
        lines?.[ch]?.val ?? null,
        !!lines?.[ch]?.shadow,
        !!lines?.[ch]?.frame,
      ]),
    )
  return normLines(rawLines) === normLines(f.borderLines)
}

function rawTabsUnchanged(raw: string | undefined, allStops: TabStop[]): boolean {
  // display-only w:ptab mirrors and style-inherited stops are not part of w:tabs
  const stops = allStops.filter((s) => !s.rel && !s.inherited)
  const rawStops: TabStop[] = []
  if (raw) {
    const inner = raw.replace(/^<w:tabs[^>]*>/, '').replace(/<\/w:tabs>$/, '')
    for (const kid of splitXmlChildren(inner)) {
      if (kid.name !== 'w:tab') continue
      const pos = parseInt(rawAttr(kid.xml, 'w:pos') ?? '', 10)
      if (!Number.isFinite(pos)) continue
      rawStops.push({ pos, val: (rawAttr(kid.xml, 'w:val') ?? 'left') as TabStop['val'] })
      const leader = rawAttr(kid.xml, 'w:leader')
      if (leader && leader !== 'none') rawStops[rawStops.length - 1].leader = leader as never
    }
  }
  if (rawStops.length !== stops.length) return false
  const normLeader = (l?: string) => (l && l !== 'none' ? l : undefined)
  return rawStops.every(
    (r, i) =>
      r.pos === stops[i].pos &&
      r.val === stops[i].val &&
      normLeader(r.leader) === normLeader(stops[i].leader),
  )
}

function rawFramePrUnchanged(raw: string | undefined, f: ParaFormat): boolean {
  // positioned frames (P19) rebuild whenever set: comparing every attribute
  // buys nothing — the generated framePr is byte-stable already
  if (f.frame !== undefined) return raw !== undefined && raw === framePrXml(f.frame)
  const val = rawAttr(raw, 'w:dropCap')
  const rawDc =
    val === 'drop' || val === 'margin'
      ? { type: val, lines: parseInt(rawAttr(raw, 'w:lines') ?? '3', 10) || 3 }
      : undefined
  const rawGap = rawDc ? parseInt(rawAttr(raw, 'w:hSpace') ?? '', 10) || 0 : 0
  return (
    rawDc?.type === f.dropCap?.type &&
    rawDc?.lines === f.dropCap?.lines &&
    rawGap === (f.dropCap?.hSpace ?? 0)
  )
}

/** True when re-parsing the raw child yields the current model value, i.e. the group was not edited */
/**
 * Same indent in two format models (twips, rounded like the emitted attributes).
 * Used when the caller knows the paragraph's parsed format: a raw w:ind whose
 * character-unit attributes (w:firstLineChars…) resolved to the model's twips
 * has no twips twin to compare against, so raw-vs-model would rebuild it and
 * lose the character unit on an unrelated format edit.
 */
function sameIndent(a: ParaFormat, b: ParaFormat): boolean {
  const norm = (v: number | undefined) => (v !== undefined ? Math.round(v) : undefined)
  return (
    norm(a.indentLeft) === norm(b.indentLeft) &&
    norm(a.indentRight) === norm(b.indentRight) &&
    norm(a.indentFirstLine) === norm(b.indentFirstLine)
  )
}

/**
 * Cancel attributes for the character-unit indents a paragraph was laid out
 * with but the new format no longer carries. A rebuilt w:ind carries twips,
 * and Word keeps preferring a `*Chars` from the style chain over a twips twin
 * (probed) — so dropping a character indent must write the explicit
 * `w:firstLineChars="0"` Word itself writes for pt indents, or the style's
 * character indent supersedes the edit on reload. Components the new format
 * still carries are re-emitted in the rebuilt w:ind (the character unit edits
 * in place) and must NOT be cancelled.
 */
function charIndentCancelAttrs(original: CharIndents, format: ParaFormat | undefined): string[] {
  const kept = format?.charIndents
  const out: string[] = []
  if (original.left && !kept?.left) out.push('w:leftChars="0"')
  if (original.right && !kept?.right) out.push('w:rightChars="0"')
  // firstLineChars and hangingChars are one component (the special indent):
  // a new character special replaces the old one outright — no cancel
  if ((original.hanging || original.firstLine) && !kept?.hanging && !kept?.firstLine) {
    out.push(original.hanging ? 'w:hangingChars="0"' : 'w:firstLineChars="0"')
  }
  return out
}

function pprGroupUnchanged(
  tag: string,
  raw: string | undefined,
  f: ParaFormat,
  original?: ParaFormat,
): boolean {
  switch (tag) {
    case 'w:pageBreakBefore':
      return rawBool(raw) === !!f.pageBreakBefore
    case 'w:keepNext':
    case 'w:keepLines':
    case 'w:widowControl':
    case 'w:suppressLineNumbers':
    case 'w:contextualSpacing':
    case 'w:kinsoku':
    case 'w:topLinePunct':
    case 'w:wordWrap':
    case 'w:overflowPunct':
    case 'w:autoSpaceDE':
    case 'w:autoSpaceDN':
    case 'w:suppressAutoHyphens':
    case 'w:snapToGrid':
    case 'w:mirrorIndents': {
      const field = PPR_FLAG_FIELDS[tag]
      const v = f[field]
      // model lost a flag the parsed paragraph had: the user reset it to the style
      if (v === undefined) return original?.[field] === undefined
      return raw !== undefined && rawBool(raw) === v
    }
    case 'w:bidi':
      return rawBool(raw) === !!f.bidi
    case 'w:jc': {
      const val = rawAttr(raw, 'w:val')
      let rawAlign = val ? JC_TO_ALIGN[val] : undefined
      if ((f.bidi || f.bidiInherited) && (rawAlign === 'left' || rawAlign === 'right')) {
        rawAlign = rawAlign === 'left' ? 'right' : 'left'
      }
      return rawAlign === f.align
    }
    case 'w:spacing':
      return rawSpacingUnchanged(raw, f)
    case 'w:ind':
      if (original !== undefined) {
        if (sameIndent(original, f)) return true
        // laid out with character-unit indents (possibly from the style alone, with no
        // raw w:ind at all): the edit must rebuild so the cancel attributes get written
        if (original.charIndents) return false
      }
      return rawIndUnchanged(raw, f)
    case 'w:pBdr':
      return rawPBdrUnchanged(raw, f)
    case 'w:shd': {
      const fill = rawAttr(raw, 'w:fill')
      const val = rawAttr(raw, 'w:val')
      const pattern = val && val !== 'clear' && val !== 'nil' ? val : undefined
      const color = pattern ? rawAttr(raw, 'w:color')?.replace(/^#/, '') : undefined
      return (
        (fill && fill !== 'auto' ? fill : undefined) === f.shadingFill &&
        pattern === f.shadingPattern?.val &&
        (color && color !== 'auto' ? color : undefined) === f.shadingPattern?.color
      )
    }
    case 'w:tabs':
      return rawTabsUnchanged(raw, f.tabStops ?? [])
    case 'w:framePr':
      return rawFramePrUnchanged(raw, f)
    case 'w:rPr': {
      const m = raw ? /<w:sz\b[^>]*w:val="(\d+)"/.exec(raw) : null
      return (m ? parseInt(m[1], 10) : undefined) === f.emptyRunSizeHalfPoints
    }
    case 'w:textAlignment':
      return rawAttr(raw, 'w:val') === f.textAlignment
    default:
      return false
  }
}

/**
 * Merge format-model changes into an original <w:pPr> slice. Like mergeRPrModel,
 * each managed child is compared group by group: when its raw bytes re-parse to
 * the current model value the group was not edited and keeps its original bytes,
 * so unmodeled attributes (w:firstLineChars, w:afterLines, autospacing, border
 * colors, shading patterns…) survive. Only genuinely changed groups are rebuilt
 * from the model (at their schema position) — a rebuilt w:ind is written in twips
 * and, when the paragraph was laid out with character-unit indents, carries the
 * `*Chars="0"` that cancels them, so the user's new indent wins over the CJK
 * char-unit variant Word would otherwise prefer. Everything unmanaged — keepNext,
 * paragraph-mark rPr, pPrChange revision records... — keeps its original bytes.
 * When format.tabStops is set, w:tabs is also managed (replaced or removed).
 * When format.dropCap is set, w:framePr is also managed.
 * `original` is the paragraph's parsed format when the caller has it: an indent
 * equal to it keeps its raw bytes even when those bytes carry no twips twin
 * (character-unit indents resolved at parse time), and its `charIndents` say
 * which cancel attributes an indent edit has to write.
 */
/** pPr groups a header/footer cell paragraph models, by the format fields that feed them */
const CELL_PPR_GROUPS: Array<[tag: string, fields: Array<keyof ParaFormat>]> = [
  ['w:jc', ['align']],
  ['w:spacing', ['spaceBefore', 'spaceAfter', 'lineSpacing', 'lineRule', 'lineRawTwips']],
  ['w:ind', ['indentLeft', 'indentFirstLine']],
  ['w:tabs', ['tabStops']],
  ['w:kinsoku', ['kinsoku']],
  ['w:wordWrap', ['wordWrapDirect']],
  ['w:topLinePunct', ['topLinePunct']],
  ['w:overflowPunct', ['overflowPunctDirect']],
  ['w:autoSpaceDE', ['autoSpaceDE']],
  ['w:autoSpaceDN', ['autoSpaceDN']],
]

/** attributes of w:ind / w:spacing the cell model drives, by group with the
 *  ParaFormat fields that drive it: only a group whose fields changed is rebuilt
 *  (a cleared field goes away with its character-unit and autospacing twins);
 *  the other groups, and attributes outside any group (right indent, …), keep
 *  their raw value — the parser drops explicit zeros, so an untouched
 *  `w:after="0"` must never be regenerated from the model */
const CELL_MODELED_ATTR_GROUPS: Record<
  string,
  Array<{ attrs: string[]; fields: Array<keyof ParaFormat> }>
> = {
  'w:ind': [
    { attrs: ['w:left', 'w:start', 'w:leftChars', 'w:startChars'], fields: ['indentLeft'] },
    {
      attrs: ['w:firstLine', 'w:hanging', 'w:firstLineChars', 'w:hangingChars'],
      fields: ['indentFirstLine'],
    },
  ],
  'w:spacing': [
    { attrs: ['w:before', 'w:beforeLines', 'w:beforeAutospacing'], fields: ['spaceBefore'] },
    { attrs: ['w:after', 'w:afterLines', 'w:afterAutospacing'], fields: ['spaceAfter'] },
    { attrs: ['w:line', 'w:lineRule'], fields: ['lineSpacing', 'lineRule', 'lineRawTwips'] },
  ],
}

function elementAttrs(xml: string): Array<[string, string]> {
  return [...xml.matchAll(/([\w:]+)="([^"]*)"/g)].map((m) => [m[1], m[2]])
}

/** replace one self-closing pPr child by attribute merge: the groups in
 *  `rebuilt` come from `freshEl` (absent group = cleared), everything else keeps
 *  its raw value; an element left without attributes is dropped, a new one lands
 *  in schema order */
function mergeCellElementAttrs(
  pPr: string,
  tag: string,
  freshEl: string | undefined,
  rebuilt: ReadonlySet<string>,
): string {
  const elRe = new RegExp(`<${tag}(?:\\s[^>]*)?/>`)
  const rawEl = elRe.exec(pPr)?.[0]
  const fresh = elementAttrs(freshEl ?? '').filter(([n]) => rebuilt.has(n))
  const freshNames = new Set(fresh.map(([n]) => n))
  const kept = elementAttrs(rawEl ?? '').filter(([n]) => !rebuilt.has(n) && !freshNames.has(n))
  const attrs = [...fresh, ...kept]
  if (attrs.length === 0) return rawEl ? pPr.replace(rawEl, '') : pPr
  const merged = `<${tag} ${attrs.map(([n, v]) => `${n}="${v}"`).join(' ')}/>`
  if (rawEl) return pPr.replace(rawEl, merged)
  const open = /^<w:pPr(?: [^>]*)?>/.exec(pPr)?.[0] ?? '<w:pPr>'
  const inner = pPr.startsWith(open) ? pPr.slice(open.length, pPr.length - '</w:pPr>'.length) : ''
  const rank = (n: string) => PPR_CHILD_ORDER.indexOf(n)
  const children = splitXmlChildren(inner).map((c) => c.xml)
  const names = splitXmlChildren(inner).map((c) => c.name)
  let at = children.length
  for (let i = 0; i < names.length; i++) {
    if (rank(names[i]) > rank(tag)) {
      at = i
      break
    }
  }
  children.splice(at, 0, merged)
  return `${open}${children.join('')}</w:pPr>`
}

/**
 * Paragraph format edit inside a layout-table cell: only the groups whose
 * fields differ from `original` are rewritten, so unmodeled pPr children
 * (bidi, borders, shading, page break) and style-resolved values the editor
 * merely reports stay as they were. w:jc / w:tabs go through mergePPrFormat;
 * w:ind / w:spacing merge attribute-wise so their unmodeled attributes survive.
 */
export function mergeCellParaFormat(
  rawPPr: string,
  format: ParaFormat,
  original: ParaFormat | undefined,
): string {
  const differs = (k: keyof ParaFormat) =>
    JSON.stringify(format[k] ?? null) !== JSON.stringify(original?.[k] ?? null)
  const tags = new Set(CELL_PPR_GROUPS.filter(([, f]) => f.some(differs)).map(([tag]) => tag))
  if (tags.size === 0) return rawPPr
  let out = !rawPPr || /^<w:pPr(?:\s[^>]*)?\/>$/.test(rawPPr) ? '<w:pPr></w:pPr>' : rawPPr
  const realStops = (format.tabStops ?? []).filter((ts) => !ts.rel && !ts.inherited)
  if (tags.has('w:tabs') && realStops.length === 0) {
    // mergePPrFormat only manages w:tabs while stops exist: clearing them all is done here
    out = out.replace(/<w:tabs>[\s\S]*?<\/w:tabs>|<w:tabs\/>/, '')
    tags.delete('w:tabs')
  }
  const viaMerge = new Set(
    [...tags].filter((t) => t === 'w:jc' || t === 'w:tabs' || t in PPR_FLAG_FIELDS),
  )
  if (viaMerge.size > 0) out = mergePPrFormat(out, format, undefined, viaMerge) || '<w:pPr></w:pPr>'
  const fresh = formatPPrChildren(format)
  for (const tag of ['w:spacing', 'w:ind']) {
    if (!tags.has(tag)) continue
    const rebuilt = new Set(
      CELL_MODELED_ATTR_GROUPS[tag].filter((g) => g.fields.some(differs)).flatMap((g) => g.attrs),
    )
    out = mergeCellElementAttrs(out, tag, fresh.find((c) => c.name === tag)?.xml, rebuilt)
  }
  return /^<w:pPr(?: [^>]*)?><\/w:pPr>$/.test(out) ? '' : out
}

export function mergePPrFormat(
  rawPPr: string,
  format: ParaFormat | undefined,
  original?: ParaFormat,
  /** restrict the managed groups (others keep their raw bytes whatever the format says) */
  onlyTags?: ReadonlySet<string>,
): string {
  const open = /^<w:pPr(?: [^>]*)?>/.exec(rawPPr)?.[0]
  const fresh = formatPPrChildren(format)
  // an indent edit on a paragraph laid out with character-unit indents: the
  // twips-only rebuild also cancels the components the edit drops
  // (`w:firstLineChars="0"`…), or Word — and this parser — would keep
  // resolving the character indent over the new value; components the new
  // format still carries are re-emitted (with their *Chars attribute) instead
  if (original?.charIndents && !sameIndent(original, format ?? {})) {
    const cancel = charIndentCancelAttrs(original.charIndents, format)
    if (cancel.length > 0) {
      const at = fresh.findIndex((c) => c.name === 'w:ind')
      const xml =
        at === -1
          ? `<w:ind ${cancel.join(' ')}/>`
          : fresh[at].xml.replace(/\/>$/, ` ${cancel.join(' ')}/>`)
      if (at === -1) fresh.push({ name: 'w:ind', xml })
      else fresh[at] = { name: 'w:ind', xml }
    }
  }
  if (!open) {
    // '<w:pPr/>' or unrecognized: rebuild from the format model alone, in schema order
    // (formatPPrChildren emits by concern, and CT_PPr order is not optional)
    const sorted = [...fresh].sort(
      (a, b) => PPR_CHILD_ORDER.indexOf(a.name) - PPR_CHILD_ORDER.indexOf(b.name),
    )
    return sorted.length > 0 ? `<w:pPr>${sorted.map((c) => c.xml).join('')}</w:pPr>` : ''
  }
  const inner = rawPPr.slice(open.length, rawPPr.length - '</w:pPr>'.length)
  // build set of managed tags for this format (base + conditional)
  const managedTags = new Set(FORMAT_MANAGED_TAGS)
  if (format?.tabStops !== undefined) managedTags.add('w:tabs')
  for (const tag of PPR_FLAG_TAGS) {
    const field = PPR_FLAG_FIELDS[tag]
    if (format?.[field] !== undefined || original?.[field] !== undefined) managedTags.add(tag)
  }
  if (format?.dropCap !== undefined || format?.frame !== undefined) managedTags.add('w:framePr')
  if (format?.textDirection !== undefined) managedTags.add('w:textDirection')
  if (format?.textAlignment !== undefined || original?.textAlignment !== undefined)
    managedTags.add('w:textAlignment')
  // only when the model carries a size: otherwise the paragraph-mark rPr stays unmanaged
  if (format?.emptyRunSizeHalfPoints !== undefined) managedTags.add('w:rPr')
  if (onlyTags) for (const tag of [...managedTags]) if (!onlyTags.has(tag)) managedTags.delete(tag)
  const rawChildren = splitXmlChildren(inner)
  const rawOf = (tag: string) => rawChildren.find((c) => c.name === tag)?.xml
  const rebuilt = new Set(
    [...managedTags].filter((tag) => !pprGroupUnchanged(tag, rawOf(tag), format ?? {}, original)),
  )
  const rank = (n: string) => PPR_CHILD_ORDER.indexOf(n)
  // the interleave below walks freshOut once with a monotonic index, so it has to arrive in
  // schema order; formatPPrChildren emits by concern and leaves framePr and tabs until last
  const freshOut = fresh
    .filter((c) => rebuilt.has(c.name))
    .sort((a, b) => rank(a.name) - rank(b.name))
  const kept = rawChildren.filter((c) => !rebuilt.has(c.name))
  const parts: string[] = []
  let fi = 0
  let prevRank = -1
  for (const child of kept) {
    const own = rank(child.name)
    const effective = own === -1 ? prevRank : Math.max(own, prevRank)
    while (fi < freshOut.length && rank(freshOut[fi].name) < effective)
      parts.push(freshOut[fi++].xml)
    parts.push(child.xml)
    prevRank = effective
  }
  while (fi < freshOut.length) parts.push(freshOut[fi++].xml)
  if (parts.length === 0) return ''
  return `${open}${parts.join('')}</w:pPr>`
}

/**
 * Strip <w:pPrChange>...</w:pPrChange> from a raw <w:pPr> slice.
 * Used when the editor accepts or rejects a paragraph-property revision,
 * so the saved file no longer contains the pPrChange record.
 */
/** hex color of the paragraph mark (pPr/rPr w:color), ignoring the pPrChange history */
export function paraMarkColor(rawPPr: string): string | undefined {
  const live = stripPPrChange(rawPPr)
  const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(live)?.[1]
  const color = rPr ? /<w:color\b[^>]*\bw:val="#?([^"]+)"/.exec(rPr)?.[1] : undefined
  return color && color !== 'auto' ? color.toUpperCase() : undefined
}

/**
 * Drop the paragraph mark's w:color (pPr/rPr). Word types new text at the end
 * of a paragraph in the mark's formatting, so a mark left red after every red
 * run was recolored would resurrect the red on the next keystroke.
 */
export function stripParaMarkColor(rawPPr: string): string {
  const changeStart = rawPPr.indexOf('<w:pPrChange')
  const live = changeStart === -1 ? rawPPr : rawPPr.slice(0, changeStart)
  const tail = changeStart === -1 ? '' : rawPPr.slice(changeStart)
  const stripped = live.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (_m, inner: string) => {
    // self-closing or paired form
    const rest = inner.replace(/<w:color\b[^>]*(?:\/>|>[\s\S]*?<\/w:color>)/, '')
    return rest.trim() ? `<w:rPr>${rest}</w:rPr>` : ''
  })
  return stripped + tail
}

/**
 * w:del / w:ins marker for a paragraph mark the editor tracked ('' when the
 * record is absent or malformed). Without an allocator the id is left empty
 * for fillRevisionIds to mint at save, like row/cell revisions.
 */
export function paraMarkRevisionXml(
  kind: 'del' | 'ins',
  record: string | null | undefined,
  mintId: () => string = () => '',
): string {
  if (!record) return ''
  try {
    const info = JSON.parse(record) as { author?: string; date?: string }
    const date = info.date ? ` w:date="${escapeXmlAttr(info.date)}"` : ''
    return `<w:${kind} w:id="${mintId()}" w:author="${escapeXmlAttr(info.author ?? '')}"${date}/>`
  } catch {
    return ''
  }
}

/** add a tracked paragraph-mark marker to raw pPr bytes (the mark's rPr opens with the revision markers) */
export function setParaMarkRevision(rawPPr: string, markerXml: string): string {
  if (!markerXml) return rawPPr
  const changeStart = rawPPr.indexOf('<w:pPrChange')
  let live = changeStart === -1 ? rawPPr : rawPPr.slice(0, changeStart)
  const tail = changeStart === -1 ? '' : rawPPr.slice(changeStart)
  if (!live || /^<w:pPr(?:\s[^>]*)?\/>$/.test(live)) live = '<w:pPr></w:pPr>'
  if (/<w:rPr\/>/.test(live)) live = live.replace(/<w:rPr\/>/, `<w:rPr>${markerXml}</w:rPr>`)
  else if (live.includes('<w:rPr>')) live = live.replace('<w:rPr>', `<w:rPr>${markerXml}`)
  else {
    const rPr = `<w:rPr>${markerXml}</w:rPr>`
    const sectAt = live.indexOf('<w:sectPr')
    if (sectAt !== -1) live = live.slice(0, sectAt) + rPr + live.slice(sectAt)
    else if (live.endsWith('</w:pPr>')) live = live.slice(0, -'</w:pPr>'.length) + rPr + '</w:pPr>'
    else live += rPr
  }
  return live + tail
}

/** drop the paragraph mark's tracked w:del / w:ins record (pPr/rPr) once the revision is accepted or rejected */
export function stripParaMarkRevision(rawPPr: string, kind: 'del' | 'ins'): string {
  const changeStart = rawPPr.indexOf('<w:pPrChange')
  const live = changeStart === -1 ? rawPPr : rawPPr.slice(0, changeStart)
  const tail = changeStart === -1 ? '' : rawPPr.slice(changeStart)
  const tags = kind === 'del' ? 'del|moveFrom' : 'ins|moveTo'
  const stripped = live.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (_m, inner: string) => {
    const rest = inner.replace(new RegExp(`<w:(?:${tags})\\b[^>]*/>`), '')
    return rest.trim() ? `<w:rPr>${rest}</w:rPr>` : ''
  })
  return stripped + tail
}

/**
 * Resolve the paragraph mark's tracked formatting change (pPr/rPr/rPrChange):
 * accept drops the record, reject puts the recorded old rPr back (the mark's
 * own ins/del/move markers stay).
 */
export function resolveParaMarkRPrChange(rawPPr: string, mode: 'accept' | 'reject'): string {
  const changeStart = rawPPr.indexOf('<w:pPrChange')
  const live = changeStart === -1 ? rawPPr : rawPPr.slice(0, changeStart)
  const tail = changeStart === -1 ? '' : rawPPr.slice(changeStart)
  const open = live.indexOf('<w:rPr>')
  const recordAt = live.indexOf('<w:rPrChange', open)
  if (open === -1 || recordAt === -1) return rawPPr
  // the record nests its own <w:rPr>, so the mark's rPr closes after the record does
  const selfClosing = /^<w:rPrChange\b[^>]*\/>/.exec(live.slice(recordAt))
  const closeTag = live.indexOf('</w:rPrChange>', recordAt)
  if (!selfClosing && closeTag === -1) return rawPPr
  const recordEnd = selfClosing
    ? recordAt + selfClosing[0].length
    : closeTag + '</w:rPrChange>'.length
  const close = live.indexOf('</w:rPr>', recordEnd)
  if (close === -1) return rawPPr
  const record = live.slice(recordAt, recordEnd)
  const rest = live.slice(open + '<w:rPr>'.length, recordAt) + live.slice(recordEnd, close)
  let next = rest
  if (mode === 'reject') {
    const markers = rest.match(/<w:(?:ins|del|moveFrom|moveTo)\b[^>]*\/>/g)?.join('') ?? ''
    const old = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(record)?.[1] ?? ''
    next = markers + old
  }
  const rPr = next.trim() ? `<w:rPr>${next}</w:rPr>` : ''
  return live.slice(0, open) + rPr + live.slice(close + '</w:rPr>'.length) + tail
}

/**
 * The paragraph's <w:pPr> element, depth-aware: a w:pPrChange child nests
 * another <w:pPr>, so a plain non-greedy match would truncate mid-record.
 */
export function rawPprOfParagraph(pXml: string): string {
  const open = /<w:pPr(?: [^>]*)?>/.exec(pXml)
  if (!open) return ''
  if (open[0].endsWith('/>')) return open[0]
  const start = open.index
  let depth = 0
  const re = /<w:pPr(?: [^>]*)?>|<\/w:pPr>/g
  re.lastIndex = start
  let m: RegExpExecArray | null
  while ((m = re.exec(pXml)) !== null) {
    if (m[0].startsWith('</')) depth -= 1
    else depth += 1
    if (depth === 0) return pXml.slice(start, re.lastIndex)
  }
  return ''
}

/**
 * The pPr the record reverts to: the old <w:pPr> element carried inside the
 * paragraph's own w:pPrChange, or null when the record has none. Rejecting a
 * paragraph-property revision restores it wholesale.
 */
export function pPrChangeOldXml(rawPPr: string): string | null {
  const start = rawPPr.indexOf('<w:pPrChange')
  if (start === -1) return null
  const m = /<w:pPr(?: [^>]*)?>[\s\S]*?<\/w:pPr>/.exec(rawPPr.slice(start))
  return m ? m[0] : null
}

export function stripPPrChange(rawPPr: string): string {
  // pPrChange nests a <w:pPr> inside it, so we need depth-aware removal.
  // Pattern: <w:pPrChange ...> ... </w:pPrChange>
  const start = rawPPr.indexOf('<w:pPrChange')
  if (start === -1) return rawPPr
  const end = rawPPr.indexOf('</w:pPrChange>', start)
  if (end === -1) return rawPPr
  return rawPPr.slice(0, start) + rawPPr.slice(end + '</w:pPrChange>'.length)
}

/** Build a pPrChange child from the editor's JSON snapshot of the previous properties. */
function revisionPPrChangeXml(changeJson: string): string | null {
  let change: Record<string, unknown>
  try {
    change = JSON.parse(changeJson) as Record<string, unknown>
  } catch {
    return null
  }
  const old = (change.old ?? {}) as Record<string, unknown>
  const oldFormat = (old.format && typeof old.format === 'object' ? old.format : old) as ParaFormat
  const children: PPrChild[] = []
  if (old.styleId) {
    children.push({
      name: 'w:pStyle',
      xml: `<w:pStyle w:val="${escapeXmlAttr(String(old.styleId))}"/>`,
    })
  }
  if (old.type === 'docListItem' && old.numId) {
    const ilvl = Math.min(Math.max(Number(old.ilvl) || 0, 0), 8)
    children.push({
      name: 'w:numPr',
      xml: `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${escapeXmlAttr(String(old.numId))}"/></w:numPr>`,
    })
  }
  children.push(...formatPPrChildren(oldFormat))
  children.sort((a, b) => PPR_CHILD_ORDER.indexOf(a.name) - PPR_CHILD_ORDER.indexOf(b.name))
  const attrs =
    ` w:id="${escapeXmlAttr(String(change.id ?? '0'))}"` +
    ` w:author="${escapeXmlAttr(String(change.author ?? ''))}"` +
    (change.date ? ` w:date="${escapeXmlAttr(String(change.date))}"` : '') +
    (change.dateUtc ? ` w16du:dateUtc="${escapeXmlAttr(String(change.dateUtc))}"` : '')
  return `<w:pPrChange${attrs}><w:pPr>${children.map((child) => child.xml).join('')}</w:pPr></w:pPrChange>`
}

/**
 * Insert or replace the live paragraph-property revision in a raw pPr slice.
 * The revision is last in CT_PPr schema order.
 */
export function setPPrChange(rawPPr: string, changeJson: string): string {
  const change = revisionPPrChangeXml(changeJson)
  if (!change) return rawPPr
  const clean = stripPPrChange(rawPPr)
  if (!clean || clean === '<w:pPr/>') return `<w:pPr>${change}</w:pPr>`
  const close = clean.lastIndexOf('</w:pPr>')
  if (close === -1) return `<w:pPr>${change}</w:pPr>`
  return `${clean.slice(0, close)}${change}${clean.slice(close)}`
}

/**
 * Generate an OOXML <w:p> fragment for an edited/new block.
 * Only references styles that already exist in the original document, so the
 * patched file never needs styles.xml modifications. When `rawPPr` is set the
 * paragraph properties pass through byte-identical instead of being rebuilt.
 */
export function generateParagraphXml(block: GeneratedBlock, ctx: GenerateContext): string {
  const crossStarts = (block.commentStarts ?? [])
    .map((id) => `<w:commentRangeStart w:id="${escapeXmlAttr(id)}"/>`)
    .join('')
  const crossEnds = (block.commentEnds ?? [])
    .map(
      (id) =>
        `<w:commentRangeEnd w:id="${escapeXmlAttr(id)}"/>` +
        `<w:r><w:commentReference w:id="${escapeXmlAttr(id)}"/></w:r>`,
    )
    .join('')
  const marks = bookmarksXml(ctx, [...(block.hiddenBookmarks ?? []), ...(block.bookmarks ?? [])])
  const content =
    marks.starts + crossStarts + generateRunsXml(block.runs, ctx) + crossEnds + marks.ends
  if (block.rawPPr !== undefined) return `<w:p>${block.rawPPr}${content}</w:p>`

  const children: PPrChild[] = []
  let styleId: string | undefined
  if (block.type === 'heading') {
    const level = Math.min(Math.max(block.level ?? 1, 1), 9)
    if (block.outlineOnly) {
      styleId = block.styleId
      children.push({ name: 'w:outlineLvl', xml: `<w:outlineLvl w:val="${level - 1}"/>` })
    } else styleId = block.styleId ?? ctx.headingStyleIds.get(level)
  } else if (block.type === 'listItem') {
    styleId = block.styleId ?? ctx.listParagraphStyleId
  } else {
    styleId = block.styleId
  }
  if (styleId)
    children.push({ name: 'w:pStyle', xml: `<w:pStyle w:val="${escapeXmlAttr(styleId)}"/>` })
  if (block.type === 'listItem' && block.list) {
    const ilvl = Math.min(Math.max(block.list.ilvl, 0), 8)
    children.push({
      name: 'w:numPr',
      xml: `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${escapeXmlAttr(block.list.numId)}"/></w:numPr>`,
    })
  }
  children.push(...formatPPrChildren(block.format))
  if (block.pPrChange) {
    const revision = revisionPPrChangeXml(block.pPrChange)
    if (revision) children.push({ name: 'w:pPrChange', xml: revision })
  }
  const markRevision =
    paraMarkRevisionXml('del', block.paraMarkDel, ctx.allocateRevisionId) +
    paraMarkRevisionXml('ins', block.paraMarkIns, ctx.allocateRevisionId)
  if (markRevision) {
    const rPr = children.find((c) => c.name === 'w:rPr')
    if (rPr) rPr.xml = rPr.xml.replace('<w:rPr>', `<w:rPr>${markRevision}`)
    else children.push({ name: 'w:rPr', xml: `<w:rPr>${markRevision}</w:rPr>` })
  }
  children.sort((a, b) => PPR_CHILD_ORDER.indexOf(a.name) - PPR_CHILD_ORDER.indexOf(b.name))
  const pPr = children.length > 0 ? `<w:pPr>${children.map((c) => c.xml).join('')}</w:pPr>` : ''
  return `<w:p>${pPr}${content}</w:p>`
}

/**
 * Fallback id source for a caller that has no document to coordinate with: a
 * counter, so ids stay unique among themselves. A hash of the name cannot do
 * that job: w:id is unique per part, the hash space overlaps the small ids
 * Word itself hands out, and two names collide by birthday at a few tens of
 * thousands of bookmarks — after which Word mis-pairs the bookmarks and a
 * cross-reference lands on the wrong target.
 */
let standaloneBookmarkSeq = 0

function bookmarksXml(
  ctx: Pick<GenerateContext, 'allocateBookmarkId'> | null,
  names: string[],
): { starts: string; ends: string } {
  // the bookmark spans the paragraph's text: a REF field to an empty bookmark shows nothing in Word
  let starts = ''
  let ends = ''
  for (const name of names) {
    const id = ctx?.allocateBookmarkId?.(name) ?? ++standaloneBookmarkSeq
    starts += `<w:bookmarkStart w:id="${id}" w:name="${escapeXmlAttr(name)}"/>`
    ends += `<w:bookmarkEnd w:id="${id}"/>`
  }
  return { starts, ends }
}

export interface TableGenOptions {
  /**
   * Style the first row as a header: light shading plus an empty bold run.
   * The bold run matters because patchTableCellTexts reuses the first run's
   * rPr when it fills in cell texts, so header texts come out bold.
   */
  headerRow?: boolean
}

/** shading fill used for generated table header rows (hex without '#') */
export const TABLE_HEADER_FILL = 'F2F2F2'

/** CT_TcPr child schema order (modeled subset + relative positions of common unmodeled items) */
const TCPR_ORDER = [
  'w:cnfStyle',
  'w:tcW',
  'w:gridSpan',
  'w:hMerge',
  'w:vMerge',
  'w:tcBorders',
  'w:shd',
  'w:noWrap',
  'w:tcMar',
  'w:textDirection',
  'w:tcFitText',
  'w:vAlign',
  'w:hideMark',
  'w:cellIns',
  'w:cellDel',
  'w:cellMerge',
  'w:tcPrChange',
]

/** Replace/insert/delete a tcPr child in schema order (xml=null deletes) */
function setTcPrChild(children: PPrChild[], name: string, xml: string | null): void {
  const idx = children.findIndex((c) => c.name === name)
  if (xml === null) {
    if (idx >= 0) children.splice(idx, 1)
    return
  }
  if (idx >= 0) {
    children[idx] = { name, xml }
    return
  }
  const orderIdx = TCPR_ORDER.indexOf(name)
  let insertAt = children.length
  for (let i = 0; i < children.length; i++) {
    const oi = TCPR_ORDER.indexOf(children[i].name)
    if (oi >= 0 && orderIdx >= 0 && oi > orderIdx) {
      insertAt = i
      break
    }
  }
  children.splice(insertAt, 0, { name, xml })
}

function cellBordersXml(borders: NonNullable<TableCell['borders']>): string {
  const side = (name: 'top' | 'left' | 'bottom' | 'right' | 'tl2br' | 'tr2bl') => {
    const b = borders[name]
    if (!b) return ''
    const sz =
      b.style === 'none' || b.style === 'nil' ? '' : ` w:sz="${b.szEighths ?? 4}" w:space="0"`
    const color =
      b.style === 'none' || b.style === 'nil'
        ? ''
        : ` w:color="${escapeXmlAttr(b.color ?? 'auto')}"`
    return `<w:${name} w:val="${escapeXmlAttr(b.style)}"${sz}${color}/>`
  }
  // CT_TcBorders is a sequence: top, left, bottom, right, tl2br, tr2bl
  return `<w:tcBorders>${side('top')}${side('left')}${side('bottom')}${side('right')}${side('tl2br')}${side('tr2bl')}</w:tcBorders>`
}

/** w:tcMar / w:tblCellMar with only the declared sides; null when nothing is declared */
function cellMarginsXml(tag: string, margins: CellMargins | undefined): string | null {
  const sides = (['top', 'left', 'bottom', 'right'] as const)
    .filter((side) => margins?.[side] != null)
    .map((side) => `<w:${side} w:w="${Math.max(0, Math.round(margins![side]!))}" w:type="dxa"/>`)
  return sides.length ? `<${tag}>${sides.join('')}</${tag}>` : null
}

/** colSpan arrives from parsed files/ops models: Infinity would emit
 *  w:val="Infinity" and blow up the grid build (Array.from({length: Infinity})
 *  throws). Clamp to 1..1000 at every emit site. */
function cellSpan(cell: TableCell): number {
  // colSpan arrives from parsed files/ops models as a number, but XML attr
  // plumbing can leave a numeric string behind: coerce before validating.
  const span = Number(cell.colSpan ?? 1)
  if (!Number.isFinite(span)) return 1
  return Math.min(Math.max(1, Math.floor(span)), 1000)
}

/** a property-change record the editor resolved: null = accepted, else its `resolved` verdict */
function resolvedProps(
  xml: string,
  changeName: string,
  change: PropChangeInfo | null | undefined,
): string {
  const mode = change === null ? 'accept' : change?.resolved
  return mode ? resolvePropChangeXml(xml, changeName, mode) : xml
}

function cellMergeXml(record: PropChangeInfo): string {
  const attr = (name: string, value: string | undefined) =>
    value === undefined ? '' : ` ${name}="${escapeXmlAttr(value)}"`
  return (
    `<w:cellMerge${attr('w:id', record.id ?? '0')}${attr('w:author', record.author)}` +
    `${attr('w:date', record.date)}${attr('w16du:dateUtc', record.dateUtc)}` +
    `${attr('w:vMerge', record.vMerge ?? 'cont')}${attr('w:vMergeOrig', record.vMergeOrig)}/>`
  )
}

function tableCellXml(
  cell: TableCell,
  width: number,
  tableWidth: number,
  verticalMerge?: 'restart' | 'continue',
  ctx?: Pick<GenerateContext, 'allocateBookmarkId'>,
): string {
  // Surgical rawTcPr patch: modeled children are replaced/inserted in schema order;
  // unmodeled ones (tcMar/textDirection/noWrap…) keep their original bytes
  const rawTcPr = cell.rawTcPr
    ? resolvedProps(
        resolvedProps(cell.rawTcPr, 'w:tcPrChange', cell.tcPrChange),
        'w:cellMerge',
        cell.cellMerge,
      )
    : ''
  const rawInner = rawTcPr.replace(/^<w:tcPr(?:\s[^>]*)?>/, '').replace(/<\/w:tcPr>$/, '')
  const children: PPrChild[] = rawTcPr.endsWith('</w:tcPr>') ? splitXmlChildren(rawInner) : []
  setTcPrChild(
    children,
    'w:tcW',
    cell.widthPct && tableWidth > 0
      ? `<w:tcW w:w="${Math.round((width / tableWidth) * 5000)}" w:type="pct"/>`
      : `<w:tcW w:w="${width}" w:type="dxa"/>`,
  )
  const span = cellSpan(cell)
  setTcPrChild(children, 'w:gridSpan', span > 1 ? `<w:gridSpan w:val="${span}"/>` : null)
  const merge = verticalMerge ?? cell.vMerge
  setTcPrChild(
    children,
    'w:vMerge',
    merge === 'restart'
      ? '<w:vMerge w:val="restart"/>'
      : merge === 'continue'
        ? '<w:vMerge/>'
        : null,
  )
  // a merged-away cell rebuilt from its restart cell's bytes lost the merge record it carried
  const cellMerge = cell.cellMerge
  if (cellMerge && !cellMerge.resolved && merge === 'continue' && !rawTcPr.includes('<w:cellMerge'))
    setTcPrChild(children, 'w:cellMerge', cellMergeXml(cellMerge))
  // With borders unmodeled (undefined), leave the raw tcBorders alone; clearing borders
  // uses explicit none
  if (cell.borders) setTcPrChild(children, 'w:tcBorders', cellBordersXml(cell.borders))
  setTcPrChild(
    children,
    'w:shd',
    cell.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${escapeXmlAttr(cell.fill)}"/>` : null,
  )
  setTcPrChild(
    children,
    'w:vAlign',
    cell.vAlign && cell.vAlign !== 'top' ? `<w:vAlign w:val="${cell.vAlign}"/>` : null,
  )
  if (cell.tcPrEdited) {
    setTcPrChild(
      children,
      'w:textDirection',
      cell.textDirection && cell.textDirection !== 'lrTb'
        ? `<w:textDirection w:val="${cell.textDirection}"/>`
        : null,
    )
    setTcPrChild(children, 'w:tcMar', cellMarginsXml('w:tcMar', cell.cellMarTwips))
    setTcPrChild(children, 'w:noWrap', cell.noWrap ? '<w:noWrap/>' : null)
    setTcPrChild(children, 'w:tcFitText', cell.fitText ? '<w:tcFitText/>' : null)
  }
  const tcPr = children.map((c) => c.xml)
  const paragraphs = cell.richParas?.length
    ? cell.richParas
    : (cell.paras.length > 0 ? cell.paras : ['']).map((text) => ({
        align: cell.align,
        runs: text === '' ? [] : [{ text, bold: cell.bold, color: cell.color }],
      }))
  const paraXmls = paragraphs.map((paragraph) => {
    const list = 'list' in paragraph ? paragraph.list : undefined
    const numPr = list
      ? `<w:numPr><w:ilvl w:val="${list.ilvl}"/><w:numId w:val="${escapeXmlAttr(list.numId)}"/></w:numPr>`
      : ''
    // no cell.align fallback here: it is a display aggregate of the paragraphs' own
    // jc values, and writing it back would stamp w:jc into paragraphs that never had one
    // (rich paragraphs carry their own align; the plain-paras fallback above sets it).
    // The parsed format, not just w:jc: hand-building the pPr here dropped w:bidi and
    // discarded every other paragraph property the model carries.
    const styleId = 'styleId' in paragraph ? paragraph.styleId : undefined
    const pStyle = styleId ? `<w:pStyle w:val="${escapeXmlAttr(styleId)}"/>` : ''
    const pPr = mergePPrFormat(`<w:pPr>${pStyle}${numPr}</w:pPr>`, paragraph)
    // the cell's bookmark anchors ride the rebuilt paragraph: a caller writing
    // into an existing document supplies the document's allocator, whose ids
    // never collide with the bookmarks the untouched blocks keep (REF fields
    // match by name, but bookmarkStart / bookmarkEnd pair by id)
    const marks = bookmarksXml(ctx ?? null, [
      ...(('hiddenBookmarks' in paragraph && paragraph.hiddenBookmarks) || []),
      ...(('bookmarks' in paragraph && paragraph.bookmarks) || []),
    ])
    return `<w:p>${pPr}${marks.starts}${runsXml(paragraph.runs, null, undefined, ctx ?? null)}${marks.ends}</w:p>`
  })
  // nested tables are regenerated from the model at their paragraph anchors (reverse
  // insertion keeps anchor indexes valid); OOXML requires tc to end with w:p
  const nested = cell.nestedTables ?? []
  const items = paraXmls.map((xml) => ({ tbl: false, xml }))
  for (let i = nested.length - 1; i >= 0; i--) {
    const at = Math.min(cell.nestedTableAnchors?.[i] ?? paraXmls.length, paraXmls.length)
    items.splice(at, 0, { tbl: true, xml: generateTableModelXml(nested[i], undefined, ctx) })
  }
  const tail = items.length === 0 || items[items.length - 1].tbl ? '<w:p/>' : ''
  const content = items.map((item) => item.xml).join('')
  return `<w:tc><w:tcPr>${tcPr.join('')}</w:tcPr>${content}${tail}</w:tc>`
}

/**
 * Generate a complete table from the editable display model. This is used only
 * after structural edits; untouched tables and text-only edits keep their
 * original XML through the byte-preserving patch paths.
 */
const TBL_PR_ORDER = [
  'w:tblStyle',
  'w:tblpPr',
  'w:tblOverlap',
  'w:bidiVisual',
  'w:tblStyleRowBandSize',
  'w:tblStyleColBandSize',
  'w:tblW',
  'w:jc',
  'w:tblCellSpacing',
  'w:tblInd',
  'w:tblBorders',
  'w:shd',
  'w:tblLayout',
  'w:tblCellMar',
  'w:tblLook',
  'w:tblCaption',
  'w:tblDescription',
] as const

/** Replace/remove/insert one tblPr child while keeping the OOXML schema order. */
function setTblPrChild(tblPr: string, name: string, xml: string | null): string {
  const selfClosing = /^<w:tblPr(?:\s[^>]*)?\/>$/.test(tblPr)
  const open = selfClosing ? tblPr.replace(/\/>$/, '>') : /^<w:tblPr(?:\s[^>]*)?>/.exec(tblPr)?.[0]
  if (!open) return tblPr
  const inner = selfClosing ? '' : tblPr.slice(open.length, tblPr.lastIndexOf('</w:tblPr>'))
  const children = splitXmlChildren(inner).filter((child) => child.name !== name)
  if (xml) {
    const order = TBL_PR_ORDER.indexOf(name as (typeof TBL_PR_ORDER)[number])
    const at = children.findIndex((child) => {
      const childOrder = TBL_PR_ORDER.indexOf(child.name as (typeof TBL_PR_ORDER)[number])
      return order >= 0 && childOrder >= 0 && childOrder > order
    })
    children.splice(at < 0 ? children.length : at, 0, { name, xml })
  }
  return `${open}${children.map((child) => child.xml).join('')}</w:tblPr>`
}

function tableLookXml(look: NonNullable<TableModel['tableLook']>): string {
  const val =
    (look.firstRow ? 0x20 : 0) |
    (look.lastRow ? 0x40 : 0) |
    (look.firstColumn ? 0x80 : 0) |
    (look.lastColumn ? 0x100 : 0) |
    (look.bandedRows ? 0 : 0x200) |
    (look.bandedColumns ? 0 : 0x400)
  return (
    `<w:tblLook w:val="${val.toString(16).toUpperCase().padStart(4, '0')}"` +
    ` w:firstRow="${look.firstRow ? 1 : 0}" w:lastRow="${look.lastRow ? 1 : 0}"` +
    ` w:firstColumn="${look.firstColumn ? 1 : 0}" w:lastColumn="${look.lastColumn ? 1 : 0}"` +
    ` w:noHBand="${look.bandedRows ? 0 : 1}" w:noVBand="${look.bandedColumns ? 0 : 1}"/>`
  )
}

/** the table's own <w:tblPr> (depth-aware: a w:tblPrChange nests another tblPr inside it) */
function originalTblPrOf(tableXml: string): string | undefined {
  const open = /<w:tbl[\s>]/.exec(tableXml)
  if (!open) return undefined
  const innerStart = tableXml.indexOf('>', open.index) + 1
  const innerEnd = tableXml.lastIndexOf('</w:tbl>')
  if (innerStart <= 0 || innerEnd < innerStart) return undefined
  return splitXmlChildren(tableXml.slice(innerStart, innerEnd)).find((c) => c.name === 'w:tblPr')
    ?.xml
}

export function generateTableModelXml(
  model: TableModel,
  originalTableXml?: string,
  ctx?: Pick<GenerateContext, 'allocateBookmarkId'>,
): string {
  // the widest row decides the column count; folding it in a loop keeps the
  // argument count off the call (a spread here blew the stack past ~125k rows)
  let columnCount = Math.max(1, model.colWidthsPct?.length ?? 0)
  for (const row of model.rows) {
    const span = row.reduce((sum, cell) => sum + cellSpan(cell), 0)
    if (span > columnCount) columnCount = span
  }
  const percentages =
    model.colWidthsPct?.length === columnCount
      ? model.colWidthsPct
      : Array.from({ length: columnCount }, () => 100 / columnCount)
  const totalPct = percentages.reduce((sum, value) => sum + value, 0) || 100
  const widths =
    model.colWidthsTwips?.length === columnCount
      ? model.colWidthsTwips.map((value) => Math.max(1, Math.round(value)))
      : percentages.map((value) => Math.max(1, Math.round((value / totalPct) * 9360)))
  const totalWidth = widths.reduce((sum, value) => sum + value, 0)
  const border = (name: string) => `<w:${name} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
  const sides = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const
  // model borders (e.g. 'none' for borderless source tables) win over the
  // self-contained single-line default
  const borders =
    '<w:tblBorders>' +
    sides
      .map((name) => {
        const b = model.borders?.[name]
        if (!b) return model.borders ? '' : border(name)
        return (
          `<w:${name} w:val="${escapeXmlAttr(b.style)}"` +
          (b.szEighths != null ? ` w:sz="${Math.round(b.szEighths)}"` : '') +
          ` w:space="0" w:color="${escapeXmlAttr(b.color ?? 'auto')}"/>`
        )
      })
      .join('') +
    '</w:tblBorders>'
  const marSides = ['top', 'left', 'bottom', 'right'] as const
  const cellMar = model.cellMarTwips
    ? '<w:tblCellMar>' +
      marSides
        .map((name) => {
          const w = model.cellMarTwips?.[name]
          return w != null ? `<w:${name} w:w="${Math.round(w)}" w:type="dxa"/>` : ''
        })
        .join('') +
      '</w:tblCellMar>'
    : ''
  const gridChange = model.tblGridChange
  const gridRecord =
    gridChange && !gridChange.resolved
      ? `<w:tblGridChange w:id="${escapeXmlAttr(gridChange.id ?? '0')}"><w:tblGrid>${gridChange.oldXml ?? ''}</w:tblGrid></w:tblGridChange>`
      : ''
  const grid = `<w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}${gridRecord}</w:tblGrid>`
  const rows = model.rows
    .map((row, ri) => {
      let gridColumn = 0
      const cells: string[] = []
      // gridGap placeholders are the display stand-ins for w:gridBefore/w:gridAfter:
      // they advance the grid but are never written as w:tc
      const edges = { before: 0, wBefore: 0, after: 0, wAfter: 0 }
      for (const cell of row) {
        const span = cellSpan(cell)
        const width = widths
          .slice(gridColumn, gridColumn + span)
          .reduce((sum, value) => sum + value, 0)
        gridColumn += span
        if (cell.gridGap) {
          if (cells.length === 0) {
            edges.before += span
            edges.wBefore += width
          } else {
            edges.after += span
            edges.wAfter += width
          }
          continue
        }
        cells.push(tableCellXml(cell, width, totalWidth, undefined, ctx))
      }
      // trPr: original bytes passed through (tblHeader/cantSplit etc.); row height is
      // replaced/inserted/removed per the model
      let trPr = resolvedProps(model.rawTrPrs?.[ri] ?? '', 'w:trPrChange', model.rowPrChanges?.[ri])
      if (trPr && !trPr.endsWith('</w:trPr>')) trPr = ''
      let tblPrEx = resolvedProps(
        model.rawTblPrExs?.[ri] ?? '',
        'w:tblPrExChange',
        model.rowTblPrExChanges?.[ri],
      )
      if (tblPrEx && !tblPrEx.endsWith('</w:tblPrEx>')) tblPrEx = ''
      // the placeholders are the truth for the grid edges: stale raw edge tags go, the
      // computed ones lead the trPr (CT_TrPr order: cnfStyle, divId, gridBefore, gridAfter, wBefore, wAfter, …)
      trPr = trPr.replace(/<w:(?:gridBefore|gridAfter|wBefore|wAfter)\b[^>]*\/>/g, '')
      if (edges.before > 0 || edges.after > 0) {
        const edgeXml =
          (edges.before > 0 ? `<w:gridBefore w:val="${edges.before}"/>` : '') +
          (edges.after > 0 ? `<w:gridAfter w:val="${edges.after}"/>` : '') +
          (edges.before > 0 ? `<w:wBefore w:w="${edges.wBefore}" w:type="dxa"/>` : '') +
          (edges.after > 0 ? `<w:wAfter w:w="${edges.wAfter}" w:type="dxa"/>` : '')
        if (!trPr) trPr = `<w:trPr>${edgeXml}</w:trPr>`
        else {
          const head =
            /^<w:trPr(?:\s[^>]*)?>(?:<w:cnfStyle\b[^>]*\/>)?(?:<w:divId\b[^>]*\/>)?/.exec(trPr)!
          trPr = trPr.slice(0, head[0].length) + edgeXml + trPr.slice(head[0].length)
        }
      } else if (/^<w:trPr(?:\s[^>]*)?>\s*<\/w:trPr>$/.test(trPr)) {
        trPr = ''
      }
      const repeatHeader = model.repeatHeaderRows?.[ri]
      if (repeatHeader !== null && repeatHeader !== undefined) {
        trPr = trPr.replace(/<w:tblHeader(?:\s[^>]*)?\/>/g, '')
        if (repeatHeader) {
          const tag = '<w:tblHeader/>'
          trPr = trPr ? trPr.replace('</w:trPr>', `${tag}</w:trPr>`) : `<w:trPr>${tag}</w:trPr>`
        } else if (/^<w:trPr(?:\s[^>]*)?>\s*<\/w:trPr>$/.test(trPr)) {
          trPr = ''
        }
      }
      const cantSplit = model.rowCantSplit?.[ri]
      if (cantSplit !== null && cantSplit !== undefined) {
        trPr = trPr.replace(/<w:cantSplit(?:\s[^>]*)?\/>/g, '')
        if (cantSplit) {
          // CT_TrPr: cantSplit precedes trHeight and tblHeader
          const tag = '<w:cantSplit/>'
          const at = trPr.search(/<w:trHeight\b|<w:tblHeader\b|<\/w:trPr>/)
          trPr = trPr ? trPr.slice(0, at) + tag + trPr.slice(at) : `<w:trPr>${tag}</w:trPr>`
        } else if (/^<w:trPr(?:\s[^>]*)?>\s*<\/w:trPr>$/.test(trPr)) {
          trPr = ''
        }
      }
      const h = model.rowHeightsTwips?.[ri]
      if (h != null && h > 0) {
        const rule = model.rowHeightRules?.[ri] === 'exact' ? 'exact' : 'atLeast'
        const tag = `<w:trHeight w:val="${Math.round(h)}" w:hRule="${rule}"/>`
        if (!trPr) trPr = `<w:trPr>${tag}</w:trPr>`
        else if (/<w:trHeight[^>]*\/>/.test(trPr)) trPr = trPr.replace(/<w:trHeight[^>]*\/>/, tag)
        else trPr = trPr.replace('</w:trPr>', `${tag}</w:trPr>`)
      } else if (trPr) {
        trPr = trPr.replace(/<w:trHeight[^>]*\/>/, '')
        if (/^<w:trPr(?:\s[^>]*)?>\s*<\/w:trPr>$/.test(trPr)) trPr = ''
      }
      return `<w:tr>${tblPrEx}${trPr}${cells.join('')}</w:tr>`
    })
    .join('')
  const originalTblPr = originalTableXml ? originalTblPrOf(originalTableXml) : undefined
  const tblInd =
    model.indentTwips && Math.round(model.indentTwips) !== 0
      ? `<w:tblInd w:w="${Math.round(model.indentTwips)}" w:type="dxa"/>`
      : ''
  let tblPr = resolvedProps(
    originalTblPr ??
      `<w:tblPr><w:tblW w:w="${totalWidth}" w:type="dxa"/>${tblInd}${borders}</w:tblPr>`,
    'w:tblPrChange',
    model.tblPrChange,
  )
  // Table style reference: '' = remove, non-empty = replace/insert (tblStyle is the
  // first tblPr child)
  if (model.tblStyleId !== undefined) {
    tblPr = setTblPrChild(
      tblPr,
      'w:tblStyle',
      model.tblStyleId === '' ? null : `<w:tblStyle w:val="${escapeXmlAttr(model.tblStyleId)}"/>`,
    )
  }
  // Floating positioning + wrapping. null explicitly returns the table to the
  // text flow; undefined leaves an imported tblpPr byte-for-byte intact.
  const requestedFloatSide =
    model.floatSide === undefined && model.floatPos
      ? model.floatPos.xTwips > 4680
        ? 'right'
        : 'left'
      : model.floatSide
  if (requestedFloatSide !== undefined || (!originalTblPr && model.floatPos)) {
    if (!requestedFloatSide) {
      tblPr = setTblPrChild(tblPr, 'w:tblpPr', null)
      tblPr = setTblPrChild(tblPr, 'w:tblOverlap', null)
    } else {
      const pos = model.floatPos ?? {
        xTwips: requestedFloatSide === 'right' ? Math.max(0, 9360 - totalWidth) : 0,
        yTwips: 0,
      }
      const distance = pos.distanceTwips ?? {}
      const attr = (name: string, value: number | undefined) =>
        value == null ? '' : ` w:${name}="${Math.max(0, Math.round(value))}"`
      const tag =
        `<w:tblpPr${attr('leftFromText', distance.left)}${attr('rightFromText', distance.right)}` +
        `${attr('topFromText', distance.top)}${attr('bottomFromText', distance.bottom)}` +
        ` w:vertAnchor="${pos.vertAnchor ?? 'page'}"` +
        ` w:horzAnchor="${pos.horzAnchor ?? 'page'}"` +
        ` w:tblpX="${Math.round(pos.xTwips)}" w:tblpY="${Math.round(pos.yTwips)}"/>`
      tblPr = setTblPrChild(tblPr, 'w:tblpPr', tag)
      tblPr = setTblPrChild(tblPr, 'w:tblOverlap', '<w:tblOverlap w:val="overlap"/>')
    }
  }
  const effectiveAutoFit =
    model.autoFit ?? (!originalTblPr ? (model.autoLayout ? 'contents' : 'fixed') : null)
  if (effectiveAutoFit) {
    tblPr = setTblPrChild(
      tblPr,
      'w:tblW',
      effectiveAutoFit === 'window'
        ? '<w:tblW w:w="5000" w:type="pct"/>'
        : effectiveAutoFit === 'contents'
          ? '<w:tblW w:w="0" w:type="auto"/>'
          : `<w:tblW w:w="${totalWidth}" w:type="dxa"/>`,
    )
    tblPr = setTblPrChild(
      tblPr,
      'w:tblLayout',
      `<w:tblLayout w:type="${effectiveAutoFit === 'fixed' ? 'fixed' : 'autofit'}"/>`,
    )
  }
  if (model.cellMarTwips !== undefined) {
    tblPr = setTblPrChild(tblPr, 'w:tblCellMar', cellMar || null)
  }
  if (model.tableLook) {
    tblPr = setTblPrChild(tblPr, 'w:tblLook', tableLookXml(model.tableLook))
  } else if (model.tblStyleId && !/<w:tblLook[\s/>]/.test(tblPr)) {
    tblPr = setTblPrChild(
      tblPr,
      'w:tblLook',
      tableLookXml({
        firstRow: true,
        lastRow: false,
        firstColumn: true,
        lastColumn: false,
        bandedRows: true,
        bandedColumns: false,
      }),
    )
  }
  // Table alignment (w:jc): explicit align replaces/removes the original;
  // undefined keeps whatever the original tblPr carried
  if (model.align) {
    tblPr = setTblPrChild(
      tblPr,
      'w:jc',
      model.align === 'left' ? null : `<w:jc w:val="${model.align}"/>`,
    )
  }
  if (model.caption !== undefined) {
    tblPr = setTblPrChild(
      tblPr,
      'w:tblCaption',
      model.caption ? `<w:tblCaption w:val="${escapeXmlAttr(model.caption)}"/>` : null,
    )
  }
  if (model.description !== undefined) {
    tblPr = setTblPrChild(
      tblPr,
      'w:tblDescription',
      model.description ? `<w:tblDescription w:val="${escapeXmlAttr(model.description)}"/>` : null,
    )
  }
  return `<w:tbl>${tblPr}${grid}${rows}</w:tbl>`
}

/** rId / docPr placeholder of a picture inserted by the editor; saveDocx resolves
 *  it against SaveOptions.newPictures once the bytes land as a media part */
export const NEW_PICTURE_TOKEN_PREFIX = 'aidocs-new-pic-'
/** after a token, in the svgBlip's r:embed: the picture's SVG part rather than its bitmap */
export const NEW_PICTURE_SVG_SUFFIX = '-svg'

const SVG_BLIP_EXT_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}'
const SVG_BLIP_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'

export interface InlinePictureXmlOptions {
  rId: string
  /** relationship of the picture's SVG part, when the bitmap at rId is its fallback */
  svgRId?: string
  /** wp:docPr / pic:cNvPr id (a placeholder token until the save mints a real one) */
  docPrId: string | number
  widthPx: number
  heightPx: number
  altText?: string
  rotDeg?: number
  flipH?: boolean
  flipV?: boolean
}

/** The drawing re-sized to widthPx x heightPx: wp:extent and the picture's own
 *  a:ext follow (effectExtent is left alone). Unchanged when the extent already
 *  reads these sizes, so an untouched picture saves byte-identical. */
export function resizeInlinePictureXml(xml: string, widthPx: number, heightPx: number): string {
  if (!/<wp:inline[\s>]/.test(xml)) return xml
  return resizeDrawingExtents(xml, widthPx, heightPx)
}

/** resizeInlinePictureXml for a floating (wp:anchor) run picture; only a lone
 *  picture: a group's extent does not follow its first picture */
export function resizeAnchoredPictureXml(xml: string, widthPx: number, heightPx: number): string {
  if (!/<wp:anchor[\s>]/.test(xml) || /<wpg:/.test(xml)) return xml
  if ((xml.match(/<pic:pic[\s>]/g) ?? []).length !== 1) return xml
  return resizeDrawingExtents(xml, widthPx, heightPx)
}

function resizeDrawingExtents(xml: string, widthPx: number, heightPx: number): string {
  if (!(widthPx > 0) || !(heightPx > 0)) return xml
  const cx = Math.max(1, Math.round(widthPx * EMU_PER_PX))
  const cy = Math.max(1, Math.round(heightPx * EMU_PER_PX))
  const extent = /<wp:extent\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(xml)
  if (!extent) return xml
  const samePx = (emu: string, px: number) =>
    Math.round((Number(emu) / EMU_PER_PX) * 100) / 100 === px
  if (samePx(extent[1], widthPx) && samePx(extent[2], heightPx)) return xml
  const ext = (tag: string) =>
    tag.replace(/\bcx="\d+"/, `cx="${cx}"`).replace(/\bcy="\d+"/, `cy="${cy}"`)
  let out = xml.replace(/<wp:extent\b[^>]*>/, ext)
  const pic = /<pic:spPr[\s>][\s\S]*?<\/pic:spPr>/.exec(out)
  if (pic) out = out.replace(pic[0], pic[0].replace(/<a:ext\b[^>]*>/, ext))
  return out
}

/** `<w:drawing>` of an inline picture (wp:inline + pic:pic), the run and paragraph excluded */
export function inlinePictureDrawingXml(o: InlinePictureXmlOptions): string {
  const cx = Math.max(1, Math.round(o.widthPx * EMU_PER_PX))
  const cy = Math.max(1, Math.round(o.heightPx * EMU_PER_PX))
  // Word lays the drawing out against the unrotated wp:extent plus
  // wp:effectExtent: a rotated non-square picture needs the bounding-box
  // overflow recorded there (same math as patchImageParagraphXml)
  const rot = o.rotDeg ? ((Math.round(o.rotDeg) % 360) + 360) % 360 : 0
  const rad = (rot * Math.PI) / 180
  const bw = Math.abs(cx * Math.cos(rad)) + Math.abs(cy * Math.sin(rad))
  const bh = Math.abs(cx * Math.sin(rad)) + Math.abs(cy * Math.cos(rad))
  const eeX = Math.max(0, Math.round((bw - cx) / 2))
  const eeY = Math.max(0, Math.round((bh - cy) / 2))
  const id = String(o.docPrId)
  const name = `Picture ${id}`
  const blip = o.svgRId
    ? `<a:blip r:embed="${o.rId}"><a:extLst><a:ext uri="${SVG_BLIP_EXT_URI}">` +
      `<asvg:svgBlip xmlns:asvg="${SVG_BLIP_NS}" r:embed="${o.svgRId}"/></a:ext></a:extLst></a:blip>`
    : `<a:blip r:embed="${o.rId}"/>`
  return (
    '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="${eeX}" t="${eeY}" r="${eeX}" b="${eeY}"/>` +
    `<wp:docPr id="${id}" name="${name}"${o.altText ? ` descr="${escapeXmlAttr(o.altText)}"` : ''}/>` +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm${rot ? ` rot="${rot * 60000}"` : ''}${o.flipH ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
  )
}

/**
 * Generate a self-contained w:tbl fragment (inline borders, no style reference),
 * so the patched file never needs styles.xml modifications.
 */
export function generateTableXml(rows: number, cols: number, opts: TableGenOptions = {}): string {
  const safeRows = Math.min(Math.max(rows, 1), 50)
  const safeCols = Math.min(Math.max(cols, 1), 20)
  const totalWidth = 9360 // 6.5in content width on US Letter
  const colWidth = Math.floor(totalWidth / safeCols)
  const border = (name: string) => `<w:${name} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
  const borders =
    '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('') +
    '</w:tblBorders>'
  const grid = `<w:tblGrid>${`<w:gridCol w:w="${colWidth}"/>`.repeat(safeCols)}</w:tblGrid>`
  const cell = (header: boolean) =>
    `<w:tc><w:tcPr><w:tcW w:w="${colWidth}" w:type="dxa"/>` +
    (header ? `<w:shd w:val="clear" w:color="auto" w:fill="${TABLE_HEADER_FILL}"/>` : '') +
    '</w:tcPr>' +
    (header ? '<w:p><w:r><w:rPr><w:b/></w:rPr></w:r></w:p>' : '<w:p/>') +
    '</w:tc>'
  const row = (header: boolean) => `<w:tr>${cell(header).repeat(safeCols)}</w:tr>`
  const body = opts.headerRow
    ? row(true) + row(false).repeat(Math.max(safeRows - 1, 0))
    : row(false).repeat(safeRows)
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${totalWidth}" w:type="dxa"/>${borders}` +
    `<w:tblLayout w:type="fixed"/></w:tblPr>${grid}${body}</w:tbl>` +
    // Word requires a paragraph after a table before another table/sectPr;
    // an empty trailing paragraph keeps the body valid in every position.
    '<w:p/>'
  )
}

export interface TocEntry {
  /** heading level 1-9 */
  level: number
  text: string
  /** page number computed by real pagination, saved as the field's result */
  pageNo?: number
  /** the entry's level shows no page number (TOC \n) */
  noPage?: boolean
}

/**
 * Generate a real TOC field as one w:p fragment per line. The entries and
 * their pages are the field's result, as Word saves one; only a table whose
 * pages are not known yet is marked dirty for Word to build.
 */
export function generateTocFieldXml(entries: TocEntry[], table: FieldTableLayout = {}): string[] {
  if (entries.length === 0) return []
  let deepest = 1
  for (const entry of entries) {
    if (entry.level > deepest) deepest = entry.level
  }
  const maxLevel = Math.min(deepest, 9)
  return tocFieldLines(
    entries,
    table.instr ?? ` TOC \\o "1-${maxLevel}" \\h \\z \\u `,
    (entry) => `TOC${Math.min(Math.max(entry.level, 1), 9)}`,
    table,
  )
}

/**
 * A table of figures: the same field as a table of contents, built from the
 * captions of one label (TOC \c) instead of the headings.
 */
export function generateTableOfFiguresFieldXml(
  entries: Array<{ text: string; pageNo?: number; noPage?: boolean }>,
  label: string,
  table: FieldTableLayout = {},
): string[] {
  // the label is quoted inside the instruction: a quote of its own would end it early
  const quoted = label.replace(/"/g, '')
  return tocFieldLines(
    entries,
    table.instr ?? ` TOC \\h \\z \\c "${quoted}" `,
    () => 'TableofFigures',
    table,
  )
}

/**
 * The run's phonetic guide with the run's formatting on its base text. The
 * formatting of a guide lives on the runs inside it, not on the run around
 * it; a guide that already says how its base looks is written as it is.
 */
function rubyWithBaseFormat(run: Run, insideLink: boolean): string {
  const { xml } = run.ruby!
  const base = /<w:rubyBase>\s*<w:r>(?!\s*<w:rPr[\s>])/.exec(xml)
  if (!base) return xml
  const rPr = runRPrXml({ ...run, ruby: undefined }, insideLink)
  if (!rPr) return xml
  const at = base.index + base[0].length
  return xml.slice(0, at) + rPr + xml.slice(at)
}

/**
 * The same guide with another reading: everything else it says (alignment,
 * sizes, the formatting of its runs) stays. Null when the fragment has no
 * reading text to replace.
 */
export function setRubyReadingXml(xml: string, rt: string): string | null {
  return setRubyPartText(xml, 'w:rt', rt)
}

/** where a reading sits over its base text (w:rubyAlign), the five Word's dialog offers */
export const RUBY_ALIGNS = [
  'center',
  'distributeLetter',
  'distributeSpace',
  'left',
  'right',
] as const
export type RubyAlign = (typeof RUBY_ALIGNS)[number]

/** the w:rubyAlign of a guide; null when it has none or one the dialog does not offer */
export function rubyAlignOf(xml: string): RubyAlign | null {
  const val = /<w:rubyAlign\b[^>]*\bw:val="([^"]*)"/.exec(xml)?.[1]
  return (RUBY_ALIGNS as readonly string[]).includes(val ?? '') ? (val as RubyAlign) : null
}

/** the guide with its alignment changed; null when the fragment has no w:rubyPr to hold one */
export function setRubyAlignXml(xml: string, align: RubyAlign): string | null {
  const tag = `<w:rubyAlign w:val="${align}"/>`
  if (/<w:rubyAlign\b[^>]*\/>/.test(xml)) return xml.replace(/<w:rubyAlign\b[^>]*\/>/, tag)
  // w:rubyAlign is the first child of w:rubyPr
  return /<w:rubyPr>/.test(xml) ? xml.replace('<w:rubyPr>', `<w:rubyPr>${tag}`) : null
}

/** how a guide's reading is set: its font, its size and how far it is raised, in half-points */
export interface RubyReadingStyle {
  font?: string
  sizeHalfPoints?: number
  raiseHalfPoints?: number
  baseSizeHalfPoints?: number
}

const rubyPrInt = (xml: string, tag: string): number | undefined => {
  const val = new RegExp(`<w:${tag}\\b[^>]*\\bw:val="(\\d+)"`).exec(xml)?.[1]
  return val === undefined ? undefined : parseInt(val, 10)
}

export function rubyReadingStyleOf(xml: string): RubyReadingStyle {
  const rt = /<w:rt>([\s\S]*?)<\/w:rt>/.exec(xml)?.[1] ?? ''
  const fonts = /<w:rFonts\b[^>]*>/.exec(rt)?.[0] ?? ''
  const font = /\bw:(?:ascii|eastAsia|hAnsi)="([^"]*)"/.exec(fonts)?.[1]
  return {
    font: font ? decodeXmlText(font) : undefined,
    sizeHalfPoints: rubyPrInt(xml, 'hps'),
    raiseHalfPoints: rubyPrInt(xml, 'hpsRaise'),
    baseSizeHalfPoints: rubyPrInt(xml, 'hpsBaseText'),
  }
}

/** the guide with its reading set another way; null when the fragment has no w:rubyPr or reading run */
export function setRubyReadingStyleXml(
  xml: string,
  style: Pick<RubyReadingStyle, 'font' | 'sizeHalfPoints' | 'raiseHalfPoints'>,
): string | null {
  if (!/<w:rubyPr>/.test(xml)) return null
  const found = /<w:rt>([\s\S]*?)<\/w:rt>/.exec(xml)
  if (!found) return null
  // w:rubyPr holds rubyAlign, hps, hpsRaise, hpsBaseText, lid in that order
  const setPr = (src: string, tag: string, val: number, before: string): string => {
    const el = `<w:${tag} w:val="${Math.max(0, Math.round(val))}"/>`
    const has = new RegExp(`<w:${tag}\\b[^>]*/>`)
    if (has.test(src)) return src.replace(has, el)
    const next = new RegExp(`<w:(?:${before})\\b|</w:rubyPr>`)
    return src.replace(next, (m) => el + m)
  }
  let inner = found[1]
  const { font, sizeHalfPoints: size, raiseHalfPoints: raise } = style
  if (font !== undefined || size !== undefined) {
    inner = inner.replace(
      /<w:r>(\s*<w:rPr>([\s\S]*?)<\/w:rPr>|\s*<w:rPr\/>)?/g,
      (_m, _whole, pr) => {
        let rPr: string = pr ?? ''
        if (font !== undefined) {
          rPr = rPr.replace(/<w:rFonts\b[^>]*\/>/, '')
          if (font) {
            const f = escapeXmlAttr(font)
            const el = `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`
            const lead = /^\s*<w:rStyle\b[^>]*\/>/.exec(rPr)?.[0] ?? ''
            rPr = lead + el + rPr.slice(lead.length)
          }
        }
        if (size !== undefined) {
          const sz = Math.max(2, Math.round(size))
          rPr = rPr.replace(/<w:sz(?:Cs)?\b[^>]*\/>/g, '')
          // sz and szCs sit before the trailing run properties Word writes after them
          const tail =
            /<w:(?:highlight|u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath)\b/.exec(
              rPr,
            )
          const at = tail ? tail.index : rPr.length
          rPr = `${rPr.slice(0, at)}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${rPr.slice(at)}`
        }
        return rPr ? `<w:r><w:rPr>${rPr}</w:rPr>` : '<w:r>'
      },
    )
  }
  let out =
    xml.slice(0, found.index) + `<w:rt>${inner}</w:rt>` + xml.slice(found.index + found[0].length)
  if (size !== undefined) out = setPr(out, 'hps', Math.max(2, size), 'hpsRaise|hpsBaseText|lid')
  if (raise !== undefined) out = setPr(out, 'hpsRaise', raise, 'hpsBaseText|lid')
  return out
}

/** the guide with its base text changed, in the first base run's formatting; null when it has no base run */
export function setRubyBaseXml(xml: string, base: string): string | null {
  return setRubyPartText(xml, 'w:rubyBase', base)
}

function setRubyPartText(xml: string, part: 'w:rt' | 'w:rubyBase', text: string): string | null {
  const found = new RegExp(`<${part}>([\\s\\S]*?)</${part}>`).exec(xml)
  if (!found) return null
  let placed = false
  const inner = found[1].replace(/<w:t(\s[^>]*)?>[^<]*<\/w:t>|<w:t(\s[^>]*)?\/>/g, () => {
    if (placed) return ''
    placed = true
    return `<w:t xml:space="preserve">${escapeXmlText(text)}</w:t>`
  })
  if (!placed) return null
  return (
    xml.slice(0, found.index) +
    `<${part}>${inner}</${part}>` +
    xml.slice(found.index + found[0].length)
  )
}

/**
 * A phonetic guide (w:ruby) for base text: the reading sits above it at half
 * the base size, raised clear of it, as Word writes a new one. The fragment
 * goes inside a w:r.
 */
export function generateRubyXml(
  base: string,
  rt: string,
  baseSizeHalfPoints = 21,
  lang = 'zh-CN',
  align: RubyAlign = 'center',
): string {
  const size = Math.max(2, Math.round(baseSizeHalfPoints))
  const reading = Math.max(2, Math.round(size / 2))
  return (
    `<w:ruby><w:rubyPr><w:rubyAlign w:val="${align}"/>` +
    `<w:hps w:val="${reading}"/><w:hpsRaise w:val="${Math.max(0, size - 2)}"/>` +
    `<w:hpsBaseText w:val="${size}"/><w:lid w:val="${escapeXmlAttr(lang)}"/></w:rubyPr>` +
    `<w:rt><w:r><w:rPr><w:sz w:val="${reading}"/><w:szCs w:val="${reading}"/></w:rPr>` +
    `<w:t xml:space="preserve">${escapeXmlText(rt)}</w:t></w:r></w:rt>` +
    `<w:rubyBase><w:r><w:t xml:space="preserve">${escapeXmlText(base)}</w:t></w:r></w:rubyBase></w:ruby>`
  )
}

/**
 * A table of authorities for one category of citations: the category's heading,
 * then its citations with their pages (TOA \c). The heading is the first line of
 * the field result, as Word writes it with \h.
 */
export function generateTableOfAuthoritiesFieldXml(
  heading: string,
  /** a citation made in several places lists its pages as text ("3, 7") */
  entries: Array<{ text: string; pageNo?: number | string }>,
  category: number,
  table: FieldTableLayout = {},
): string[] {
  if (entries.length === 0) return []
  return generateFieldTableXml(
    [
      { text: heading, style: 'TOAHeading' },
      ...entries.map((e) => ({ ...e, style: 'TableofAuthorities' })),
    ],
    { ...table, instr: table.instr ?? ` TOA \\h \\c "${Math.round(category)}" \\p ` },
  )
}

export type FieldTableLeader = 'none' | 'dot' | 'hyphen' | 'underscore'

/** How the lines of a generated TOC / TOA / INDEX field are laid out. */
export interface FieldTableLayout {
  /** the field instruction, when it is not the generator's default */
  instr?: string
  /** leader of the right tab the page number sits behind (default dot) */
  leader?: FieldTableLeader
  /** the page number follows the text after this separator instead of a right tab */
  pageSep?: string
  /** position of the right tab, twips (default 9350) */
  tabPos?: number
}

export interface FieldTableLine {
  text: string
  pageNo?: number | string
  /** the line never takes a page number */
  noPage?: boolean
  /** paragraph style id of the line */
  style?: string
  /** left indent, twips */
  indent?: number
  /** the page text in formatted pieces (an index entry's bold or italic pages); joined, they are `pageNo` */
  pageRuns?: FieldTablePageRun[]
  /** the text in formatted pieces (a citation's original formatting); joined, they are `text` */
  textRuns?: FieldTextRun[]
}

export type FieldTablePageRun = FieldTextRun

const fieldTextRPr = (run: FieldTextRun, noProof = false): string => {
  const inner =
    (run.bold ? '<w:b/>' : '') +
    (run.italic ? '<w:i/>' : '') +
    (run.underline ? '<w:u w:val="single"/>' : '') +
    (noProof ? '<w:noProof/>' : '')
  return inner ? `<w:rPr>${inner}</w:rPr>` : ''
}

/** a field code as runs: the formatted pieces of its quoted long text stand in for that text */
function fieldCodeRuns(code: string, formatted: FieldTextRun[] | undefined): FieldTextRun[] {
  const text = formatted?.map((r) => r.text).join('') ?? ''
  const at = text ? code.indexOf(`"${text}"`) : -1
  if (at === -1) return [{ text: code }]
  return [
    { text: code.slice(0, at + 1) },
    ...formatted!,
    { text: code.slice(at + 1 + text.length) },
  ]
}

/** a TOC / TOA / INDEX field as one w:p per line, the field begun in the first and ended in the last */
export function generateFieldTableXml(
  lines: FieldTableLine[],
  table: FieldTableLayout & { instr: string },
): string[] {
  return tocFieldLines(lines, table.instr, (line) => line.style ?? '', table)
}

function tocFieldLines<
  E extends {
    text: string
    pageNo?: number | string
    noPage?: boolean
    indent?: number
    pageRuns?: FieldTablePageRun[]
    textRuns?: FieldTextRun[]
  },
>(
  entries: E[],
  instr: string,
  styleOf: (entry: E) => string,
  table: FieldTableLayout = {},
): string[] {
  const tabbed = table.pageSep === undefined && entries.some((e) => !e.noPage)
  const tabs = tabbed
    ? `<w:tabs><w:tab w:val="right" w:leader="${table.leader ?? 'dot'}" w:pos="${table.tabPos ?? 9350}"/></w:tabs>`
    : ''
  const pPr = (entry: E) => {
    const style = styleOf(entry)
    return (
      `<w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ''}${tabs}` +
      (entry.indent ? `<w:ind w:left="${entry.indent}"/>` : '') +
      '<w:rPr><w:noProof/></w:rPr></w:pPr>'
    )
  }
  const run = (inner: string) => `<w:r><w:rPr><w:noProof/></w:rPr>${inner}</w:r>`
  const entryRuns = (entry: E) => {
    const text = entry.textRuns?.length
      ? entry.textRuns
          .map(
            (piece) =>
              `<w:r>${fieldTextRPr(piece, true)}<w:t xml:space="preserve">${escapeXmlText(piece.text)}</w:t></w:r>`,
          )
          .join('')
      : run(`<w:t xml:space="preserve">${escapeXmlText(entry.text)}</w:t>`)
    if (entry.noPage) return text
    if (entry.pageRuns?.length) {
      const pages = entry.pageRuns
        .map(
          (p) =>
            `<w:r><w:rPr>${p.bold ? '<w:b/>' : ''}${p.italic ? '<w:i/>' : ''}<w:noProof/></w:rPr>` +
            `<w:t xml:space="preserve">${escapeXmlText(p.text)}</w:t></w:r>`,
        )
        .join('')
      const lead =
        table.pageSep !== undefined
          ? run(`<w:t xml:space="preserve">${escapeXmlText(table.pageSep)}</w:t>`)
          : run('<w:tab/>')
      return text + lead + pages
    }
    const page = entry.pageNo !== undefined ? escapeXmlText(String(entry.pageNo)) : ''
    // separator and page share the last run: a refreshed page number rewrites that run alone
    if (table.pageSep !== undefined)
      return (
        text +
        run(`<w:t xml:space="preserve">${page ? escapeXmlText(table.pageSep) + page : ''}</w:t>`)
      )
    return text + run('<w:tab/>') + (page ? run(`<w:t>${page}</w:t>`) : '')
  }
  // built before the pages were laid out (they come for every line or for none): there is
  // no result to show, so Word is asked to build the table
  const paged = (e: E) => !!e.pageRuns?.length || (e.pageNo !== undefined && e.pageNo !== '')
  const unpaged = !entries.some(paged) && entries.some((e) => !e.noPage)
  const begin =
    `<w:r><w:fldChar w:fldCharType="begin"${unpaged ? ' w:dirty="true"' : ''}/></w:r>` +
    `<w:r><w:instrText xml:space="preserve">${escapeXmlText(instr)}</w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

  return entries.map((entry, i) => {
    const first = i === 0 ? begin : ''
    const last = i === entries.length - 1 ? end : ''
    return `<w:p>${pPr(entry)}${first}${entryRuns(entry)}${last}</w:p>`
  })
}

/** a generated line whose page follows a separator: its last text run holds both */
export function patchFieldTablePageXml(xml: string, tail: string): string {
  const at = xml.lastIndexOf('<w:t xml:space="preserve">')
  if (at === -1) return xml
  const open = at + '<w:t xml:space="preserve">'.length
  const close = xml.indexOf('</w:t>', open)
  return close === -1 ? xml : xml.slice(0, open) + escapeXmlText(tail) + xml.slice(close)
}

export type SeqNumberFormat = 'ARABIC' | 'alphabetic' | 'ALPHABETIC' | 'roman' | 'ROMAN'

export const SEQ_NUMBER_FORMATS: readonly SeqNumberFormat[] = [
  'ARABIC',
  'alphabetic',
  'ALPHABETIC',
  'roman',
  'ROMAN',
]

const ROMAN_PARTS: Array<[number, string]> = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
]

/** a SEQ field's number as its `\*` format switch shows it (Word repeats the letter past z: aa, bb) */
export function formatSeqNumber(n: number, format: SeqNumberFormat): string {
  if (format === 'ARABIC' || n < 1) return String(n)
  if (format === 'alphabetic' || format === 'ALPHABETIC') {
    const letter = String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1)
    return format === 'ALPHABETIC' ? letter.toUpperCase() : letter
  }
  let rest = n
  let out = ''
  for (const [value, digits] of ROMAN_PARTS) {
    while (rest >= value) {
      out += digits
      rest -= value
    }
  }
  return format === 'ROMAN' ? out : out.toLowerCase()
}

export interface CaptionOptions {
  format?: SeqNumberFormat
  /** Word's "Exclude label from caption": the number alone, still counted under the label */
  excludeLabel?: boolean
  /** the number is a guess (the label's numbering is not computed here): Word renumbers on open */
  dirty?: boolean
}

/**
 * Caption paragraph: `<label> <SEQ label> <text>`, e.g. "Figure 1 System architecture".
 * The number is the SEQ field's result, saved as Word saves it.
 */
export function generateCaptionXml(
  label: string,
  number: number,
  text: string,
  options: CaptionOptions = {},
): string {
  const format = options.format ?? 'ARABIC'
  const rPr = '<w:rPr><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
  const run = (inner: string) => `<w:r>${rPr}${inner}</w:r>`
  // a label with a space has to be quoted or SEQ reads its first word as the identifier
  const id = /\s/.test(label) ? `"${label.replace(/"/g, '')}"` : label
  return (
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="80" w:after="200"/></w:pPr>' +
    (options.excludeLabel ? '' : run(`<w:t xml:space="preserve">${escapeXmlText(label)} </w:t>`)) +
    run(`<w:fldChar w:fldCharType="begin"${options.dirty ? ' w:dirty="true"' : ''}/>`) +
    run(
      `<w:instrText xml:space="preserve"> SEQ ${escapeXmlText(id)} \\* ${format} </w:instrText>`,
    ) +
    run('<w:fldChar w:fldCharType="separate"/>') +
    run(`<w:t>${formatSeqNumber(number, format)}</w:t>`) +
    run('<w:fldChar w:fldCharType="end"/>') +
    (text ? run(`<w:t xml:space="preserve"> ${escapeXmlText(text)}</w:t>`) : '') +
    '</w:p>'
  )
}

/**
 * INDEX field as one w:p per cached entry line, alphabetically sorted. The
 * lines carry no pages, so the field is dirty and Word builds it on open.
 */
export function generateIndexFieldXml(terms: string[]): string[] {
  const unique = [...new Set(terms.map((t) => t.trim()).filter(Boolean))]
  if (unique.length === 0) return []
  unique.sort((a, b) => a.localeCompare(b, 'zh-CN'))
  return generateFieldTableXml(
    unique.map((text) => ({ text })),
    { instr: ' INDEX \\c "2" ', tabPos: 4300 },
  )
}

/**
 * The instruction runs of a field. A mail-merge IF, SKIPIF or NEXTIF carries its nested
 * MERGEFIELD operands as «Name»; Word only reads them as nested fields, so
 * each one is written back as a field of its own.
 */
function fieldInstructionXml(instr: string): string {
  const text = (part: string): string =>
    `<w:r><w:instrText xml:space="preserve">${escapeXmlText(part)}</w:instrText></w:r>`
  if (!MERGE_RULE_INSTR_RE.test(instr)) return text(` ${instr} `)
  return ` ${instr} `
    .split(/(«[^»]+»)/)
    .map((part) => {
      const name = /^«([^»]+)»$/.exec(part)?.[1]
      if (name === undefined) return part === '' ? '' : text(part)
      const code = /\s/.test(name) ? `"${name}"` : name
      return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        text(` MERGEFIELD ${code} `) +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        `<w:r><w:t>${escapeXmlText(part)}</w:t></w:r>` +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
      )
    })
    .join('')
}

function generateRunsXml(runs: Run[], ctx: GenerateContext): string {
  return runsXml(runs, ctx.allocateHyperlinkRel, ctx.allocateRevisionId, ctx)
}

/** OOXML runs without relationship allocation (header/footer parts, ...) */
export function inlineRunsXml(runs: Run[]): string {
  return runsXml(runs, null)
}

/**
 * OOXML runs for a run list. `allocate` mints rIds for new external links;
 * pass null where no relationship allocation is possible (textbox patches) —
 * links without an existing rId then degrade to plain runs.
 */
function runsXml(
  runs: Run[],
  allocate: ((href: string) => string) | null,
  allocateRevisionId?: () => string,
  bookmarkCtx: Pick<GenerateContext, 'allocateBookmarkId'> | null = null,
): string {
  // comment ranges: re-emit start/end/reference markers around the first..last
  // run each id covers, so editing a commented paragraph keeps its comments
  const firstOf = new Map<string, number>()
  const lastOf = new Map<string, number>()
  runs.forEach((run, i) => {
    for (const id of run.commentIds ?? []) {
      if (!firstOf.has(id)) firstOf.set(id, i)
      lastOf.set(id, i)
    }
  })
  const startsAt = (i: number) =>
    [...firstOf]
      .filter(([, at]) => at === i)
      .map(([id]) => `<w:commentRangeStart w:id="${escapeXmlAttr(id)}"/>`)
      .join('')
  const endsAt = (i: number) =>
    [...lastOf]
      .filter(([, at]) => at === i)
      .map(
        ([id]) =>
          `<w:commentRangeEnd w:id="${escapeXmlAttr(id)}"/>` +
          `<w:r><w:commentReference w:id="${escapeXmlAttr(id)}"/></w:r>`,
      )
      .join('')

  const parts: string[] = []

  // a NOTEREF field finds its note through a bookmark on the mark itself, in a link or not
  const fragment = (run: Run, insideLink: boolean): string => {
    const onMark = bookmarksXml(bookmarkCtx, run.noteRef?.bookmarks ?? [])
    return onMark.starts + runFragmentXml(run, insideLink) + onMark.ends
  }

  // serialize run indices [from, to) handling hyperlink grouping + comment markers
  const emitRange = (from: number, to: number) => {
    let i = from
    while (i < to) {
      const run = runs[i]
      if (run.link) {
        // group consecutive runs sharing the same link target; runs parsed from
        // distinct w:hyperlink elements (own rId) stay separate so no relationship is orphaned
        const group: Run[] = []
        const groupStart = i
        const href = run.link.href
        let rId = run.link.rId
        while (i < to && runs[i].link && runs[i].link!.href === href) {
          const next = runs[i].link!.rId
          if (group.length > 0 && next !== undefined && rId !== undefined && next !== rId) break
          rId = rId ?? next
          group.push(runs[i])
          i++
        }
        // a comment boundary inside the link group widens to the whole hyperlink
        for (let j = groupStart; j < i; j++) parts.push(startsAt(j))
        const tooltip = group[0]?.link?.tooltip
        const tipAttr = tooltip ? ` w:tooltip="${escapeXmlAttr(tooltip)}"` : ''
        if (href.startsWith('#')) {
          const inner = group.map((r) => fragment(r, true)).join('')
          parts.push(
            `<w:hyperlink w:anchor="${escapeXmlAttr(href.slice(1))}"${tipAttr}>${inner}</w:hyperlink>`,
          )
        } else {
          const finalRId = rId ?? allocate?.(href)
          if (finalRId) {
            const inner = group.map((r) => fragment(r, true)).join('')
            parts.push(
              `<w:hyperlink r:id="${escapeXmlAttr(finalRId)}"${tipAttr}>${inner}</w:hyperlink>`,
            )
          } else {
            parts.push(group.map((r) => fragment(r, false)).join(''))
          }
        }
        for (let j = groupStart; j < i; j++) parts.push(endsAt(j))
      } else {
        parts.push(startsAt(i))
        parts.push(fragment(run, false))
        parts.push(endsAt(i))
        i++
      }
    }
  }

  // tracked changes: wrap consecutive runs of the same revision in w:ins / w:del.
  // An insertion later deleted (both set) nests w:del inside w:ins, as Word does.
  const revKey = (r: Run) =>
    r.ins || r.del
      ? JSON.stringify([
          r.ins?.author ?? null,
          r.ins?.date ?? null,
          r.ins?.dateUtc ?? null,
          r.ins?.id ?? null,
          r.ins?.move ?? null,
          r.del?.author ?? null,
          r.del?.date ?? null,
          r.del?.dateUtc ?? null,
          r.del?.id ?? null,
          r.del?.move ?? null,
        ])
      : ''
  let revSeq = 9001
  const mintRevisionId = allocateRevisionId ?? (() => String(revSeq++))
  // a later edit splitting one w:ins in two leaves both halves with the stored id
  const emittedIds = new Set<string>()
  const revisionId = (info: RevisionInfo) => {
    let id = info.id && !emittedIds.has(info.id) ? info.id : mintRevisionId()
    while (emittedIds.has(id)) id = mintRevisionId()
    emittedIds.add(id)
    return id
  }
  const revAttrs = (info: RevisionInfo) =>
    ` w:id="${escapeXmlAttr(revisionId(info))}"` +
    ` w:author="${escapeXmlAttr(info.author)}"` +
    (info.date ? ` w:date="${escapeXmlAttr(info.date)}"` : '') +
    (info.dateUtc ? ` w16du:dateUtc="${escapeXmlAttr(info.dateUtc)}"` : '')

  // a move wraps its runs in w:moveFrom / w:moveTo inside a named range
  // (w:moveFromRangeStart .. End); the twin side carries the same w:name
  const moveOf = (r: Run | undefined): { side: 'From' | 'To'; name: string } | null =>
    r?.del?.move
      ? { side: 'From', name: r.del.move }
      : r?.ins?.move
        ? { side: 'To', name: r.ins.move }
        : null
  const sameMove = (a: ReturnType<typeof moveOf>, b: ReturnType<typeof moveOf>) =>
    !!a && !!b && a.side === b.side && a.name === b.name
  let openRange: { side: 'From' | 'To'; id: string } | null = null
  let g = 0
  while (g < runs.length) {
    const key = revKey(runs[g])
    let end = g
    while (end < runs.length && revKey(runs[end]) === key) end++
    if (key === '') {
      emitRange(g, end)
    } else {
      const { ins, del } = runs[g]
      const move = moveOf(runs[g])
      if (move && !sameMove(move, moveOf(runs[g - 1]))) {
        const info = move.side === 'From' ? del! : ins!
        const id = mintRevisionId()
        parts.push(
          `<w:move${move.side}RangeStart w:id="${escapeXmlAttr(id)}"` +
            ` w:author="${escapeXmlAttr(info.author)}"` +
            (info.date ? ` w:date="${escapeXmlAttr(info.date)}"` : '') +
            ` w:name="${escapeXmlAttr(move.name)}"/>`,
        )
        openRange = { side: move.side, id }
      }
      const insTag = ins?.move ? 'w:moveTo' : 'w:ins'
      const delTag = del?.move ? 'w:moveFrom' : 'w:del'
      if (ins) parts.push(`<${insTag}${revAttrs(ins)}>`)
      if (del) parts.push(`<${delTag}${revAttrs(del)}>`)
      emitRange(g, end)
      if (del) parts.push(`</${delTag}>`)
      if (ins) parts.push(`</${insTag}>`)
      if (openRange && !sameMove(move, moveOf(runs[end]))) {
        parts.push(`<w:move${openRange.side}RangeEnd w:id="${escapeXmlAttr(openRange.id)}"/>`)
        openRange = null
      }
    }
    g = end
  }
  return parts.join('')
}

/**
 * One run's OOXML, including atomic constructs the plain-text run cannot
 * express: footnote/endnote reference markers and trailing XE index fields.
 */
/** Runs whose text mixes box glyphs with typed characters: each glyph becomes
 *  its own field / control, the characters between them plain runs. */
function splitGlyphRuns(
  run: Run,
  isGlyph: (ch: string) => boolean,
  glyphXml: (ch: string) => string,
  plainXml: (text: string) => string,
): string {
  let out = ''
  let plain = ''
  for (const ch of run.text) {
    if (isGlyph(ch)) {
      if (plain) out += plainXml(plain)
      plain = ''
      out += glyphXml(ch)
    } else plain += ch
  }
  if (plain) out += plainXml(plain)
  return out
}

function runFragmentXml(run: Run, insideLink: boolean): string {
  // atomic inline formula: the stored <m:oMath> fragment is already valid
  // paragraph content (patch.ts adds the xmlns:m declaration when missing)
  if (run.math) {
    return run.math.display ? `<m:oMathPara>${run.math.omml}</m:oMathPara>` : run.math.omml
  }
  // atomic phonetic guide: the <w:ruby> fragment re-wrapped in a run
  if (run.ruby) return `<w:r>${rubyWithBaseFormat(run, insideLink)}</w:r>`
  // enclosed or combined characters: the EQ field's own runs
  if (run.eqLayout) return eqLayoutFieldXml(run, insideLink)
  // atomic cell picture: the exact <w:drawing> fragment re-wrapped in a run
  if (run.image) {
    const text = run.text === '' ? '' : generateRunXml({ ...run, image: undefined }, insideLink)
    return `${text}<w:r>${run.rawRPr ?? ''}${run.image.xml}</w:r>`
  }
  if (run.sym) {
    return (
      `<w:r>${runRPrXml(run, insideLink)}<w:sym w:font="${escapeXmlAttr(run.sym.font)}"` +
      ` w:char="${escapeXmlAttr(run.sym.char)}"/></w:r>`
    )
  }
  if (run.noteRef) {
    const tag = run.noteRef.kind === 'footnote' ? 'w:footnoteReference' : 'w:endnoteReference'
    // a marker parsed from the file keeps its own rPr bytes verbatim — rStyle,
    // fonts, size — whatever else in the paragraph the save rebuilds (the
    // marker is an atom no text op can target, so there is nothing to merge).
    // A marker the app minted has no rPr and keeps the superscript it has
    // always carried.
    const markerRPr = run.rawRPr ?? '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
    const { customMark } = run.noteRef
    if (customMark === undefined)
      return `<w:r>${markerRPr}<${tag} w:id="${escapeXmlAttr(run.noteRef.id)}"/></w:r>`
    const markText = customMark
      ? `<w:t xml:space="preserve">${escapeXmlText(customMark)}</w:t>`
      : ''
    return `<w:r>${markerRPr}<${tag} w:customMarkFollows="1" w:id="${escapeXmlAttr(run.noteRef.id)}"/>${markText}</w:r>`
  }
  if (run.refField !== undefined) {
    // cross-reference: full REF field, run text is the cached display result;
    // the original instruction (with its switches) is written back verbatim
    const name = run.refField.replace(/"/g, '')
    const instr = run.refInstr ?? ` REF ${name} \\h `
    return (
      `<w:r><w:fldChar w:fldCharType="begin"${run.fldDirty ? ' w:dirty="true"' : ''}/></w:r>` +
      `<w:r><w:instrText xml:space="preserve">${escapeXmlText(instr)}</w:instrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      generateRunXml({ ...run, refField: undefined, fldDirty: undefined }, insideLink) +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )
  }
  if (run.sdtCheckboxXml) {
    // Each box glyph in the run is one control sharing the same w:sdtPr, with
    // w14:checked set from the glyph; other characters are plain text beside it.
    // No glyph left means the user typed over the box, and the control goes.
    const glyphs = sdtCheckboxGlyphs(run.sdtCheckboxXml)
    const inner = (text: string) =>
      generateRunXml({ ...run, text, sdtCheckboxXml: undefined }, insideLink)
    return splitGlyphRuns(
      run,
      (ch) => ch === glyphs.checked || ch === glyphs.unchecked,
      (ch) =>
        `<w:sdt>${syncSdtCheckbox(run.sdtCheckboxXml!, sdtCheckboxIsChecked(run.sdtCheckboxXml!, ch))}` +
        `<w:sdtContent>${inner(ch)}</w:sdtContent></w:sdt>`,
      inner,
    )
  }
  if (run.instrField !== undefined) {
    // Generic inline field: run text is the cached result; the instruction is written back verbatim.
    // A preserved begin run (form fields: w:ffData) replaces the bare begin, and the run text is a
    // synthesized glyph Word must not see as a cached result — Word draws the box from ffData.
    const instrXml =
      fieldInstructionXml(run.instrField) + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    const endXml = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    if (run.text === NO_FIELD_RESULT)
      return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        fieldInstructionXml(run.instrField) +
        endXml
      )
    if (run.zoteroFieldPart && run.zoteroFieldPart !== 'single') {
      const cachedXml = generateRunXml(
        {
          ...run,
          instrField: undefined,
          zoteroFieldId: undefined,
          zoteroFieldPart: undefined,
        },
        insideLink,
      )
      if (run.zoteroFieldPart === 'begin') {
        return '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' + instrXml + cachedXml
      }
      if (run.zoteroFieldPart === 'end') return cachedXml + endXml
      return cachedXml
    }
    if (run.fldBeginXml) {
      // Adjacent identical checkboxes merge into one text node in the editor
      // (equal marks), so each ☐/☒ glyph in the run is one field sharing the
      // same ffData; characters typed beside them survive as plain text runs.
      // The glyph is the source of truth for the checked state: an in-editor
      // ☐↔☒ edit must land in w:checked or Word keeps the old state.
      const syncedField = (checked: boolean): string => {
        const val = `<w:checked w:val="${checked ? '1' : '0'}"/>`
        let begin = run.fldBeginXml!.replace(/<w:checked(?:\s[^>]*)?\/>/g, '')
        if (begin.includes('</w:checkBox>'))
          begin = begin.replace('</w:checkBox>', `${val}</w:checkBox>`)
        else
          begin = begin.replace(/<w:checkBox((?:\s[^>]*)?)\/>/, `<w:checkBox$1>${val}</w:checkBox>`)
        return begin + instrXml + endXml
      }
      // no glyph left = the user typed over / deleted the checkbox; Word also
      // removes the form field then, so only the replacement text survives
      return splitGlyphRuns(
        run,
        (ch) => ch === '☐' || ch === '☒',
        (ch) => syncedField(ch === '☒'),
        (plain) =>
          generateRunXml(
            { ...run, text: plain, instrField: undefined, fldBeginXml: undefined },
            insideLink,
          ),
      )
    }
    return (
      `<w:r><w:fldChar w:fldCharType="begin"${run.fldDirty ? ' w:dirty="true"' : ''}/></w:r>` +
      instrXml +
      generateRunXml({ ...run, instrField: undefined, fldDirty: undefined }, insideLink) +
      endXml
    )
  }
  let xml = run.text === '' ? '' : generateRunXml(run, insideLink)
  const kept = run.taInstr !== undefined ? `TA ${run.taInstr}` : run.tcInstr && `TC ${run.tcInstr}`
  if (kept) {
    // a citation or contents entry rides the index entry's slot: its instruction goes back as it was read
    xml +=
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      fieldCodeRuns(` ${kept} `, run.taRuns)
        .map(
          (piece) =>
            `<w:r>${fieldTextRPr(piece)}<w:instrText xml:space="preserve">${escapeXmlText(piece.text)}</w:instrText></w:r>`,
        )
        .join('') +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  } else if (run.xeTerm !== undefined) {
    // embedded quotes would break the field instruction syntax; drop them
    const term = run.xeTerm.replace(/"/g, '')
    xml +=
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      `<w:r><w:instrText xml:space="preserve"> XE "${escapeXmlText(term)}" ${
        run.xeSwitches ? escapeXmlText(run.xeSwitches) + ' ' : ''
      }</w:instrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  }
  return xml
}

const EMU_PER_POINT = 12700

/** a w14 colour with its transparency, as the reader takes it back */
function w14ColorXml(color: string, alpha: number | undefined): string {
  const inner =
    alpha !== undefined && alpha < 1 ? `<w14:alpha w14:val="${Math.round(alpha * 100000)}"/>` : ''
  const hex = escapeXmlAttr(color)
  return inner
    ? `<w14:srgbClr w14:val="${hex}">${inner}</w14:srgbClr>`
    : `<w14:srgbClr w14:val="${hex}"/>`
}

function w14TextFillXml(color: string): string {
  return `<w14:textFill><w14:solidFill>${w14ColorXml(color, undefined)}</w14:solidFill></w14:textFill>`
}

/** the literal colour of a raw solid w14:textFill; '' when the fill is solid in a scheme colour */
function rawSolidFill(xml: string | undefined): string | undefined {
  const solid = /^<w14:textFill>\s*<w14:solidFill>\s*<w14:(srgbClr|schemeClr)\b[^>]*>/.exec(
    xml ?? '',
  )
  if (!solid) return undefined
  return solid[1] === 'srgbClr' ? (rawAttr(solid[0], 'w14:val') ?? '').toUpperCase() : ''
}

function w14GlowXml(glow: TextGlow): string {
  const rad = Math.round(glow.radiusPt * EMU_PER_POINT)
  return `<w14:glow w14:rad="${rad}">${w14ColorXml(glow.color, glow.alpha)}</w14:glow>`
}

function w14ShadowXml(shadow: TextShadow): string {
  return (
    `<w14:shadow w14:blurRad="${Math.round(shadow.blurPt * EMU_PER_POINT)}" ` +
    `w14:dist="${Math.round(shadow.distPt * EMU_PER_POINT)}" ` +
    `w14:dir="${Math.round(shadow.dirDeg * 60000)}" ` +
    'w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">' +
    `${w14ColorXml(shadow.color, shadow.alpha)}</w14:shadow>`
  )
}

function w14ReflectionXml(reflection: TextReflection): string {
  const pct = (v: number) => Math.round(v * 100000)
  return (
    `<w14:reflection w14:blurRad="${Math.round(reflection.blurPt * EMU_PER_POINT)}" ` +
    `w14:stA="${pct(reflection.startAlpha)}" w14:stPos="0" w14:endA="300" ` +
    `w14:endPos="${pct(reflection.endPos)}" ` +
    `w14:dist="${Math.round(reflection.distPt * EMU_PER_POINT)}" ` +
    'w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000" w14:sy="-100000" ' +
    'w14:kx="0" w14:ky="0" w14:algn="bl"/>'
  )
}

function w14TextOutlineXml(outline: TextOutline): string {
  return (
    `<w14:textOutline w14:w="${Math.round(outline.widthPt * EMU_PER_POINT)}" ` +
    'w14:cap="rnd" w14:cmpd="sng" w14:algn="ctr">' +
    `<w14:solidFill>${w14ColorXml(outline.color, outline.alpha)}</w14:solidFill>` +
    '<w14:prstDash w14:val="solid"/><w14:bevel/></w14:textOutline>'
  )
}

/** CT_RPr child sequence (subset), for schema-ordered assembly and merging */
export const RPR_CHILD_ORDER = [
  'w:rStyle',
  'w:rFonts',
  'w:b',
  'w:bCs',
  'w:i',
  'w:iCs',
  'w:caps',
  'w:smallCaps',
  'w:strike',
  'w:dstrike',
  'w:outline',
  'w:shadow',
  'w:emboss',
  'w:imprint',
  'w:noProof',
  'w:snapToGrid',
  'w:vanish',
  'w:webHidden',
  'w:color',
  'w:spacing',
  'w:w',
  'w:kern',
  'w:position',
  'w:sz',
  'w:szCs',
  'w:highlight',
  'w:u',
  'w:effect',
  'w:bdr',
  'w:shd',
  'w:fitText',
  'w:vertAlign',
  'w:rtl',
  'w:cs',
  'w:em',
  'w:lang',
  'w:eastAsianLayout',
  'w:specVanish',
  'w:oMath',
  'w14:glow',
  'w14:shadow',
  'w14:reflection',
  'w14:textOutline',
  'w14:textFill',
  'w14:ligatures',
  'w14:numForm',
  'w14:numSpacing',
  'w14:stylisticSets',
  'w14:cntxtAlts',
  'w:rPrChange',
]

/** rPr children modeled by the run model, grouped by modeled field (tags within a group are kept or rebuilt together) */
const RUN_MANAGED_GROUPS: Array<{ key: string; tags: string[] }> = [
  { key: 'rStyle', tags: ['w:rStyle'] },
  { key: 'rFonts', tags: ['w:rFonts'] },
  { key: 'bold', tags: ['w:b', 'w:bCs'] },
  { key: 'italic', tags: ['w:i', 'w:iCs'] },
  { key: 'strike', tags: ['w:strike'] },
  { key: 'caps', tags: ['w:caps', 'w:smallCaps'] },
  { key: 'dstrike', tags: ['w:dstrike'] },
  { key: 'vanish', tags: ['w:vanish'] },
  { key: 'color', tags: ['w:color'] },
  { key: 'spacing', tags: ['w:spacing'] },
  { key: 'scale', tags: ['w:w'] },
  { key: 'kern', tags: ['w:kern'] },
  { key: 'position', tags: ['w:position'] },
  { key: 'size', tags: ['w:sz', 'w:szCs'] },
  { key: 'highlight', tags: ['w:highlight'] },
  { key: 'underline', tags: ['w:u'] },
  { key: 'bdr', tags: ['w:bdr'] },
  { key: 'fitText', tags: ['w:fitText'] },
  { key: 'eaLayout', tags: ['w:eastAsianLayout'] },
  { key: 'lang', tags: ['w:lang'] },
  { key: 'noProof', tags: ['w:noProof'] },
  { key: 'shading', tags: ['w:shd'] },
  { key: 'vertAlign', tags: ['w:vertAlign'] },
  { key: 'rtl', tags: ['w:rtl'] },
  { key: 'em', tags: ['w:em'] },
  { key: 'otLigatures', tags: ['w14:ligatures'] },
  { key: 'otNumForm', tags: ['w14:numForm'] },
  { key: 'otNumSpacing', tags: ['w14:numSpacing'] },
  { key: 'otSets', tags: ['w14:stylisticSets'] },
  { key: 'otAlts', tags: ['w14:cntxtAlts'] },
  { key: 'fxGlow', tags: ['w14:glow'] },
  { key: 'fxShadow', tags: ['w14:shadow'] },
  { key: 'fxReflection', tags: ['w14:reflection'] },
  { key: 'fxOutline', tags: ['w14:textOutline'] },
  { key: 'fxFill', tags: ['w14:textFill'] },
  { key: 'rPrChange', tags: ['w:rPrChange'] },
]

const rawAttr = (xml: string | undefined, attr: string): string | undefined =>
  xml ? new RegExp(` ${attr}="([^"]*)"`).exec(xml)?.[1] : undefined

/** the switched-on w14:styleSet ids of a raw w14:stylisticSets, as parse reads them */
function rawStyleSets(xml: string | undefined): number[] {
  return [...(xml ?? '').matchAll(/<w14:styleSet\b[^>]*>/g)]
    .filter((m) => !/^(0|false|off)$/.test(rawAttr(m[0], 'w14:val') ?? ''))
    .map((m) => parseInt(rawAttr(m[0], 'w14:id') ?? '', 10))
    .filter((id) => id >= 1 && id <= 20)
}

type RunBorder = NonNullable<Run['bdr']>

/** the run's own w:bdr as parse reads it: null for none/nil, undefined when there is no element */
function rawRunBorder(xml: string | undefined): RunBorder | null | undefined {
  if (!xml) return undefined
  const val = rawAttr(xml, 'w:val')
  if (!val || val === 'none' || val === 'nil') return null
  const color = rawAttr(xml, 'w:color')
  const space = parseInt(rawAttr(xml, 'w:space') ?? '', 10)
  return {
    val,
    sz: parseInt(rawAttr(xml, 'w:sz') ?? '', 10) || 4,
    ...(color && color !== 'auto' ? { color: color.replace(/^#/, '') } : {}),
    ...(space > 0 ? { space } : {}),
  }
}

function sameRunBorder(a: RunBorder | null | undefined, b: RunBorder | null): boolean {
  if (!a || !b) return a === b
  return (
    a.val === b.val && a.sz === b.sz && a.color === b.color && (a.space ?? 0) === (b.space ?? 0)
  )
}

/** w:bdr for a run; null writes the explicit "no border" that overrides a style's */
function runBorderXml(bdr: RunBorder | null): string {
  if (!bdr) return '<w:bdr w:val="nil"/>'
  return `<w:bdr w:val="${escapeXmlAttr(bdr.val)}" w:sz="${bdr.sz}" w:space="${bdr.space ?? 0}" w:color="${escapeXmlAttr(bdr.color ?? 'auto')}"/>`
}

/** String version of boolProp: element present means true unless w:val negates it */
function rawBool(xml: string | undefined): boolean {
  return rawOnOff(xml) === true
}

/** Tri-state read of a toggle element: undefined when absent */
function rawOnOff(xml: string | undefined): boolean | undefined {
  if (!xml) return undefined
  const val = rawAttr(xml, 'w:val')
  if (val === undefined) return true
  return !['0', 'false', 'none', 'off'].includes(val.toLowerCase())
}

const rawInt = (xml: string | undefined): number | undefined => {
  const v = rawAttr(xml, 'w:val')
  if (v === undefined) return undefined
  const n = parseInt(v, 10)
  return Number.isNaN(n) ? undefined : n
}

/** Mirrors the parse side: w:caps wins over w:smallCaps; either explicitly off reads as 'none' */
function rawCaps(caps: string | undefined, smallCaps: string | undefined): Run['caps'] {
  const c = rawOnOff(caps)
  const sc = rawOnOff(smallCaps)
  if (c) return 'all'
  if (sc) return 'small'
  if (c === false || sc === false) return 'none'
  return undefined
}

const onOffXml = (tag: string, on: boolean): string => (on ? `<${tag}/>` : `<${tag} w:val="0"/>`)

/** rFonts built from the model alone (no original element to preserve).
 * Latin-only keeps eastAsia empty; a lone primary font fills every slot (legacy behavior);
 * a complex-script font fills the cs slot instead of mirroring the ascii font. */
function freshRFontsXml(
  font: string | undefined,
  fontAscii: string | undefined,
  fontCs?: string,
  eastAsiaFont?: string,
): string {
  // Older callers use a lone primary font for every slot. Explicit slot edits
  // must leave the other slots absent so their style/theme inheritance survives.
  const legacy = font && fontAscii === undefined && eastAsiaFont === undefined ? font : undefined
  const ascii = fontAscii ?? legacy
  // a Latin-only run carries the same face in both slots; writing it as eastAsia
  // would pin CJK to the Latin font (see mergeRFontsXml)
  const ea = eastAsiaFont ?? (fontAscii !== undefined && font === fontAscii ? undefined : font)
  const cs = fontCs ?? legacy
  return `<w:rFonts${ascii ? ` w:ascii="${escapeXmlAttr(ascii)}"` : ''}${ea ? ` w:eastAsia="${escapeXmlAttr(ea)}"` : ''}${ascii ? ` w:hAnsi="${escapeXmlAttr(ascii)}"` : ''}${cs ? ` w:cs="${escapeXmlAttr(cs)}"` : ''}/>`
}

/**
 * Rebuild rFonts merging the model into the original attributes: only the slots the
 * model holds are overwritten (their theme attrs dropped so the explicit value wins);
 * cs/hint and any untouched slot keep their original values.
 */
function mergeRFontsXml(rawXml: string, run: Run): string {
  const attrs = new Map<string, string>()
  for (const m of rawXml.matchAll(/ ([\w:]+)="([^"]*)"/g)) attrs.set(m[1], m[2])
  // run.font may just be the parse-side derivation of an ascii-only element; writing
  // it back would invent an eastAsia slot that pins CJK to the old Latin font. Only
  // write eastAsia when the slot already existed or the user actually changed it.
  const rawPrimary = attrs.get('w:eastAsia') ?? attrs.get('w:ascii') ?? attrs.get('w:hAnsi')
  const hadEastAsia = attrs.has('w:eastAsia') || attrs.has('w:eastAsiaTheme')
  const set = (slot: string, theme: string, value: string) => {
    attrs.set(slot, escapeXmlAttr(value))
    attrs.delete(theme)
  }
  // slots still equal to their theme-resolved value are untouched: keep the theme attrs
  if (run.fontAscii && run.fontAscii !== run.themeRFonts?.fontAscii) {
    set('w:ascii', 'w:asciiTheme', run.fontAscii)
    set('w:hAnsi', 'w:hAnsiTheme', run.fontAscii)
  }
  if (
    run.font &&
    run.font !== run.themeRFonts?.font &&
    (hadEastAsia || run.font !== rawPrimary || run.eastAsiaFont !== undefined)
  ) {
    set('w:eastAsia', 'w:eastAsiaTheme', run.font)
  }
  if (run.fontCs) set('w:cs', 'w:cstheme', run.fontCs)
  return `<w:rFonts${[...attrs].map(([k, v]) => ` ${k}="${v}"`).join('')}/>`
}

function revisionRPrChangeXml(run: Run): string | null {
  const change = run.rPrChange
  if (!change) return null
  const attrs =
    ` w:id="${escapeXmlAttr(change.id ?? '0')}"` +
    ` w:author="${escapeXmlAttr(change.author)}"` +
    (change.date ? ` w:date="${escapeXmlAttr(change.date)}"` : '') +
    (change.dateUtc ? ` w16du:dateUtc="${escapeXmlAttr(change.dateUtc)}"` : '')
  if (change.oldRPr !== undefined) {
    const rPr = change.oldRPr ? `<w:rPr>${change.oldRPr}</w:rPr>` : '<w:rPr/>'
    return `<w:rPrChange${attrs}>${rPr}</w:rPrChange>`
  }
  const old = change.old ?? {}
  const props: string[] = []
  if (old.styleId) props.push(`<w:rStyle w:val="${escapeXmlAttr(old.styleId)}"/>`)
  if (old.font || old.fontAscii) props.push(freshRFontsXml(old.font, old.fontAscii))
  // no Cs twins here: this nested w:rPr records what the run looked like before the tracked
  // revision, and old.bold cannot tell w:b from w:b plus w:bCs. Adding them would invent a
  // complex-script flag the document never had, so rejecting the revision would bold Arabic
  // that was never bold.
  if (old.bold) props.push('<w:b/>')
  if (old.italic) props.push('<w:i/>')
  if (old.strike) props.push('<w:strike/>')
  if (old.color) props.push(`<w:color w:val="${escapeXmlAttr(old.color)}"/>`)
  if (old.charSpacingTwips) props.push(`<w:spacing w:val="${old.charSpacingTwips}"/>`)
  if (old.charScalePct) props.push(`<w:w w:val="${old.charScalePct}"/>`)
  if (old.sizeHalfPoints) {
    props.push(`<w:sz w:val="${old.sizeHalfPoints}"/>`)
    props.push(`<w:szCs w:val="${old.sizeHalfPoints}"/>`)
  }
  if (old.highlight) props.push(`<w:highlight w:val="${escapeXmlAttr(old.highlight)}"/>`)
  if (old.underline) props.push('<w:u w:val="single"/>')
  if (old.vertAlign) props.push(`<w:vertAlign w:val="${old.vertAlign}"/>`)
  return `<w:rPrChange${attrs}><w:rPr>${props.join('')}</w:rPr></w:rPrChange>`
}

/** Fresh rPr children for the modeled fields (one-to-one with what buildRun reads on the parse side) */
function modelRPrChildren(run: Run, insideLink: boolean): PPrChild[] {
  const out: PPrChild[] = []
  // a link run keeps the document's own character style (localized ids like "ae")
  const styleId = run.styleId ?? (insideLink && !run.link?.plain ? 'Hyperlink' : undefined)
  if (styleId) out.push({ name: 'w:rStyle', xml: `<w:rStyle w:val="${escapeXmlAttr(styleId)}"/>` })
  if (run.font || run.fontAscii || run.fontCs) {
    out.push({
      name: 'w:rFonts',
      xml: freshRFontsXml(run.font, run.fontAscii, run.fontCs, run.eastAsiaFont),
    })
  }
  // the Cs twins carry the same flag for complex-script text; without them clicking Bold
  // on Arabic or Hebrew changes nothing on screen, which is what Word writes too
  if (run.bold) {
    out.push({ name: 'w:b', xml: '<w:b/>' }, { name: 'w:bCs', xml: '<w:bCs/>' })
  }
  if (run.italic) {
    out.push({ name: 'w:i', xml: '<w:i/>' }, { name: 'w:iCs', xml: '<w:iCs/>' })
  }
  if (run.strike) out.push({ name: 'w:strike', xml: '<w:strike/>' })
  if (run.caps === 'all') out.push({ name: 'w:caps', xml: '<w:caps/>' })
  else if (run.caps === 'small') out.push({ name: 'w:smallCaps', xml: '<w:smallCaps/>' })
  else if (run.caps === 'none') {
    out.push(
      { name: 'w:caps', xml: onOffXml('w:caps', false) },
      { name: 'w:smallCaps', xml: onOffXml('w:smallCaps', false) },
    )
  }
  if (run.dstrike !== undefined)
    out.push({ name: 'w:dstrike', xml: onOffXml('w:dstrike', run.dstrike) })
  if (run.vanishOwn !== undefined)
    out.push({ name: 'w:vanish', xml: onOffXml('w:vanish', run.vanishOwn) })
  if (run.bdrOwn !== undefined) out.push({ name: 'w:bdr', xml: runBorderXml(run.bdrOwn) })
  if (run.fitTextTwips)
    out.push({ name: 'w:fitText', xml: `<w:fitText w:val="${Math.round(run.fitTextTwips)}"/>` })
  if (run.eaLayout && (run.eaLayout.combine || run.eaLayout.vert)) {
    const { combine, vert, id } = run.eaLayout
    out.push({
      name: 'w:eastAsianLayout',
      xml:
        `<w:eastAsianLayout${id === undefined ? '' : ` w:id="${id}"`}` +
        `${combine ? ' w:combine="1"' : ''}${vert ? ' w:vert="1"' : ''}/>`,
    })
  }
  if (run.lang && (run.lang.val || run.lang.eastAsia || run.lang.bidi)) {
    const slot = (attr: string, value?: string) =>
      value ? ` ${attr}="${escapeXmlAttr(value)}"` : ''
    out.push({
      name: 'w:lang',
      xml: `<w:lang${slot('w:val', run.lang.val)}${slot('w:eastAsia', run.lang.eastAsia)}${slot('w:bidi', run.lang.bidi)}/>`,
    })
  }
  if (run.noProof) out.push({ name: 'w:noProof', xml: '<w:noProof/>' })
  if (run.color)
    out.push({ name: 'w:color', xml: `<w:color w:val="${escapeXmlAttr(run.color)}"/>` })
  if (run.charSpacingTwips !== undefined)
    out.push({ name: 'w:spacing', xml: `<w:spacing w:val="${run.charSpacingTwips}"/>` })
  if (run.charScalePct) out.push({ name: 'w:w', xml: `<w:w w:val="${run.charScalePct}"/>` })
  if (run.kernHalfPoints !== undefined)
    out.push({ name: 'w:kern', xml: `<w:kern w:val="${run.kernHalfPoints}"/>` })
  if (run.positionHalfPoints)
    out.push({ name: 'w:position', xml: `<w:position w:val="${run.positionHalfPoints}"/>` })
  if (run.sizeHalfPoints) {
    out.push({ name: 'w:sz', xml: `<w:sz w:val="${run.sizeHalfPoints}"/>` })
    out.push({ name: 'w:szCs', xml: `<w:szCs w:val="${run.sizeHalfPoints}"/>` })
  }
  if (run.highlight)
    out.push({ name: 'w:highlight', xml: `<w:highlight w:val="${escapeXmlAttr(run.highlight)}"/>` })
  if (run.underline) {
    const color =
      run.underlineColor && run.underlineColor !== 'auto'
        ? ` w:color="${escapeXmlAttr(run.underlineColor)}"`
        : ''
    out.push({
      name: 'w:u',
      xml: `<w:u w:val="${escapeXmlAttr(run.underlineStyle ?? 'single')}"${color}/>`,
    })
  }
  if (run.shading) {
    out.push({
      name: 'w:shd',
      xml: `<w:shd w:val="clear" w:color="auto" w:fill="${escapeXmlAttr(run.shading)}"/>`,
    })
  }
  if (run.vertAlign)
    out.push({ name: 'w:vertAlign', xml: `<w:vertAlign w:val="${run.vertAlign}"/>` })
  if (run.rtl) out.push({ name: 'w:rtl', xml: '<w:rtl/>' })
  if (run.em && run.em !== 'none') out.push({ name: 'w:em', xml: `<w:em w:val="${run.em}"/>` })
  // the effects the editor set on the run; a cleared one writes nothing and its element goes
  if (run.fxOwn?.glow) out.push({ name: 'w14:glow', xml: w14GlowXml(run.fxOwn.glow) })
  if (run.fxOwn?.shadow) out.push({ name: 'w14:shadow', xml: w14ShadowXml(run.fxOwn.shadow) })
  if (run.fxOwn?.reflection)
    out.push({ name: 'w14:reflection', xml: w14ReflectionXml(run.fxOwn.reflection) })
  if (run.fxOwn?.outline)
    out.push({ name: 'w14:textOutline', xml: w14TextOutlineXml(run.fxOwn.outline) })
  if (run.fxOwn?.fill && run.color)
    out.push({ name: 'w14:textFill', xml: w14TextFillXml(run.color) })
  const otf = run.otf
  if (otf?.ligatures && otf.ligatures !== 'none')
    out.push({
      name: 'w14:ligatures',
      xml: `<w14:ligatures w14:val="${escapeXmlAttr(otf.ligatures)}"/>`,
    })
  if (otf?.numForm && otf.numForm !== 'default')
    out.push({ name: 'w14:numForm', xml: `<w14:numForm w14:val="${otf.numForm}"/>` })
  if (otf?.numSpacing && otf.numSpacing !== 'default')
    out.push({ name: 'w14:numSpacing', xml: `<w14:numSpacing w14:val="${otf.numSpacing}"/>` })
  if (otf?.stylisticSets?.length)
    out.push({
      name: 'w14:stylisticSets',
      xml: `<w14:stylisticSets>${otf.stylisticSets
        .map((id) => `<w14:styleSet w14:id="${Math.round(id)}"/>`)
        .join('')}</w14:stylisticSets>`,
    })
  if (otf?.cntxtAlts) out.push({ name: 'w14:cntxtAlts', xml: '<w14:cntxtAlts/>' })
  const rPrChange = revisionRPrChangeXml(run)
  if (rPrChange) out.push({ name: 'w:rPrChange', xml: rPrChange })
  return out
}

/**
 * Merge the raw rPr slice with the run model: groups whose model value matches the raw
 * encoding keep their original bytes (double underline/themeColor/all four rFonts slots
 * do not degrade); mismatched (edited) groups are rebuilt from the model; children the
 * model does not cover (bdr/em/lang/effects…) are always kept.
 */
export function mergeRPrModel(rawRPr: string, run: Run, insideLink: boolean): string {
  const open = /^<w:rPr(?: [^>]*)?>/.exec(rawRPr)?.[0]
  const fresh = modelRPrChildren(run, insideLink)
  if (!open) {
    // '<w:rPr/>' or unrecognizable: rebuild from the model
    return fresh.length > 0 ? `<w:rPr>${fresh.map((c) => c.xml).join('')}</w:rPr>` : ''
  }
  const inner = rawRPr.slice(open.length, rawRPr.length - '</w:rPr>'.length)
  const rawChildren = splitXmlChildren(inner)
  const rawOf = (tag: string) => rawChildren.find((c) => c.name === tag)?.xml
  // rtl runs decode bold/italic/size from the Cs twins on the parse side;
  // compare against the same elements or every untouched rtl run would get "rebuilt".
  // run.cs can be dropped by round-trips that rebuild the Run (editor marks), so an
  // explicit raw w:rtl re-selects the Cs set on its own.
  const cs = !!run.cs || rawBool(rawOf('w:rtl'))

  // Compare model vs raw-encoded values group by group; equal → keep the original bytes
  // (drop the group from fresh)
  const rebuiltTags = new Set<string>()
  const freshByGroup = new Map<string, PPrChild[]>()
  for (const g of RUN_MANAGED_GROUPS) freshByGroup.set(g.key, [])
  for (const f of fresh) {
    const g = RUN_MANAGED_GROUPS.find((grp) => grp.tags.includes(f.name))!
    freshByGroup.get(g.key)!.push(f)
  }
  const groupEqual = (key: string): boolean => {
    switch (key) {
      case 'rStyle': {
        const raw = rawAttr(rawOf('w:rStyle'), 'w:val')
        if (run.styleId) return raw === run.styleId
        // a run that already sat in the document's hyperlink (rId) stays unstyled if it was unstyled
        return raw === undefined
          ? !insideLink || !!run.link?.rId || !!run.link?.plain
          : raw === 'Hyperlink'
      }
      case 'rFonts': {
        // mirrors the parse side: primary = eastAsia ?? ascii ?? hAnsi, latin = ascii ?? hAnsi,
        // complex-script = literal w:cs
        const attrs = rawOf('w:rFonts')
        const ascii = rawAttr(attrs, 'w:ascii') ?? rawAttr(attrs, 'w:hAnsi')
        // an absent model fontCs is "untouched", not a removal: editor-rebuilt
        // Runs drop unmodeled fields, and losing w:cs on an unrelated edit
        // would corrupt complex-script runs (mergeRFontsXml likewise only
        // writes w:cs when the model carries one)
        // a model value resolved from a theme ref counts as untouched: the literal
        // attrs can never equal it, and rebuilding would materialize the theme link
        return (
          ((rawAttr(attrs, 'w:eastAsia') ?? ascii) === run.font ||
            (run.font !== undefined && run.font === run.themeRFonts?.font)) &&
          (ascii === run.fontAscii ||
            (run.fontAscii !== undefined && run.fontAscii === run.themeRFonts?.fontAscii)) &&
          (run.eastAsiaFont === undefined ||
            rawAttr(attrs, 'w:eastAsia') === run.eastAsiaFont ||
            run.eastAsiaFont === run.themeRFonts?.font) &&
          (run.fontCs === undefined || rawAttr(attrs, 'w:cs') === run.fontCs)
        )
      }
      case 'bold':
        return rawBool(rawOf(cs ? 'w:bCs' : 'w:b')) === !!run.bold
      case 'italic':
        return rawBool(rawOf(cs ? 'w:iCs' : 'w:i')) === !!run.italic
      case 'strike':
        return rawBool(rawOf('w:strike')) === !!run.strike
      case 'caps':
        return rawCaps(rawOf('w:caps'), rawOf('w:smallCaps')) === run.caps
      case 'dstrike':
        return rawBool(rawOf('w:dstrike')) === !!run.dstrike
      // `vanish` may be style-inherited; only the run's own value is compared
      case 'vanish':
        return run.vanishOwn === undefined || rawOnOff(rawOf('w:vanish')) === run.vanishOwn
      case 'spacing':
        return rawInt(rawOf('w:spacing')) === run.charSpacingTwips
      case 'scale': {
        const raw = rawInt(rawOf('w:w'))
        return (raw && raw !== 100 ? raw : undefined) === run.charScalePct
      }
      case 'kern': {
        const raw = rawAttr(rawOf('w:kern'), 'w:val')
        return (raw === undefined ? undefined : parseInt(raw, 10) || 0) === run.kernHalfPoints
      }
      case 'position':
        return (rawInt(rawOf('w:position')) || undefined) === run.positionHalfPoints
      case 'color': {
        const raw = rawAttr(rawOf('w:color'), 'w:val')
        // a theme-resolved model value never equals the cached literal; rebuilding
        // would materialize the theme link (same rule as rFonts above)
        return (
          raw === run.color ||
          (run.themeColor !== undefined && run.color === run.themeColor) ||
          // the colour was read from the fill: the file has no w:color to compare
          (raw === undefined &&
            !!run.color &&
            rawSolidFill(rawOf('w14:textFill')) === run.color.toUpperCase())
        )
      }
      case 'size': {
        const raw = rawAttr(rawOf(cs ? 'w:szCs' : 'w:sz'), 'w:val')
        return (raw ? parseInt(raw, 10) || undefined : undefined) === run.sizeHalfPoints
      }
      case 'highlight': {
        const raw = rawAttr(rawOf('w:highlight'), 'w:val')
        return (raw === 'none' ? undefined : raw) === run.highlight
      }
      case 'shading': {
        const raw = rawAttr(rawOf('w:shd'), 'w:fill')
        return (raw && raw !== 'auto' ? raw : undefined) === run.shading
      }
      // undefined is "not touched": what the file says about these stays, switches the model
      // does not carry (brackets, compression) included
      case 'fitText':
        return (
          run.fitTextTwips === undefined || (rawInt(rawOf('w:fitText')) || 0) === run.fitTextTwips
        )
      case 'eaLayout': {
        if (run.eaLayout === undefined) return true
        const raw = rawOf('w:eastAsianLayout')
        const on = (attr: string) => /^(1|true|on)$/.test(rawAttr(raw, attr) ?? '')
        return on('w:combine') === !!run.eaLayout.combine && on('w:vert') === !!run.eaLayout.vert
      }
      case 'lang': {
        if (run.lang === undefined) return true
        const raw = rawOf('w:lang')
        return (['val', 'eastAsia', 'bidi'] as const).every(
          (slot) => (rawAttr(raw, `w:${slot}`) || undefined) === (run.lang?.[slot] || undefined),
        )
      }
      case 'noProof':
        return run.noProof === undefined || rawBool(rawOf('w:noProof')) === run.noProof
      // `bdr` may be style-inherited; only the run's own border is compared
      case 'bdr':
        return run.bdrOwn === undefined || sameRunBorder(rawRunBorder(rawOf('w:bdr')), run.bdrOwn)
      case 'underline': {
        // w:u is not a boolean prop: no w:val (or val="none") means no
        // underline, regardless of other attributes like w:color. Must match
        // underlineProp() used at parse time.
        const val = rawAttr(rawOf('w:u'), 'w:val')
        if ((val !== undefined && val !== 'none') !== !!run.underline) return false
        if (!run.underline) return true
        // a model without a style or colour never asked for another line: the file's own stays
        return (
          (run.underlineStyle === undefined || val === run.underlineStyle) &&
          (run.underlineColor === undefined ||
            (rawAttr(rawOf('w:u'), 'w:color') ?? 'auto') === run.underlineColor)
        )
      }
      case 'vertAlign': {
        const raw = rawAttr(rawOf('w:vertAlign'), 'w:val')
        const modeled = raw === 'superscript' || raw === 'subscript' ? raw : undefined
        return modeled === run.vertAlign
      }
      case 'rtl':
        return rawBool(rawOf('w:rtl')) === !!run.rtl
      // undefined is "not touched": what the file says stays; the default value removes it
      case 'em':
        return run.em === undefined || (rawAttr(rawOf('w:em'), 'w:val') ?? 'none') === run.em
      case 'otLigatures':
        return (
          run.otf?.ligatures === undefined ||
          (rawAttr(rawOf('w14:ligatures'), 'w14:val') ?? 'none') === run.otf.ligatures
        )
      case 'otNumForm':
        return (
          run.otf?.numForm === undefined ||
          (rawAttr(rawOf('w14:numForm'), 'w14:val') ?? 'default') === run.otf.numForm
        )
      case 'otNumSpacing':
        return (
          run.otf?.numSpacing === undefined ||
          (rawAttr(rawOf('w14:numSpacing'), 'w14:val') ?? 'default') === run.otf.numSpacing
        )
      case 'otSets':
        return (
          run.otf?.stylisticSets === undefined ||
          rawStyleSets(rawOf('w14:stylisticSets')).join() === run.otf.stylisticSets.join()
        )
      case 'otAlts': {
        if (run.otf?.cntxtAlts === undefined) return true
        const raw = rawOf('w14:cntxtAlts')
        return (
          (!!raw && !/^(0|false|off)$/.test(rawAttr(raw, 'w14:val') ?? '')) === run.otf.cntxtAlts
        )
      }
      // undefined is "not touched in the editor": what the file says stays
      case 'fxGlow':
        return run.fxOwn?.glow === undefined
      case 'fxShadow':
        return run.fxOwn?.shadow === undefined
      case 'fxReflection':
        return run.fxOwn?.reflection === undefined
      case 'fxOutline':
        return run.fxOwn?.outline === undefined
      // Word shows the fill, not w:color: a solid fill of the file follows a change of colour
      case 'fxFill': {
        if (run.fxOwn?.fill !== undefined) return false
        const rawFill = rawSolidFill(rawOf('w14:textFill'))
        if (rawFill === undefined || groupEqual('color')) return true
        return rawFill === '' && rawOf('w:color') === undefined
      }
      case 'rPrChange':
        return !!rawOf('w:rPrChange') === !!run.rPrChange
      default:
        return true
    }
  }
  const freshOut: PPrChild[] = []
  for (const g of RUN_MANAGED_GROUPS) {
    if (groupEqual(g.key)) continue
    for (const t of g.tags) rebuiltTags.add(t)
    const rawRFonts = g.key === 'rFonts' ? rawOf('w:rFonts') : undefined
    if (rawRFonts && (run.font || run.fontAscii)) {
      // edited font slots merge into the original element instead of replacing it
      freshOut.push({ name: 'w:rFonts', xml: mergeRFontsXml(rawRFonts, run) })
    } else if (g.key === 'fxFill' && run.fxOwn?.fill === undefined) {
      if (run.color) freshOut.push({ name: 'w14:textFill', xml: w14TextFillXml(run.color) })
    } else {
      freshOut.push(...freshByGroup.get(g.key)!)
    }
  }

  const kept = rawChildren.filter((c) => !rebuiltTags.has(c.name))
  const rank = (n: string) => RPR_CHILD_ORDER.indexOf(n)
  const parts: string[] = []
  let fi = 0
  let prevRank = -1
  for (const child of kept) {
    const own = rank(child.name)
    const effective = own === -1 ? prevRank : Math.max(own, prevRank)
    while (fi < freshOut.length && rank(freshOut[fi].name) < effective)
      parts.push(freshOut[fi++].xml)
    parts.push(child.xml)
    prevRank = effective
  }
  while (fi < freshOut.length) parts.push(freshOut[fi++].xml)
  if (parts.length === 0) return ''
  return `${open}${parts.join('')}</w:rPr>`
}

// ---- new drawing paragraph generators ----

/**
 * Build a minimal <w:p> fragment that contains a floating WPS text-box anchored
 * at the cursor position with the given dimensions.
 *
 * @param widthEmu  horizontal size in EMU  (default ~5 cm = 1800000)
 * @param heightEmu vertical size in EMU    (default ~3 cm = 1080000)
 * @param id        wp:docPr id / name suffix (caller must keep unique)
 * @param fillHex   6-char solid fill hex colour (default "FFFFFF" = white)
 * @param borderHex 6-char border hex colour     (default "000000" = black)
 */
export function buildTextboxParagraphXml(opts?: {
  widthEmu?: number
  heightEmu?: number
  id?: number
  fillHex?: string
  borderHex?: string
  /** the text runs down the box (Word's vertical text box is `eaVert`) */
  vert?: TextboxVert
}): string {
  const widthEmu = opts?.widthEmu ?? 1800000
  const heightEmu = opts?.heightEmu ?? 1080000
  const id = opts?.id ?? 1
  const fillHex = opts?.fillHex ?? 'FFFFFF'
  const borderHex = opts?.borderHex ?? '000000'

  const spPr =
    `<wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill>` +
    `<a:ln><a:solidFill><a:srgbClr val="${borderHex}"/></a:solidFill></a:ln>` +
    `</wps:spPr>`

  const wsp =
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:cNvSpPr txBox="1"/>` +
    spPr +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    (opts?.vert ? `<wps:bodyPr vert="${opts.vert}"/>` : `<wps:bodyPr/>`) +
    `</wps:wsp>`

  const graphicData =
    `<a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${wsp}</a:graphicData>`

  const graphic = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${graphicData}</a:graphic>`

  const anchor =
    `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapSquare wrapText="bothSides"/>` +
    `<wp:docPr id="${id}" name="TextBox ${id}"/>` +
    graphic +
    `</wp:anchor>`

  // Requires="wps" must resolve at the AlternateContent scope, or Word
  // reports the whole file as unreadable content
  const mcNs =
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
  const mcChoice =
    `<mc:Choice Requires="wps">` + `<w:drawing>${anchor}</w:drawing>` + `</mc:Choice>`

  const vmlRect =
    `<v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute;width:${widthEmu / EMU_PER_PT}pt;height:${heightEmu / EMU_PER_PT}pt" filled="t" stroked="t">` +
    `<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p></w:txbxContent></v:textbox>` +
    `</v:rect>`

  const mcFallback = `<mc:Fallback>` + `<w:pict>${vmlRect}</w:pict>` + `</mc:Fallback>`

  const alternateContent =
    `<mc:AlternateContent ${mcNs}>` + mcChoice + mcFallback + `</mc:AlternateContent>`

  return `<w:p><w:r>${alternateContent}</w:r></w:p>`
}

/** one paragraph of anchored-textbox content: real runs plus ParaFormat-owned pPr */
export interface TextboxContentParagraph {
  runs: Run[]
  format?: ParaFormat
}

export interface AnchoredTextboxOptions {
  /**
   * 'paragraph': the box rides its holder paragraph through the flow (card
   * regions); 'page': absolute page coordinates (newsletter regions).
   */
  anchor: 'paragraph' | 'page'
  /**
   * EMU offsets. paragraph anchor: x from the text column, y from the holder
   * paragraph; page anchor: both from the page edges.
   */
  xEmu: number
  yEmu: number
  widthEmu: number
  heightEmu: number
  paragraphs: TextboxContentParagraph[]
  /** wp:docPr id / name suffix (caller must keep unique) */
  id?: number
  /** solid interior fill (RRGGBB); omitted = transparent */
  fillHex?: string
  /** border color (RRGGBB); omitted = no outline */
  borderHex?: string
  /** text insets in EMU; omitted = Word defaults (91440 horizontal / 45720 vertical) */
  insetsEmu?: { l: number; t: number; r: number; b: number }
  /** corner radius in EMU; maps to a:prstGeom roundRect with the matching adj */
  cornerRadiusEmu?: number
  /** written as relativeHeight base + zOrder (P16 z discipline) */
  zOrder?: number
  behindDoc?: boolean
  /** default: paragraph anchor wraps topAndBottom, page anchor wraps none */
  wrap?: 'topAndBottom' | 'none'
  /** exact line height (twips) for the holder <w:p>, so the anchor row adds no visible height */
  holderLineTwips?: number
  /** spacing-before (twips) on the holder <w:p> — the region's slot in the caller's spacing chain */
  holderSpacingBeforeTwips?: number
  /** page-break-before on the holder <w:p> (region starts its page) */
  holderPageBreakBefore?: boolean
}

/**
 * Build a <w:p> holding a floating WPS text box whose content is real,
 * editable w:p paragraphs (P20 region container). Unlike
 * buildTextboxParagraphXml (cursor-position UI insert), this variant takes
 * explicit geometry, fill, insets and z-order — pdf2docx uses it to rebuild
 * measured card/column regions.
 */
export function buildAnchoredTextboxParagraphXml(opts: AnchoredTextboxOptions): string {
  const id = opts.id ?? 1
  const { xEmu, yEmu, widthEmu, heightEmu } = opts

  const contentParaXml = (p: TextboxContentParagraph): string => {
    const children = formatPPrChildren(p.format)
    children.sort((a, b) => PPR_CHILD_ORDER.indexOf(a.name) - PPR_CHILD_ORDER.indexOf(b.name))
    const pPr = children.length > 0 ? `<w:pPr>${children.map((c) => c.xml).join('')}</w:pPr>` : ''
    return `<w:p>${pPr}${inlineRunsXml(p.runs)}</w:p>`
  }
  const txbxContent = `<w:txbxContent>${opts.paragraphs.map(contentParaXml).join('')}</w:txbxContent>`

  // roundRect radius = adj / 100000 * min(cx, cy); clamp to the spec's half-side max
  const geom = opts.cornerRadiusEmu
    ? `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${Math.min(
        50000,
        Math.max(
          0,
          Math.round((opts.cornerRadiusEmu / Math.max(1, Math.min(widthEmu, heightEmu))) * 100000),
        ),
      )}"/></a:avLst></a:prstGeom>`
    : `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
  const fill = opts.fillHex
    ? `<a:solidFill><a:srgbClr val="${escapeXmlAttr(opts.fillHex)}"/></a:solidFill>`
    : `<a:noFill/>`
  const line = opts.borderHex
    ? `<a:ln><a:solidFill><a:srgbClr val="${escapeXmlAttr(opts.borderHex)}"/></a:solidFill></a:ln>`
    : `<a:ln><a:noFill/></a:ln>`
  const spPr =
    `<wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    geom +
    fill +
    line +
    `</wps:spPr>`

  const ins = opts.insetsEmu ?? { l: 91440, t: 45720, r: 91440, b: 45720 }
  // fixed box: a:noAutofit keeps overflowing content from growing the region
  const bodyPr =
    `<wps:bodyPr wrap="square" lIns="${Math.round(ins.l)}" tIns="${Math.round(ins.t)}" ` +
    `rIns="${Math.round(ins.r)}" bIns="${Math.round(ins.b)}" anchor="t"><a:noAutofit/></wps:bodyPr>`

  const wsp =
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:cNvSpPr txBox="1"/>` +
    spPr +
    `<wps:txbx>${txbxContent}</wps:txbx>` +
    bodyPr +
    `</wps:wsp>`

  const graphic =
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${wsp}</a:graphicData>` +
    `</a:graphic>`

  const relFrom = opts.anchor === 'page' ? 'page' : undefined
  const posH =
    `<wp:positionH relativeFrom="${relFrom ?? 'column'}">` +
    `<wp:posOffset>${Math.round(xEmu)}</wp:posOffset></wp:positionH>`
  const posV =
    `<wp:positionV relativeFrom="${relFrom ?? 'paragraph'}">` +
    `<wp:posOffset>${Math.round(yEmu)}</wp:posOffset></wp:positionV>`
  const wrapMode = opts.wrap ?? (opts.anchor === 'page' ? 'none' : 'topAndBottom')
  const wrapEl = wrapMode === 'topAndBottom' ? '<wp:wrapTopAndBottom/>' : '<wp:wrapNone/>'

  const anchor =
    `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="${251658240 + (opts.zOrder ?? 0)}" behindDoc="${opts.behindDoc ? 1 : 0}" ` +
    `locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    posH +
    posV +
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    wrapEl +
    `<wp:docPr id="${id}" name="Region ${id}"/>` +
    graphic +
    `</wp:anchor>`

  // Requires="wps" must resolve at the AlternateContent scope, or Word
  // reports the whole file as unreadable content
  const mcNs =
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
  const vmlTag = opts.cornerRadiusEmu ? 'v:roundrect' : 'v:rect'
  const vmlFill = opts.fillHex
    ? ` fillcolor="#${escapeXmlAttr(opts.fillHex)}" filled="t"`
    : ' filled="f"'
  const vmlRect =
    `<${vmlTag} xmlns:v="urn:schemas-microsoft-com:vml" ` +
    `style="position:absolute;margin-left:${xEmu / EMU_PER_PT}pt;margin-top:${yEmu / EMU_PER_PT}pt;` +
    `width:${widthEmu / EMU_PER_PT}pt;height:${heightEmu / EMU_PER_PT}pt"${vmlFill} stroked="f">` +
    `<v:textbox>${txbxContent}</v:textbox>` +
    `</${vmlTag}>`

  const alternateContent =
    `<mc:AlternateContent ${mcNs}>` +
    `<mc:Choice Requires="wps"><w:drawing>${anchor}</w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict>${vmlRect}</w:pict></mc:Fallback>` +
    `</mc:AlternateContent>`

  const holderSpacing =
    opts.holderLineTwips !== undefined || opts.holderSpacingBeforeTwips !== undefined
      ? `<w:spacing` +
        (opts.holderSpacingBeforeTwips
          ? ` w:before="${Math.round(opts.holderSpacingBeforeTwips)}"`
          : '') +
        ` w:after="0"` +
        (opts.holderLineTwips !== undefined
          ? ` w:line="${Math.round(opts.holderLineTwips)}" w:lineRule="exact"`
          : '') +
        `/>`
      : ''
  const holderBreak = opts.holderPageBreakBefore ? '<w:pageBreakBefore/>' : ''
  const holderPPr =
    holderSpacing || holderBreak ? `<w:pPr>${holderBreak}${holderSpacing}</w:pPr>` : ''
  return `<w:p>${holderPPr}<w:r>${alternateContent}</w:r></w:p>`
}

/**
 * Build a <w:p> fragment for a floating WPS shape (prstGeom) with optional
 * text content. Same anchor structure as buildTextboxParagraphXml.
 */
export function buildShapeParagraphXml(opts: {
  prst: string
  widthEmu?: number
  heightEmu?: number
  id?: number
  fillHex?: string
  borderHex?: string
  withTextbox?: boolean
}): string {
  const widthEmu = opts.widthEmu ?? 1800000
  const heightEmu = opts.heightEmu ?? 1080000
  const id = opts.id ?? 1
  const fillHex = opts.fillHex ?? '4472C4'
  const borderHex = opts.borderHex ?? '2F5496'

  const spPr =
    `<wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    `<a:prstGeom prst="${opts.prst}"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill>` +
    `<a:ln><a:solidFill><a:srgbClr val="${borderHex}"/></a:solidFill></a:ln>` +
    `</wps:spPr>`

  // Word centers autoshape text both ways (a survey of 279 Word-authored documents
  // puts anchor="ctr" on 72% and w:jc="center" on 74% of them). A text box is the
  // opposite case and stays top-left, which is why buildTextboxParagraphXml differs.
  const seededPara = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p>`
  const txbxPart = opts.withTextbox
    ? `<wps:txbx><w:txbxContent>${seededPara}</w:txbxContent></wps:txbx>`
    : ''

  // Gallery shapes carry a style block, and its a:fontRef is what gives them light
  // text on the accent fill — Word writes no color on the runs. The fill/line refs
  // are inert here because spPr states both explicitly, but CT_ShapeStyle requires
  // all four children in this order.
  const style =
    `<wps:style>` +
    `<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>` +
    `<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>` +
    `<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>` +
    `<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>` +
    `</wps:style>`

  const wsp =
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:cNvPr id="${id}" name="${opts.prst} ${id}"/>` +
    `<wps:cNvSpPr/>` +
    spPr +
    style +
    txbxPart +
    `<wps:bodyPr anchor="ctr"/>` +
    `</wps:wsp>`

  const graphicData =
    `<a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${wsp}</a:graphicData>`

  const graphic = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${graphicData}</a:graphic>`

  const anchor =
    `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapSquare wrapText="bothSides"/>` +
    `<wp:docPr id="${id}" name="${opts.prst} ${id}"/>` +
    graphic +
    `</wp:anchor>`

  const mcNs2 =
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
  const mcChoice =
    `<mc:Choice Requires="wps">` + `<w:drawing>${anchor}</w:drawing>` + `</mc:Choice>`

  const vmlRect =
    `<v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute;width:${widthEmu / EMU_PER_PT}pt;height:${heightEmu / EMU_PER_PT}pt" filled="t" stroked="t">` +
    // The VML twin renders in place of the shape on older Word builds, so it has to
    // carry the same centering (v:textbox takes the vertical half as an attribute)
    (opts.withTextbox
      ? `<v:textbox style="v-text-anchor:middle"><w:txbxContent>${seededPara}</w:txbxContent></v:textbox>`
      : '') +
    `</v:rect>`

  const mcFallback = `<mc:Fallback>` + `<w:pict>${vmlRect}</w:pict>` + `</mc:Fallback>`

  const alternateContent =
    `<mc:AlternateContent ${mcNs2}>` + mcChoice + mcFallback + `</mc:AlternateContent>`

  return `<w:p><w:r>${alternateContent}</w:r></w:p>`
}

/**
 * Build a <w:p> fragment containing a floating WordArt WPS text box.
 * The shape has no background fill; the text runs carry large size (36pt)
 * and the specified solid color. Style is approximated (no stroke/effects in
 * the saved run — the caller picks a readable solid color) — Word can open
 * the result. Presets live in the UI layer (@genoffice/ui wordart-presets).
 */
export function buildWordArtParagraphXml(opts: {
  text?: string
  /** 6-digit hex without '#'; defaults to the Office accent blue. */
  colorHex?: string
  italic?: boolean
  widthEmu?: number
  heightEmu?: number
  id?: number
}): string {
  const widthEmu = opts.widthEmu ?? 2700000 // ~7.5 cm
  const heightEmu = opts.heightEmu ?? 720000 // ~2 cm
  const id = opts.id ?? 1
  const colorHex = opts.colorHex ?? '4472C4'
  const text = opts.text ?? 'WordArt'

  // Build run props: large font + color (schema order: b < i < color < sz)
  const rPr =
    `<w:rPr>` +
    `<w:b/>` +
    (opts.italic ? `<w:i/>` : '') +
    `<w:color w:val="${colorHex}"/>` +
    `<w:sz w:val="72"/>` + // 36pt = 72 half-points
    `<w:szCs w:val="72"/>` +
    `</w:rPr>`

  const textRun = `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r>`
  const txbxContent = `<w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${textRun}</w:p></w:txbxContent>`

  const spPr =
    `<wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:noFill/>` +
    `<a:ln><a:noFill/></a:ln>` +
    `</wps:spPr>`

  const wsp =
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    // cNvSpPr is a required wsp child; Word treats its absence as corruption
    `<wps:cNvSpPr txBox="1"/>` +
    spPr +
    `<wps:txbx>${txbxContent}</wps:txbx>` +
    `<wps:bodyPr/>` +
    `</wps:wsp>`

  const graphicData =
    `<a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${wsp}</a:graphicData>`

  const graphic = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${graphicData}</a:graphic>`

  const anchor =
    `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapSquare wrapText="bothSides"/>` +
    `<wp:docPr id="${id}" name="WordArt ${id}"/>` +
    graphic +
    `</wp:anchor>`

  // Requires="wps" must resolve at the AlternateContent scope, or Word
  // reports the whole file as unreadable content
  const mcNs =
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
  const mcChoice =
    `<mc:Choice Requires="wps">` + `<w:drawing>${anchor}</w:drawing>` + `</mc:Choice>`

  const vmlRect =
    `<v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute;width:${widthEmu / EMU_PER_PT}pt;height:${heightEmu / EMU_PER_PT}pt" filled="f" stroked="f">` +
    `<v:textbox><w:txbxContent><w:p><w:r>${rPr}<w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r></w:p></w:txbxContent></v:textbox>` +
    `</v:rect>`

  const mcFallback = `<mc:Fallback><w:pict>${vmlRect}</w:pict></mc:Fallback>`

  const alternateContent =
    `<mc:AlternateContent ${mcNs}>` + mcChoice + mcFallback + `</mc:AlternateContent>`

  return `<w:p><w:r>${alternateContent}</w:r></w:p>`
}

/** the size text has when neither it nor the editor names one (Word's 10.5 pt) */
const DEFAULT_SIZE_HALF_POINTS = 21

/**
 * The EQ field of enclosed or combined characters. One read from the file is
 * written as it was. A new one is written the way Word writes it: the parts
 * overlaid with \o, every run in the run's own formatting, the sized part
 * carrying its size and baseline shift.
 */
function eqLayoutFieldXml(run: Run, insideLink: boolean): string {
  const layout = run.eqLayout!
  if (layout.xml) return layout.xml
  const base: Run = { ...run, eqLayout: undefined, text: '' }
  const rPr = (over: Partial<Run> = {}) => runRPrXml({ ...base, ...over }, insideLink)
  const plain = rPr()
  const code = (text: string, props = plain) =>
    `<w:r>${props}<w:instrText xml:space="preserve">${escapeXmlText(text)}</w:instrText></w:r>`
  const size = run.sizeHalfPoints ?? DEFAULT_SIZE_HALF_POINTS
  let body: string
  if (layout.kind === 'enclose') {
    const enclosure =
      layout.style === 'enlarge'
        ? rPr({
            sizeHalfPoints: Math.round(size * 1.5),
            positionHalfPoints: -Math.round(size * 0.19),
          })
        : plain
    const inside =
      layout.style === 'shrink'
        ? rPr({
            sizeHalfPoints: Math.max(2, Math.round((size * 2) / 3)),
            positionHalfPoints: Math.round(size * 0.1),
          })
        : plain
    body =
      code(' eq \\o\\ac(') +
      code(layout.a, enclosure) +
      code(',') +
      code(layout.b, inside) +
      code(')')
  } else {
    // two lines of half-size text, one raised and one lowered about the baseline (points)
    const small = rPr({ sizeHalfPoints: Math.max(2, Math.round(size / 2)) })
    const up = Math.max(1, Math.round(size / 4))
    const down = Math.max(1, Math.round(size / 10))
    body =
      code(` eq \\o(\\s\\up ${up}(`) +
      code(layout.a, small) +
      code(`),\\s\\do ${down}(`) +
      code(layout.b, small) +
      code('))')
  }
  return (
    `<w:r>${plain}<w:fldChar w:fldCharType="begin"/></w:r>` +
    body +
    `<w:r>${plain}<w:fldChar w:fldCharType="end"/></w:r>`
  )
}

function runRPrXml(run: Run, insideLink: boolean): string {
  // OOXML requires rPr children in schema order:
  // rStyle < rFonts < b < i < strike < color < sz < highlight < u < vertAlign
  if (run.rawRPr !== undefined) return mergeRPrModel(run.rawRPr, run, insideLink)
  const props = modelRPrChildren(run, insideLink).map((c) => c.xml)
  return props.length > 0 ? `<w:rPr>${props.join('')}</w:rPr>` : ''
}

function generateRunXml(run: Run, insideLink: boolean): string {
  const rPr = runRPrXml(run, insideLink)

  // Translate embedded control characters back to OOXML elements.
  // Deleted runs carry their text in w:delText instead of w:t; a moved-away run keeps w:t, as Word writes it.
  const textTag = run.del && !run.del.move ? 'w:delText' : 'w:t'
  const segments: string[] = []
  let buffer = ''
  const flush = () => {
    if (buffer !== '') {
      segments.push(`<${textTag} xml:space="preserve">${escapeXmlText(buffer)}</${textTag}>`)
      buffer = ''
    }
  }
  for (const ch of run.text) {
    if (ch === '\t') {
      flush()
      segments.push('<w:tab/>')
    } else if (ch === '\n') {
      flush()
      segments.push('<w:br/>')
    } else if (ch === '\f') {
      flush()
      segments.push('<w:br w:type="page"/>')
    } else if (ch === '\v') {
      flush()
      segments.push('<w:br w:type="column"/>')
    } else if (ch === '\u001e') {
      flush()
      segments.push('<w:br w:type="textWrapping" w:clear="all"/>')
    } else if (ch === '\u2011') {
      flush()
      segments.push('<w:noBreakHyphen/>')
    } else if (ch === '\u00ad') {
      flush()
      segments.push('<w:softHyphen/>')
    } else {
      buffer += ch
    }
  }
  flush()
  return `<w:r>${rPr}${segments.join('')}</w:r>`
}
