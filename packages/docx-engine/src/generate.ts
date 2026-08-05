import type {
  GeneratedBlock,
  ImageWrap,
  ParaFormat,
  Run,
  TabStop,
  TableCell,
  TableModel,
} from './types'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

export interface GenerateContext {
  /** heading level -> styleId existing in the original styles.xml */
  headingStyleIds: Map<number, string>
  /** styleId for list paragraphs, if present in the original doc */
  listParagraphStyleId?: string
  /** allocate a new relationship id for a hyperlink target; returns rId */
  allocateHyperlinkRel: (href: string) => string
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
}

/**
 * Rewrite the display size (wp:extent + a:xfrm a:ext) and/or paragraph
 * alignment (w:jc) of an image paragraph, leaving everything else untouched.
 */
export function patchImageParagraphXml(xml: string, patch: ImagePatch): string {
  let out = xml
  if (patch.widthPx && patch.heightPx) {
    const cx = Math.max(1, Math.round(patch.widthPx * EMU_PER_PX))
    const cy = Math.max(1, Math.round(patch.heightPx * EMU_PER_PX))
    const resize = (tag: string) =>
      tag.replace(/cx="\d+"/, `cx="${cx}"`).replace(/cy="\d+"/, `cy="${cy}"`)
    out = out.replace(/<wp:extent[^>]*\/?>/, resize)
    out = out.replace(/<a:ext[^>]*\/>/, resize)
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
  // Rewrite posOffset values inside positionH / positionV (surgical)
  if (patch.posOffsetX !== undefined) {
    out = out.replace(
      /(<wp:positionH[^>]*>[\s\S]*?)<wp:posOffset>-?\d+<\/wp:posOffset>([\s\S]*?<\/wp:positionH>)/,
      `$1<wp:posOffset>${Math.round(patch.posOffsetX)}</wp:posOffset>$2`,
    )
  }
  if (patch.posOffsetY !== undefined) {
    out = out.replace(
      /(<wp:positionV[^>]*>[\s\S]*?)<wp:posOffset>-?\d+<\/wp:posOffset>([\s\S]*?<\/wp:positionV>)/,
      `$1<wp:posOffset>${Math.round(patch.posOffsetY)}</wp:posOffset>$2`,
    )
  }
  return out
}

const WRAP_ELEMENT_RE =
  /<wp:wrapNone\s*\/>|<wp:wrapSquare[^>]*\/>|<wp:wrapSquare[\s\S]*?<\/wp:wrapSquare>|<wp:wrapTight[\s\S]*?<\/wp:wrapTight>|<wp:wrapThrough[\s\S]*?<\/wp:wrapThrough>|<wp:wrapTopAndBottom\s*\/>|<wp:wrapTopAndBottom[\s\S]*?<\/wp:wrapTopAndBottom>/g

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
  posOffset?: { x: number; y: number },
  marginAlign?: { h: 'left' | 'center' | 'right'; v: 'top' | 'center' | 'bottom' },
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
    position =
      '<wp:simplePos x="0" y="0"/>' +
      `<wp:positionH relativeFrom="column"><wp:posOffset>${Math.round(posOffset.x)}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${Math.round(posOffset.y)}</wp:posOffset></wp:positionV>`
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
    ` relativeHeight="251658240" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1">`
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
    .replace(/&amp;/g, '&')
}

function textNodes(xml: string, tag: 'w:t' | 'm:t'): XmlTextNode[] {
  const nodes: XmlTextNode[] = []
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
    const replacement = node.open + escapeXmlText(text) + node.close
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
  const nodes = textNodes(xml, 'w:t')
  if (nodes.length === 0) return xml
  const tabIndex = xml.search(/<w:tab(?:\s[^>]*)?\/>/)
  if (patch.right !== undefined && tabIndex === -1) return xml
  const leftNodes = tabIndex === -1 ? nodes : nodes.filter((node) => node.start < tabIndex)
  const rightNodes = tabIndex === -1 ? [] : nodes.filter((node) => node.start > tabIndex)
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
function xmlSegments(
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
    if (c === -1 || c >= to) break
    if (o !== -1 && o < c) {
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
 * Rebuild one cell's paragraphs with new text, keeping tcPr and the first
 * paragraph's pPr / first run's rPr (so header bold, alignment, shading stay).
 * Returns null when the cell is too complex to patch safely (nested table).
 */
function patchCellXml(tcXml: string, paras: string[]): string | null {
  if (tcXml.indexOf('<w:tbl', 1) !== -1) return null
  const openTag = /^<w:tc(?: [^>]*)?>/.exec(tcXml)?.[0]
  if (!openTag) return null
  const tcPr = /<w:tcPr[\s\S]*?<\/w:tcPr>|<w:tcPr[^>]*\/>/.exec(tcXml)?.[0] ?? ''
  const firstP = /<w:p(?: [^>]*)?>[\s\S]*?<\/w:p>/.exec(tcXml)?.[0] ?? ''
  const pPr = /<w:pPr[\s\S]*?<\/w:pPr>|<w:pPr[^>]*\/>/.exec(firstP)?.[0] ?? ''
  const firstRun = /<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/.exec(firstP)?.[0] ?? ''
  const rPr = /<w:rPr[\s\S]*?<\/w:rPr>/.exec(firstRun)?.[0] ?? ''
  const body = paras
    .map((t) =>
      t === ''
        ? `<w:p>${pPr}</w:p>`
        : `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlText(t)}</w:t></w:r></w:p>`,
    )
    .join('')
  return openTag + tcPr + body + '</w:tc>'
}

/** per-cell patch: plain paragraph strings, or nested-table cell texts by nested index */
export type CellTextsPatch =
  | readonly string[]
  | {
      /** this cell's own text (optional; rewriting the outer text of a cell containing a nested table is not supported yet) */
      paras?: readonly string[] | null
      /** one cell-text grid per direct nested table (null = leave that nested table untouched) */
      nested?: ReadonlyArray<ReadonlyArray<
        ReadonlyArray<readonly string[] | null | undefined> | null | undefined
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
        ? patchCellXml(tcXml, entry as string[])
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
        ReadonlyArray<readonly string[] | null | undefined> | null | undefined
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
}

/** plain text of one w:p fragment (w:t + tabs/breaks), for change detection */
function paraPlainText(pXml: string): string {
  let out = ''
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\/>|<w:br\/>|<w:cr\/>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(pXml)) !== null) {
    if (m[0] === '<w:tab/>') out += '\t'
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

/** set or remove w:jc inside a pPr fragment (creating pPr when needed) */
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
    let pPr = /<w:pPr[\s\S]*?<\/w:pPr>|<w:pPr[^>]*\/>/.exec(template)?.[0] ?? ''
    if (patch.align !== undefined) pPr = pPrWithJc(pPr, patch.align)
    // no rel allocation inside textboxes: only links that already have an rId
    // survive re-generation, new links degrade to plain text
    parts.push(`<w:p>${pPr}${runsXml(patch.runs, null)}</w:p>`)
  }
  return open + parts.join('') + '</w:txbxContent>'
}

/**
 * Replace textbox paragraphs inside a paragraph fragment that carries
 * anchored DrawingML textboxes. `boxes[box]` = per-paragraph patches for that
 * box (visible-shape order, matching Block.textboxes), or null/undefined to
 * leave the box untouched. Word pairs every DrawingML shape with a VML twin
 * inside mc:Fallback whose w:txbxContent duplicates the content — fallback
 * copies are patched with the same paragraphs as the preceding visible box so
 * both renderings stay in sync.
 */
export function patchTextboxParas(
  paragraphXml: string,
  boxes: ReadonlyArray<ReadonlyArray<TextboxParaPatch | null | undefined> | null | undefined>,
): string {
  const segs = xmlSegments(paragraphXml, 'w:txbxContent', 0, paragraphXml.length)
  if (segs.length === 0) return paragraphXml

  const fallbacks: Array<{ start: number; end: number }> = []
  const fbRe = /<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g
  let fb: RegExpExecArray | null
  while ((fb = fbRe.exec(paragraphXml)) !== null) {
    fallbacks.push({ start: fb.index, end: fb.index + fb[0].length })
  }
  const inFallback = (pos: number) => fallbacks.some((f) => pos >= f.start && pos < f.end)

  let out = ''
  let cursor = 0
  let boxIndex = -1
  let lastVisibleCounted = false
  for (const seg of segs) {
    const segXml = paragraphXml.slice(seg.start, seg.end)
    if (!inFallback(seg.start)) {
      // display extraction skips all-empty boxes, so only boxes with text
      // consume an index; a fallback twin repeats the preceding visible box
      lastVisibleCounted = paraPlainText(segXml) !== ''
      if (lastVisibleCounted) boxIndex++
    } else if (!lastVisibleCounted) {
      continue
    }
    const paras = boxIndex >= 0 && lastVisibleCounted ? boxes[boxIndex] : null
    if (paras == null) continue
    out += paragraphXml.slice(cursor, seg.start) + patchTxbxContent(segXml, paras)
    cursor = seg.end
  }
  return out + paragraphXml.slice(cursor)
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
    const resized = drawingXml
      .replace(/(<wp:extent\b[^>]*\bcy=")\d+(")/, `$1${cy}$2`)
      .replace(/(<a:ext\b[^>]*\bcy=")\d+(")/, `$1${cy}$2`)
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

/** pPr children the ParaFormat model owns (rebuilt on format edits) */
function formatPPrChildren(format: ParaFormat | undefined): PPrChild[] {
  if (!format) return []
  const out: PPrChild[] = []
  if (format.pageBreakBefore) out.push({ name: 'w:pageBreakBefore', xml: '<w:pageBreakBefore/>' })
  if (format.borders) {
    const line = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="1" w:color="auto"/>`
    const sides: string[] = []
    // schema order inside pBdr: top, left, bottom, right
    if (format.borders.includes('t')) sides.push(line('top'))
    if (format.borders.includes('l')) sides.push(line('left'))
    if (format.borders.includes('b')) sides.push(line('bottom'))
    if (format.borders.includes('r')) sides.push(line('right'))
    if (sides.length > 0) out.push({ name: 'w:pBdr', xml: `<w:pBdr>${sides.join('')}</w:pBdr>` })
  }
  if (format.shadingFill) {
    out.push({
      name: 'w:shd',
      xml: `<w:shd w:val="clear" w:color="auto" w:fill="${escapeXmlAttr(format.shadingFill)}"/>`,
    })
  }
  if (format.bidi) out.push({ name: 'w:bidi', xml: '<w:bidi/>' })
  const spacingAttrs: string[] = []
  if (format.spaceBefore && format.spaceBefore > 0) {
    spacingAttrs.push(`w:before="${Math.round(format.spaceBefore)}"`)
  }
  // explicit 0 overrides the style's space-after in Word, so it must be written back
  if (format.spaceAfter !== undefined && format.spaceAfter >= 0) {
    spacingAttrs.push(`w:after="${Math.round(format.spaceAfter)}"`)
  }
  if ((format.lineRule === 'exact' || format.lineRule === 'atLeast') && format.lineRawTwips) {
    // Exact/at-least line height: write the twips back verbatim (previously only auto was
    // recognized, so edited paragraphs lost their line spacing)
    spacingAttrs.push(
      `w:line="${Math.round(format.lineRawTwips)}"`,
      `w:lineRule="${format.lineRule}"`,
    )
  } else if (format.lineSpacing && format.lineSpacing > 0) {
    spacingAttrs.push(`w:line="${Math.round(format.lineSpacing * 240)}"`, 'w:lineRule="auto"')
  }
  if (spacingAttrs.length > 0) {
    out.push({ name: 'w:spacing', xml: `<w:spacing ${spacingAttrs.join(' ')}/>` })
  }
  const indAttrs: string[] = []
  if (format.indentLeft && format.indentLeft > 0)
    indAttrs.push(`w:left="${Math.round(format.indentLeft)}"`)
  if (format.indentRight && format.indentRight > 0)
    indAttrs.push(`w:right="${Math.round(format.indentRight)}"`)
  if (format.indentFirstLine) {
    if (format.indentFirstLine > 0)
      indAttrs.push(`w:firstLine="${Math.round(format.indentFirstLine)}"`)
    else indAttrs.push(`w:hanging="${Math.round(-format.indentFirstLine)}"`)
  }
  if (indAttrs.length > 0) out.push({ name: 'w:ind', xml: `<w:ind ${indAttrs.join(' ')}/>` })
  if (format.align) {
    let jc: string = format.align === 'justify' ? 'both' : format.align
    // convert visual value back to Word's logical value (inverse of the parse side)
    if (format.bidi && (jc === 'left' || jc === 'right')) jc = jc === 'left' ? 'right' : 'left'
    out.push({ name: 'w:jc', xml: `<w:jc w:val="${jc}"/>` })
  }
  if (format.tabStops && format.tabStops.length > 0) {
    const tabXml = format.tabStops
      .map((ts) => {
        let xml = `<w:tab w:val="${escapeXmlAttr(ts.val)}" w:pos="${ts.pos}"`
        if (ts.leader && ts.leader !== 'none') xml += ` w:leader="${escapeXmlAttr(ts.leader)}"`
        return xml + '/>'
      })
      .join('')
    out.push({ name: 'w:tabs', xml: `<w:tabs>${tabXml}</w:tabs>` })
  }
  if (format.dropCap) {
    const { type, lines } = format.dropCap
    out.push({
      name: 'w:framePr',
      xml: `<w:framePr w:dropCap="${type}" w:lines="${lines}" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>`,
    })
  }
  return out
}

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
  const rawLeft = Number.isFinite(left) && left !== 0 ? left : undefined
  const rawRight = Number.isFinite(right) && right !== 0 ? right : undefined
  const rawFirst = hanging > 0 ? -hanging : firstLine > 0 ? firstLine : undefined
  const norm = (v: number | undefined) => (v ? Math.round(v) : undefined)
  return (
    rawLeft === norm(f.indentLeft) &&
    rawRight === norm(f.indentRight) &&
    rawFirst === norm(f.indentFirstLine)
  )
}

function rawPBdrUnchanged(raw: string | undefined, f: ParaFormat): boolean {
  let rawBorders = ''
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
      if (el && rawAttr(el.xml, 'w:val') !== 'none') rawBorders += ch
    }
  }
  const norm = (s: string | undefined) => (s ? [...new Set(s)].sort().join('') : '')
  return norm(rawBorders) === norm(f.borders)
}

function rawTabsUnchanged(raw: string | undefined, stops: TabStop[]): boolean {
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
  const val = rawAttr(raw, 'w:dropCap')
  const rawDc =
    val === 'drop' || val === 'margin'
      ? { type: val, lines: parseInt(rawAttr(raw, 'w:lines') ?? '3', 10) || 3 }
      : undefined
  return rawDc?.type === f.dropCap?.type && rawDc?.lines === f.dropCap?.lines
}

/** True when re-parsing the raw child yields the current model value, i.e. the group was not edited */
function pprGroupUnchanged(tag: string, raw: string | undefined, f: ParaFormat): boolean {
  switch (tag) {
    case 'w:pageBreakBefore':
      return rawBool(raw) === !!f.pageBreakBefore
    case 'w:bidi':
      return rawBool(raw) === !!f.bidi
    case 'w:jc': {
      const val = rawAttr(raw, 'w:val')
      let rawAlign = val ? JC_TO_ALIGN[val] : undefined
      if (f.bidi && (rawAlign === 'left' || rawAlign === 'right')) {
        rawAlign = rawAlign === 'left' ? 'right' : 'left'
      }
      return rawAlign === f.align
    }
    case 'w:spacing':
      return rawSpacingUnchanged(raw, f)
    case 'w:ind':
      return rawIndUnchanged(raw, f)
    case 'w:pBdr':
      return rawPBdrUnchanged(raw, f)
    case 'w:shd': {
      const fill = rawAttr(raw, 'w:fill')
      return (fill && fill !== 'auto' ? fill : undefined) === f.shadingFill
    }
    case 'w:tabs':
      return rawTabsUnchanged(raw, f.tabStops ?? [])
    case 'w:framePr':
      return rawFramePrUnchanged(raw, f)
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
 * from the model (at their schema position) — a rebuilt w:ind intentionally drops
 * firstLineChars/leftChars so the user's new twips indent wins over the CJK
 * char-unit variant Word would otherwise prefer. Everything unmanaged — keepNext,
 * paragraph-mark rPr, pPrChange revision records... — keeps its original bytes.
 * When format.tabStops is set, w:tabs is also managed (replaced or removed).
 * When format.dropCap is set, w:framePr is also managed.
 */
export function mergePPrFormat(rawPPr: string, format: ParaFormat | undefined): string {
  const open = /^<w:pPr(?: [^>]*)?>/.exec(rawPPr)?.[0]
  const fresh = formatPPrChildren(format)
  if (!open) {
    // '<w:pPr/>' or unrecognized: rebuild from the format model alone
    return fresh.length > 0 ? `<w:pPr>${fresh.map((c) => c.xml).join('')}</w:pPr>` : ''
  }
  const inner = rawPPr.slice(open.length, rawPPr.length - '</w:pPr>'.length)
  // build set of managed tags for this format (base + conditional)
  const managedTags = new Set(FORMAT_MANAGED_TAGS)
  if (format?.tabStops !== undefined) managedTags.add('w:tabs')
  if (format?.dropCap !== undefined) managedTags.add('w:framePr')
  const rawChildren = splitXmlChildren(inner)
  const rawOf = (tag: string) => rawChildren.find((c) => c.name === tag)?.xml
  const rebuilt = new Set(
    [...managedTags].filter((tag) => !pprGroupUnchanged(tag, rawOf(tag), format ?? {})),
  )
  const freshOut = fresh.filter((c) => rebuilt.has(c.name))
  const kept = rawChildren.filter((c) => !rebuilt.has(c.name))
  const rank = (n: string) => PPR_CHILD_ORDER.indexOf(n)
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
    (change.date ? ` w:date="${escapeXmlAttr(String(change.date))}"` : '')
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
  const content =
    bookmarksXml(block.hiddenBookmarks) +
    bookmarksXml(block.bookmarks) +
    crossStarts +
    generateRunsXml(block.runs, ctx) +
    crossEnds
  if (block.rawPPr !== undefined) return `<w:p>${block.rawPPr}${content}</w:p>`

  const children: PPrChild[] = []
  let styleId: string | undefined
  if (block.type === 'heading') {
    const level = Math.min(Math.max(block.level ?? 1, 1), 9)
    styleId = block.styleId ?? ctx.headingStyleIds.get(level)
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
  children.sort((a, b) => PPR_CHILD_ORDER.indexOf(a.name) - PPR_CHILD_ORDER.indexOf(b.name))
  const pPr = children.length > 0 ? `<w:pPr>${children.map((c) => c.xml).join('')}</w:pPr>` : ''
  return `<w:p>${pPr}${content}</w:p>`
}

/** stable 31-bit id per bookmark name (start/end pair only needs to agree with itself) */
function bookmarkIdOf(name: string): number {
  let h = 0
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) | 0
  return Math.abs(h) % 0x7fffffff
}

function bookmarksXml(names: string[] | undefined): string {
  if (!names?.length) return ''
  return names
    .map((name) => {
      const id = bookmarkIdOf(name)
      return `<w:bookmarkStart w:id="${id}" w:name="${escapeXmlAttr(name)}"/><w:bookmarkEnd w:id="${id}"/>`
    })
    .join('')
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
  const side = (name: 'top' | 'left' | 'bottom' | 'right') => {
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
  return `<w:tcBorders>${side('top')}${side('left')}${side('bottom')}${side('right')}</w:tcBorders>`
}

function tableCellXml(
  cell: TableCell,
  width: number,
  verticalMerge?: 'restart' | 'continue',
): string {
  // Surgical rawTcPr patch: modeled children are replaced/inserted in schema order;
  // unmodeled ones (tcMar/textDirection/noWrap…) keep their original bytes
  const rawInner = cell.rawTcPr
    ? cell.rawTcPr.replace(/^<w:tcPr(?:\s[^>]*)?>/, '').replace(/<\/w:tcPr>$/, '')
    : ''
  const children: PPrChild[] =
    cell.rawTcPr && cell.rawTcPr.endsWith('</w:tcPr>') ? splitXmlChildren(rawInner) : []
  setTcPrChild(children, 'w:tcW', `<w:tcW w:w="${width}" w:type="dxa"/>`)
  setTcPrChild(
    children,
    'w:gridSpan',
    cell.colSpan && cell.colSpan > 1 ? `<w:gridSpan w:val="${cell.colSpan}"/>` : null,
  )
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
  const tcPr = children.map((c) => c.xml)
  const paragraphs = cell.richParas?.length
    ? cell.richParas
    : (cell.paras.length > 0 ? cell.paras : ['']).map((text) => ({
        align: cell.align,
        runs: text === '' ? [] : [{ text, bold: cell.bold, color: cell.color }],
      }))
  const content = paragraphs
    .map((paragraph) => {
      const align = paragraph.align ?? cell.align
      const list = 'list' in paragraph ? paragraph.list : undefined
      // schema order inside pPr: numPr before jc
      const numPr = list
        ? `<w:numPr><w:ilvl w:val="${list.ilvl}"/><w:numId w:val="${escapeXmlAttr(list.numId)}"/></w:numPr>`
        : ''
      const jc = align ? `<w:jc w:val="${align === 'justify' ? 'both' : align}"/>` : ''
      const pPr = numPr || jc ? `<w:pPr>${numPr}${jc}</w:pPr>` : ''
      return `<w:p>${pPr}${runsXml(paragraph.runs, null)}</w:p>`
    })
    .join('')
  // nested tables are regenerated from the model; OOXML requires tc to end with w:p
  const nested = (cell.nestedTables ?? []).map((nt) => generateTableModelXml(nt)).join('')
  return `<w:tc><w:tcPr>${tcPr.join('')}</w:tcPr>${content}${nested ? `${nested}<w:p/>` : ''}</w:tc>`
}

/**
 * Generate a complete table from the editable display model. This is used only
 * after structural edits; untouched tables and text-only edits keep their
 * original XML through the byte-preserving patch paths.
 */
export function generateTableModelXml(model: TableModel, originalTableXml?: string): string {
  const columnCount = Math.max(
    1,
    model.colWidthsPct?.length ?? 0,
    ...model.rows.map((row) => row.reduce((sum, cell) => sum + (cell.colSpan ?? 1), 0)),
  )
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
  const borders =
    '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('') +
    '</w:tblBorders>'
  const grid = `<w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>`
  const rows = model.rows
    .map((row, ri) => {
      let gridColumn = 0
      const cells = row.map((cell) => {
        const span = Math.max(1, cell.colSpan ?? 1)
        const width = widths
          .slice(gridColumn, gridColumn + span)
          .reduce((sum, value) => sum + value, 0)
        gridColumn += span
        return tableCellXml(cell, width)
      })
      // trPr: original bytes passed through (tblHeader/cantSplit etc.); row height is
      // replaced/inserted/removed per the model
      let trPr = model.rawTrPrs?.[ri] ?? ''
      if (trPr && !trPr.endsWith('</w:trPr>')) trPr = ''
      const h = model.rowHeightsTwips?.[ri]
      if (h != null && h > 0) {
        const tag = `<w:trHeight w:val="${Math.round(h)}" w:hRule="atLeast"/>`
        if (!trPr) trPr = `<w:trPr>${tag}</w:trPr>`
        else if (/<w:trHeight[^>]*\/>/.test(trPr)) trPr = trPr.replace(/<w:trHeight[^>]*\/>/, tag)
        else trPr = trPr.replace('</w:trPr>', `${tag}</w:trPr>`)
      } else if (trPr) {
        trPr = trPr.replace(/<w:trHeight[^>]*\/>/, '')
        if (/^<w:trPr(?:\s[^>]*)?>\s*<\/w:trPr>$/.test(trPr)) trPr = ''
      }
      return `<w:tr>${trPr}${cells.join('')}</w:tr>`
    })
    .join('')
  const originalTblPr = originalTableXml?.match(
    /<w:tblPr(?:\s[^>]*)?>[\s\S]*?<\/w:tblPr>|<w:tblPr(?:\s[^>]*)?\/>/,
  )?.[0]
  let tblPr =
    originalTblPr ??
    `<w:tblPr><w:tblW w:w="${totalWidth}" w:type="dxa"/>${borders}` +
      '<w:tblLayout w:type="fixed"/></w:tblPr>'
  // Table style reference: '' = remove, non-empty = replace/insert (tblStyle is the
  // first tblPr child)
  if (model.tblStyleId !== undefined) {
    tblPr = tblPr.replace(/<w:tblStyle[^>]*\/>/, '')
    if (model.tblStyleId !== '') {
      const tag = `<w:tblStyle w:val="${escapeXmlAttr(model.tblStyleId)}"/>`
      tblPr = /<w:tblPr\/>/.test(tblPr)
        ? tblPr.replace('<w:tblPr/>', `<w:tblPr>${tag}</w:tblPr>`)
        : tblPr.replace(/(<w:tblPr(?:\s[^>]*)?>)/, `$1${tag}`)
      // Style banding/header-row switches default to 04A0 (firstRow + banding)
      if (!/<w:tblLook[\s/>]/.test(tblPr)) {
        tblPr = tblPr.replace(
          '</w:tblPr>',
          '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>',
        )
      }
    }
  }
  return `<w:tbl>${tblPr}${grid}${rows}</w:tbl>`
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
  /** page number computed by real pagination (cached text; begin is dirty, so Word still recalculates on open) */
  pageNo?: number
}

/**
 * Generate a real TOC field as one w:p fragment per line. The begin fldChar is
 * marked dirty so Word recalculates entries and page numbers on open; the
 * static entry texts serve as the visible result until then.
 */
export function generateTocFieldXml(entries: TocEntry[]): string[] {
  if (entries.length === 0) return []
  const maxLevel = Math.min(Math.max(...entries.map((e) => e.level), 1), 9)
  const pPr = (level: number) =>
    `<w:pPr><w:pStyle w:val="TOC${Math.min(Math.max(level, 1), 9)}"/>` +
    '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>' +
    '<w:rPr><w:noProof/></w:rPr></w:pPr>'
  const entryRuns = (text: string, pageNo?: number) =>
    `<w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r>` +
    '<w:r><w:rPr><w:noProof/></w:rPr><w:tab/></w:r>' +
    (pageNo !== undefined ? `<w:r><w:rPr><w:noProof/></w:rPr><w:t>${pageNo}</w:t></w:r>` : '')
  const begin =
    '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-${maxLevel}" \\h \\z \\u </w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

  return entries.map((entry, i) => {
    const first = i === 0 ? begin : ''
    const last = i === entries.length - 1 ? end : ''
    return `<w:p>${pPr(entry.level)}${first}${entryRuns(entry.text, entry.pageNo)}${last}</w:p>`
  })
}

/**
 * Caption paragraph: `<label> <SEQ label> <text>`, e.g. "Figure 1 System architecture".
 * The SEQ field is marked dirty so Word renumbers all captions on open; the
 * static number is the visible result until then.
 */
export function generateCaptionXml(label: string, number: number, text: string): string {
  const rPr = '<w:rPr><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
  const run = (inner: string) => `<w:r>${rPr}${inner}</w:r>`
  return (
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="80" w:after="200"/></w:pPr>' +
    run(`<w:t xml:space="preserve">${escapeXmlText(label)} </w:t>`) +
    run('<w:fldChar w:fldCharType="begin" w:dirty="true"/>') +
    run(
      `<w:instrText xml:space="preserve"> SEQ ${escapeXmlText(label)} \\* ARABIC </w:instrText>`,
    ) +
    run('<w:fldChar w:fldCharType="separate"/>') +
    run(`<w:t>${number}</w:t>`) +
    run('<w:fldChar w:fldCharType="end"/>') +
    (text ? run(`<w:t xml:space="preserve"> ${escapeXmlText(text)}</w:t>`) : '') +
    '</w:p>'
  )
}

/**
 * INDEX field as one w:p per cached entry line, alphabetically sorted. The
 * begin fldChar is dirty so Word rebuilds entries and page numbers on open.
 */
export function generateIndexFieldXml(terms: string[]): string[] {
  const unique = [...new Set(terms.map((t) => t.trim()).filter(Boolean))]
  if (unique.length === 0) return []
  unique.sort((a, b) => a.localeCompare(b, 'zh-CN'))
  const begin =
    '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> INDEX \\c "2" </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  const pPr =
    '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="4300"/></w:tabs>' +
    '<w:rPr><w:noProof/></w:rPr></w:pPr>'
  return unique.map((term, i) => {
    const first = i === 0 ? begin : ''
    const last = i === unique.length - 1 ? end : ''
    return (
      `<w:p>${pPr}${first}` +
      `<w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve">${escapeXmlText(term)}</w:t></w:r>` +
      '<w:r><w:rPr><w:noProof/></w:rPr><w:tab/></w:r>' +
      `${last}</w:p>`
    )
  })
}

function generateRunsXml(runs: Run[], ctx: GenerateContext): string {
  return runsXml(runs, ctx.allocateHyperlinkRel)
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
function runsXml(runs: Run[], allocate: ((href: string) => string) | null): string {
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

  // serialize run indices [from, to) handling hyperlink grouping + comment markers
  const emitRange = (from: number, to: number) => {
    let i = from
    while (i < to) {
      const run = runs[i]
      if (run.link) {
        // group consecutive runs sharing the same link target
        const group: Run[] = []
        const groupStart = i
        const href = run.link.href
        let rId = run.link.rId
        while (i < to && runs[i].link && runs[i].link!.href === href) {
          rId = rId ?? runs[i].link!.rId
          group.push(runs[i])
          i++
        }
        // a comment boundary inside the link group widens to the whole hyperlink
        for (let j = groupStart; j < i; j++) parts.push(startsAt(j))
        const tooltip = group[0]?.link?.tooltip
        const tipAttr = tooltip ? ` w:tooltip="${escapeXmlAttr(tooltip)}"` : ''
        if (href.startsWith('#')) {
          const inner = group.map((r) => runFragmentXml(r, true)).join('')
          parts.push(
            `<w:hyperlink w:anchor="${escapeXmlAttr(href.slice(1))}"${tipAttr}>${inner}</w:hyperlink>`,
          )
        } else {
          const finalRId = rId ?? allocate?.(href)
          if (finalRId) {
            const inner = group.map((r) => runFragmentXml(r, true)).join('')
            parts.push(
              `<w:hyperlink r:id="${escapeXmlAttr(finalRId)}"${tipAttr}>${inner}</w:hyperlink>`,
            )
          } else {
            parts.push(group.map((r) => runFragmentXml(r, false)).join(''))
          }
        }
        for (let j = groupStart; j < i; j++) parts.push(endsAt(j))
      } else {
        parts.push(startsAt(i))
        parts.push(runFragmentXml(run, false))
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
          r.ins?.id ?? null,
          r.del?.author ?? null,
          r.del?.date ?? null,
          r.del?.id ?? null,
        ])
      : ''
  let revSeq = 9001
  const revAttrs = (info: { author: string; date?: string; id?: string }) =>
    ` w:id="${escapeXmlAttr(info.id ?? String(revSeq++))}"` +
    ` w:author="${escapeXmlAttr(info.author)}"` +
    (info.date ? ` w:date="${escapeXmlAttr(info.date)}"` : '')

  let g = 0
  while (g < runs.length) {
    const key = revKey(runs[g])
    let end = g
    while (end < runs.length && revKey(runs[end]) === key) end++
    if (key === '') {
      emitRange(g, end)
    } else {
      const { ins, del } = runs[g]
      if (ins) parts.push(`<w:ins${revAttrs(ins)}>`)
      if (del) parts.push(`<w:del${revAttrs(del)}>`)
      emitRange(g, end)
      if (del) parts.push('</w:del>')
      if (ins) parts.push('</w:ins>')
    }
    g = end
  }
  return parts.join('')
}

/**
 * One run's OOXML, including atomic constructs the plain-text run cannot
 * express: footnote/endnote reference markers and trailing XE index fields.
 */
function runFragmentXml(run: Run, insideLink: boolean): string {
  // atomic inline formula: the stored <m:oMath> fragment is already valid
  // paragraph content (patch.ts adds the xmlns:m declaration when missing)
  if (run.math) return run.math.omml
  // atomic phonetic guide: the exact <w:ruby> fragment re-wrapped in a run
  if (run.ruby) return `<w:r>${run.ruby.xml}</w:r>`
  if (run.noteRef) {
    const tag = run.noteRef.kind === 'footnote' ? 'w:footnoteReference' : 'w:endnoteReference'
    return (
      '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>' +
      `<${tag} w:id="${escapeXmlAttr(run.noteRef.id)}"/></w:r>`
    )
  }
  if (run.refField !== undefined) {
    // cross-reference: full REF field, run text is the cached display result;
    // the original instruction (with its switches) is written back verbatim
    const name = run.refField.replace(/"/g, '')
    const instr = run.refInstr ?? ` REF ${name} \\h `
    return (
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      `<w:r><w:instrText xml:space="preserve">${escapeXmlText(instr)}</w:instrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      generateRunXml({ ...run, refField: undefined }, insideLink) +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )
  }
  if (run.instrField !== undefined) {
    // Generic inline field: run text is the cached result; the instruction is written back verbatim
    return (
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      `<w:r><w:instrText xml:space="preserve"> ${escapeXmlText(run.instrField)} </w:instrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      generateRunXml({ ...run, instrField: undefined }, insideLink) +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )
  }
  let xml = run.text === '' ? '' : generateRunXml(run, insideLink)
  if (run.xeTerm !== undefined) {
    // embedded quotes would break the field instruction syntax; drop them
    const term = run.xeTerm.replace(/"/g, '')
    xml +=
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      `<w:r><w:instrText xml:space="preserve"> XE "${escapeXmlText(term)}" </w:instrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  }
  return xml
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
  'w:rPrChange',
]

/** rPr children modeled by the run model, grouped by modeled field (tags within a group are kept or rebuilt together) */
const RUN_MANAGED_GROUPS: Array<{ key: string; tags: string[] }> = [
  { key: 'rStyle', tags: ['w:rStyle'] },
  { key: 'rFonts', tags: ['w:rFonts'] },
  { key: 'bold', tags: ['w:b', 'w:bCs'] },
  { key: 'italic', tags: ['w:i', 'w:iCs'] },
  { key: 'strike', tags: ['w:strike'] },
  { key: 'color', tags: ['w:color'] },
  { key: 'size', tags: ['w:sz', 'w:szCs'] },
  { key: 'highlight', tags: ['w:highlight'] },
  { key: 'underline', tags: ['w:u'] },
  { key: 'vertAlign', tags: ['w:vertAlign'] },
  { key: 'rPrChange', tags: ['w:rPrChange'] },
]

const rawAttr = (xml: string | undefined, attr: string): string | undefined =>
  xml ? new RegExp(` ${attr}="([^"]*)"`).exec(xml)?.[1] : undefined

/** String version of boolProp: element present means true unless w:val negates it */
function rawBool(xml: string | undefined): boolean {
  if (!xml) return false
  const val = rawAttr(xml, 'w:val')
  if (val === undefined) return true
  return !['0', 'false', 'none', 'off'].includes(val.toLowerCase())
}

function revisionRPrChangeXml(run: Run): string | null {
  const change = run.rPrChange
  if (!change) return null
  const old = change.old ?? {}
  const props: string[] = []
  if (old.styleId) props.push(`<w:rStyle w:val="${escapeXmlAttr(old.styleId)}"/>`)
  if (old.font) {
    const font = escapeXmlAttr(old.font)
    props.push(
      `<w:rFonts w:ascii="${font}" w:eastAsia="${font}" w:hAnsi="${font}" w:cs="${font}"/>`,
    )
  }
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
  const attrs =
    ` w:id="${escapeXmlAttr(change.id ?? '0')}"` +
    ` w:author="${escapeXmlAttr(change.author)}"` +
    (change.date ? ` w:date="${escapeXmlAttr(change.date)}"` : '')
  return `<w:rPrChange${attrs}><w:rPr>${props.join('')}</w:rPr></w:rPrChange>`
}

/** Fresh rPr children for the modeled fields (one-to-one with what buildRun reads on the parse side) */
function modelRPrChildren(run: Run, insideLink: boolean): PPrChild[] {
  const out: PPrChild[] = []
  if (insideLink) out.push({ name: 'w:rStyle', xml: '<w:rStyle w:val="Hyperlink"/>' })
  else if (run.styleId)
    out.push({ name: 'w:rStyle', xml: `<w:rStyle w:val="${escapeXmlAttr(run.styleId)}"/>` })
  if (run.font) {
    const f = escapeXmlAttr(run.font)
    out.push({
      name: 'w:rFonts',
      xml: `<w:rFonts w:ascii="${f}" w:eastAsia="${f}" w:hAnsi="${f}" w:cs="${f}"/>`,
    })
  }
  if (run.bold) out.push({ name: 'w:b', xml: '<w:b/>' })
  if (run.italic) out.push({ name: 'w:i', xml: '<w:i/>' })
  if (run.strike) out.push({ name: 'w:strike', xml: '<w:strike/>' })
  if (run.color)
    out.push({ name: 'w:color', xml: `<w:color w:val="${escapeXmlAttr(run.color)}"/>` })
  if (run.sizeHalfPoints) {
    out.push({ name: 'w:sz', xml: `<w:sz w:val="${run.sizeHalfPoints}"/>` })
    out.push({ name: 'w:szCs', xml: `<w:szCs w:val="${run.sizeHalfPoints}"/>` })
  }
  if (run.highlight)
    out.push({ name: 'w:highlight', xml: `<w:highlight w:val="${escapeXmlAttr(run.highlight)}"/>` })
  if (run.underline) out.push({ name: 'w:u', xml: '<w:u w:val="single"/>' })
  if (run.vertAlign)
    out.push({ name: 'w:vertAlign', xml: `<w:vertAlign w:val="${run.vertAlign}"/>` })
  const rPrChange = revisionRPrChangeXml(run)
  if (rPrChange) out.push({ name: 'w:rPrChange', xml: rPrChange })
  return out
}

/**
 * Merge the raw rPr slice with the run model: groups whose model value matches the raw
 * encoding keep their original bytes (double underline/themeColor/all four rFonts slots
 * do not degrade); mismatched (edited) groups are rebuilt from the model; children the
 * model does not cover (caps/vanish/dstrike/bdr/shd/spacing/lang…) are always kept.
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
        const modeled = insideLink ? 'Hyperlink' : run.styleId
        // The parse side does not store Hyperlink in styleId: raw=Hyperlink with an empty
        // model counts as equal
        return raw === modeled || (raw === 'Hyperlink' && !modeled)
      }
      case 'rFonts': {
        const attrs = rawOf('w:rFonts')
        const raw =
          rawAttr(attrs, 'w:eastAsia') ?? rawAttr(attrs, 'w:ascii') ?? rawAttr(attrs, 'w:hAnsi')
        return raw === run.font
      }
      case 'bold':
        return rawBool(rawOf('w:b')) === !!run.bold
      case 'italic':
        return rawBool(rawOf('w:i')) === !!run.italic
      case 'strike':
        return rawBool(rawOf('w:strike')) === !!run.strike
      case 'color': {
        const raw = rawAttr(rawOf('w:color'), 'w:val')
        return (raw === 'auto' ? undefined : raw) === run.color
      }
      case 'size': {
        const raw = rawAttr(rawOf('w:sz'), 'w:val')
        return (raw ? parseInt(raw, 10) || undefined : undefined) === run.sizeHalfPoints
      }
      case 'highlight': {
        const raw = rawAttr(rawOf('w:highlight'), 'w:val')
        return (raw === 'none' ? undefined : raw) === run.highlight
      }
      case 'underline': {
        // w:u is not a boolean prop: no w:val (or val="none") means no
        // underline, regardless of other attributes like w:color. Must match
        // underlineProp() used at parse time.
        const val = rawAttr(rawOf('w:u'), 'w:val')
        return (val !== undefined && val !== 'none') === !!run.underline
      }
      case 'vertAlign': {
        const raw = rawAttr(rawOf('w:vertAlign'), 'w:val')
        const modeled = raw === 'superscript' || raw === 'subscript' ? raw : undefined
        return modeled === run.vertAlign
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
    freshOut.push(...freshByGroup.get(g.key)!)
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

  const txbxPart = opts.withTextbox
    ? `<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p></w:txbxContent></wps:txbx>`
    : ''

  const wsp =
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:cNvSpPr/>` +
    spPr +
    txbxPart +
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
    (opts.withTextbox
      ? `<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p></w:txbxContent></v:textbox>`
      : '') +
    `</v:rect>`

  const mcFallback = `<mc:Fallback>` + `<w:pict>${vmlRect}</w:pict>` + `</mc:Fallback>`

  const alternateContent =
    `<mc:AlternateContent ${mcNs2}>` + mcChoice + mcFallback + `</mc:AlternateContent>`

  return `<w:p><w:r>${alternateContent}</w:r></w:p>`
}

/** WordArt preset definitions (style name → shape fill/stroke/text color). */
export const WORDART_PRESETS = [
  {
    id: 'wordArt-1',
    label: 'Fill - accent color',
    colorHex: '4472C4',
    borderHex: '2F5496',
    noFill: false,
  },
  {
    id: 'wordArt-2',
    label: 'Gradient - blue violet',
    colorHex: '7B2FBE',
    borderHex: '4472C4',
    noFill: false,
  },
  {
    id: 'wordArt-3',
    label: 'Outline - no fill',
    colorHex: '4472C4',
    borderHex: '4472C4',
    noFill: true,
  },
  {
    id: 'wordArt-4',
    label: 'Shadow - dark',
    colorHex: '1F3864',
    borderHex: '1F3864',
    noFill: false,
  },
  {
    id: 'wordArt-5',
    label: 'Glow - orange',
    colorHex: 'ED7D31',
    borderHex: 'C55A11',
    noFill: false,
  },
  {
    id: 'wordArt-6',
    label: 'Fill - dark red',
    colorHex: 'C00000',
    borderHex: '9B0000',
    noFill: false,
  },
] as const

/**
 * Build a <w:p> fragment containing a floating WordArt WPS text box.
 * The shape has no background fill; the text runs carry large size (36pt)
 * and the specified color. Style is approximated — Word can open the result.
 */
export function buildWordArtParagraphXml(opts: {
  text?: string
  wordArtId?: string
  widthEmu?: number
  heightEmu?: number
  id?: number
}): string {
  const widthEmu = opts.widthEmu ?? 2700000 // ~7.5 cm
  const heightEmu = opts.heightEmu ?? 720000 // ~2 cm
  const id = opts.id ?? 1
  const preset = WORDART_PRESETS.find((p) => p.id === opts.wordArtId) ?? WORDART_PRESETS[0]
  const text = opts.text ?? 'WordArt'

  // Build run props: large font + color
  const rPr =
    `<w:rPr>` +
    `<w:b/>` +
    `<w:color w:val="${preset.colorHex}"/>` +
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

function generateRunXml(run: Run, insideLink: boolean): string {
  // OOXML requires rPr children in schema order:
  // rStyle < rFonts < b < i < strike < color < sz < highlight < u < vertAlign
  let rPr: string
  if (run.rawRPr !== undefined) {
    rPr = mergeRPrModel(run.rawRPr, run, insideLink)
  } else {
    const props = modelRPrChildren(run, insideLink).map((c) => c.xml)
    rPr = props.length > 0 ? `<w:rPr>${props.join('')}</w:rPr>` : ''
  }

  // Translate embedded control characters back to OOXML elements.
  // Deleted runs carry their text in w:delText instead of w:t.
  const textTag = run.del ? 'w:delText' : 'w:t'
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
    } else {
      buffer += ch
    }
  }
  flush()
  return `<w:r>${rPr}${segments.join('')}</w:r>`
}
