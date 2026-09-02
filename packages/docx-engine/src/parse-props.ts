// Pure XML -> value readers for paragraph, run and table-cell properties
// (no document context, no recursion into block builders).
import {
  attrsOf,
  childrenOf,
  findChild,
  findChildren,
  nameOf,
  textOf,
  xmlParser,
  type XNode,
} from './xml-utils'
import {
  EMU_PER_PX,
  decodeEntities,
  decodeNumericCharRefs,
  onOffOf,
  plainText,
  stripHash,
} from './parse-xml-text'
import type {
  Block,
  CellBorders,
  CellMargins,
  CharIndents,
  ParaFormat,
  RevisionInfo,
  Run,
  TableBorders,
  TableModel,
  ThemeFonts,
} from './types'

/** No run un-hides itself and nothing anchors here (bookmarks, comments, sectPr,
 *  drawings, numbering): safe to collapse a style-vanished paragraph entirely */
export function staysVanished(xml: string): boolean {
  if (/<w:vanish\s[^>]*w:val="(?:0|false|off)"/.test(xml)) return false
  return !/<w:(?:drawing|pict|object|sectPr|bookmarkStart|commentRangeStart|commentRangeEnd|numPr)[\s/>]/.test(
    xml,
  )
}

/** Text-less run content that still affects layout (breaks, tabs, note marks,
 *  symbol chars): a "visually empty" paragraph carrying one must not collapse.
 *  w:pPr is skipped so tab-stop definitions (w:tabs > w:tab) don't count. */
const LAYOUT_RUN_CONTENT = /^w:(?:br|cr|tab|sym|footnoteReference|endnoteReference)$/

export function hasLayoutRunContent(node: XNode): boolean {
  for (const child of childrenOf(node)) {
    const name = nameOf(child)
    if (name === 'w:pPr' || name === 'w:rPr') continue
    if (name !== undefined && LAYOUT_RUN_CONTENT.test(name)) return true
    if (hasLayoutRunContent(child)) return true
  }
  return false
}

/**
 * Cross-paragraph comment range endpoints: comment ids where only one end falls in this
 * paragraph (the other end is in a different paragraph). Ranges fully within one
 * paragraph are handled by run.commentIds; this only catches cross-paragraph ones so a
 * paragraph rebuild does not leave orphaned commentRangeEnd/Start markers.
 */
export function crossParaCommentMarkers(xml: string): {
  commentStarts: string[] | undefined
  commentEnds: string[] | undefined
} {
  const ids = (re: RegExp) => [...xml.matchAll(re)].map((m) => m[1])
  const starts = ids(/<w:commentRangeStart [^>]*w:id="([^"]+)"/g)
  const ends = ids(/<w:commentRangeEnd [^>]*w:id="([^"]+)"/g)
  const onlyStarts = starts.filter((id) => !ends.includes(id))
  const onlyEnds = ends.filter((id) => !starts.includes(id))
  return {
    commentStarts: onlyStarts.length ? onlyStarts : undefined,
    commentEnds: onlyEnds.length ? onlyEnds : undefined,
  }
}

/** user bookmark names starting in this paragraph; Word internals (_Toc/_Ref/_GoBack…) split out as hiddenBookmarks */
export function bookmarkNamesOf(xml: string): {
  bookmarks: string[] | undefined
  hiddenBookmarks: string[] | undefined
} {
  const names: string[] = []
  const hidden: string[] = []
  for (const m of xml.matchAll(/<w:bookmarkStart [^>]*w:name="([^"]+)"/g)) {
    const name = decodeEntities(m[1])
    // A _ prefix marks Word internal bookmarks (_Ref/_Toc/_Hlk): hidden from the UI, but
    // they must be re-emitted when the paragraph rebuilds, otherwise REF cross-references
    // and TOC anchors pointing at them break
    const list = name.startsWith('_') ? hidden : names
    if (!list.includes(name)) list.push(name)
  }
  return {
    bookmarks: names.length > 0 ? names : undefined,
    hiddenBookmarks: hidden.length > 0 ? hidden : undefined,
  }
}

/**
 * Exact <w:pPr>...</w:pPr> slice of a paragraph (depth-aware: pPrChange nests
 * another w:pPr inside). undefined when the paragraph has no properties.
 */
export function rawPPrOf(xml: string): string | undefined {
  // the paragraph's own pPr must be the first child of w:p; a later match
  // would belong to nested content (textbox paragraphs)
  const openEnd = xml.indexOf('>') + 1
  if (openEnd === 0 || !xml.startsWith('<w:pPr', openEnd)) return undefined
  const start = openEnd
  const re = /<w:pPr(?=[\s/>])|<\/w:pPr>/g
  re.lastIndex = start
  let depth = 0
  let match: RegExpExecArray | null
  while ((match = re.exec(xml)) !== null) {
    if (match[0] === '</w:pPr>') {
      depth--
      if (depth === 0) return xml.slice(start, match.index + match[0].length)
    } else {
      // self-closing <w:pPr/> never opens
      const gt = xml.indexOf('>', match.index)
      if (xml[gt - 1] === '/') {
        if (depth === 0) return xml.slice(start, gt + 1)
        continue
      }
      depth++
    }
  }
  return undefined
}

/**
 * A text-less anchored shape no taller than ~10px is almost always a
 * decorative horizontal rule (heading underlines, dividers); render those as a
 * line instead of a drawing-object chip.
 */
export function isThinRule(xml: string): boolean {
  const m = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml)
  if (!m) return false
  const cx = parseInt(m[1], 10)
  const cy = parseInt(m[2], 10)
  // Word writes plain horizontal lines with cy="0"
  return cy <= 130000 && (cy > 0 || cx > 0)
}

/** stroke color/thickness (a:ln) + extent width of a decorative rule drawing */
export function ruleDisplayOf(
  xml: string,
): Pick<Block, 'ruleColorHex' | 'ruleThicknessPx' | 'ruleWidthPx'> {
  const out: Pick<Block, 'ruleColorHex' | 'ruleThicknessPx' | 'ruleWidthPx'> = {}
  const ln = /<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/.exec(xml)?.[0]
  if (ln) {
    const color = /<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(ln)?.[1]
    if (color) out.ruleColorHex = color.toUpperCase()
    const w = parseInt(/<a:ln\b[^>]*\bw="(\d+)"/.exec(ln)?.[1] ?? '', 10)
    if (Number.isFinite(w) && w > 0) out.ruleThicknessPx = Math.max(1, Math.round(w / EMU_PER_PX))
  }
  const cx = parseInt(/<wp:extent cx="(\d+)"/.exec(xml)?.[1] ?? '', 10)
  if (Number.isFinite(cx) && cx > 0) out.ruleWidthPx = Math.round(cx / EMU_PER_PX)
  return out
}

/**
 * Converter artifact: every shape explicitly declares noFill + noFill outline
 * and carries no picture, no text and no effects — Word renders nothing, so
 * the block renders as nothing too (still passthrough, bytes untouched).
 */
export function isInvisibleEmptyShape(xml: string): boolean {
  if (!xml.includes('<wps:wsp') || xml.includes('<a:blip') || plainText(xml).trim() !== '') {
    return false
  }
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(xml) as XNode[]
  } catch {
    return false
  }
  const shapes: XNode[] = []
  collectNodes(parsed, 'wps:wsp', shapes)
  if (shapes.length === 0) return false
  return shapes.every((shape) => {
    const spPr = findChild(shape, 'wps:spPr')
    if (!spPr || !findChild(spPr, 'a:noFill')) return false
    const ln = findChild(spPr, 'a:ln')
    if (!ln || !findChild(ln, 'a:noFill')) return false
    const effects = findChild(spPr, 'a:effectLst')
    return !effects || childrenOf(effects).length === 0
  })
}

/**
 * Legacy VML pict that draws nothing: only v:shapetype definitions (geometry
 * templates, not drawn), shapes hidden via style visibility:hidden, or white
 * strokeless placeholder rectangles (LO fixtures / converter watermark
 * furniture). Word renders nothing there — an "Embedded object" chip would
 * paint a box on an otherwise blank page.
 */
export function isInvisibleVmlPict(xml: string): boolean {
  if (!/<v:(?:shapetype|shape|rect|roundrect|oval|line|polyline)\b/.test(xml)) return false
  if (xml.includes('<v:imagedata') || xml.includes('<w:txbxContent')) return false
  const shapes = xml.match(/<v:(?:shape|rect|roundrect|oval|line|polyline)\b[^>]*>/g) ?? []
  // no drawable instances at all (shapetype-only pict) also renders nothing
  return shapes.every((tag) => {
    const style = /style="([^"]*)"/.exec(tag)?.[1] ?? ''
    if (/visibility:\s*hidden/.test(style)) return true
    const fill = /fillcolor="([^"]+)"/.exec(tag)?.[1]?.trim().toLowerCase()
    const unstroked = /\bstroked="(?:f|false|0)"/.test(tag)
    return unstroked && (fill === 'white' || fill === '#ffffff' || fill === '#fff')
  })
}

/** drop textbox content so "does the paragraph itself have text" checks work */
export function stripTextboxes(xml: string): string {
  return xml.includes('<w:txbxContent')
    ? xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '')
    : xml
}

/** page-type w:br carried by the anchor paragraph itself (not inside box content) */
export function hostPageBreak(xml: string): boolean {
  return /<w:br\s[^>]*w:type="page"/.test(stripTextboxes(xml))
}

/**
 * w:tbl or w:sdt among the txbxContent children: the display lines no longer
 * map 1:1 onto the w:p segments patch-save rewrites, so the box must stay
 * read-only (editing would drop the table / sdt shells).
 */
export function txbxHasStructuredContent(content: XNode): boolean {
  return childrenOf(content).some((c) => {
    const n = nameOf(c)
    return n === 'w:tbl' || n === 'w:sdt'
  })
}

export function collectNodes(nodes: XNode[], name: string, out: XNode[]): void {
  for (const node of nodes) {
    if (nameOf(node) === name) out.push(node)
    collectNodes(childrenOf(node), name, out)
  }
}

/** like collectNodes, but does not descend into a matched node — top-level
 *  matches only, the set xmlSegments counts on the save path */
export function collectTopNodes(nodes: XNode[], name: string, out: XNode[]): void {
  for (const node of nodes) {
    if (nameOf(node) === name) {
      out.push(node)
      continue
    }
    collectTopNodes(childrenOf(node), name, out)
  }
}

export const JC_ALIGN: Record<string, ParaFormat['align']> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  both: 'justify',
  distribute: 'distribute',
}

/** w:autoSpaceDE/DN (Word default on): false only when both are explicitly off */
export function autoSpaceOf(pPr: XNode): boolean | undefined {
  const de = onOffOf(pPr, 'w:autoSpaceDE')
  const dn = onOffOf(pPr, 'w:autoSpaceDN')
  if (de === false && dn === false) return false
  if (de === true || dn === true) return true
  return undefined
}

/** pPr <w:tabs> → TabStop[] (shared by direct paragraph format and style display) */
export function tabStopsOf(pPr: XNode): import('./types').TabStop[] | undefined {
  const tabsEl = findChild(pPr, 'w:tabs')
  if (!tabsEl) return undefined
  const stops: import('./types').TabStop[] = []
  for (const tab of findChildren(tabsEl, 'w:tab')) {
    const attrs = attrsOf(tab)
    const pos = parseInt(attrs['w:pos'] ?? '', 10)
    const val = attrs['w:val'] ?? 'left'
    if (!Number.isFinite(pos)) continue
    const validVals = ['left', 'center', 'right', 'decimal', 'bar', 'clear'] as const
    const safeVal = validVals.includes(val as (typeof validVals)[number])
      ? (val as (typeof validVals)[number])
      : 'left'
    const stop: import('./types').TabStop = { pos, val: safeVal }
    const leader = attrs['w:leader']
    if (leader && leader !== 'none') {
      const validLeaders = ['dot', 'hyphen', 'underscore', 'heavy', 'middleDot'] as const
      if (validLeaders.includes(leader as (typeof validLeaders)[number])) {
        stop.leader = leader as (typeof validLeaders)[number]
      }
    }
    stops.push(stop)
  }
  return stops.length > 0 ? stops : undefined
}

/** display-only stops for the paragraph's w:ptab elements (left-aligned ones advance nothing) */
export function ptabDisplayStops(pNode: XNode): import('./types').TabStop[] {
  const out: import('./types').TabStop[] = []
  const walk = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      const name = nameOf(c)
      if (name === 'w:pPr') continue
      if (name === 'w:ptab') {
        const a = attrsOf(c)
        const align = a['w:alignment']
        if (align !== 'center' && align !== 'right') continue
        const stop: import('./types').TabStop = {
          pos: align === 'center' ? 50 : 100,
          val: align,
          rel: 'margin',
        }
        const leader = a['w:leader']
        if (
          leader === 'dot' ||
          leader === 'hyphen' ||
          leader === 'underscore' ||
          leader === 'middleDot'
        ) {
          stop.leader = leader
        }
        out.push(stop)
      } else walk(c)
    }
  }
  walk(pNode)
  return out
}

/**
 * True when every field in the paragraph is an XE (index entry) marker.
 * Multi-fragment instructions or fldSimple fields fail the check, so anything
 * unusual falls back to the protected-passthrough path.
 */
/** paragraphs whose only fields are XE / REF stay editable (extractRuns round-trips them) */
/** Simple instructions foldable into an editable inline-field run (the cached result is the display text) */
export const SIMPLE_INLINE_FIELD_RE =
  /^\s*(DATE|TIME|CREATEDATE|SAVEDATE|NUMPAGES|FILENAME|AUTHOR|PAGE)\b/

/** HYPERLINK "url" (optional \o "tip"): the only field form folded into an editable link run;
 * any other switch (\l bookmark, \t frame...) keeps the protected-passthrough path */
export function convertibleHyperlink(instr: string): { href: string; tooltip?: string } | null {
  const m = /^\s*HYPERLINK\s+"([^"\\]+)"\s*(?:\\o\s+"([^"]*)"\s*)?$/.exec(
    decodeNumericCharRefs(instr),
  )
  if (!m) return null
  return { href: m[1], ...(m[2] ? { tooltip: m[2] } : {}) }
}

/** every field instruction is EMBED/LINK (the object-field forms of legacy OLE) */
export function onlyOleFields(xml: string): boolean {
  if (xml.includes('<w:fldSimple')) return false
  const instrs = xml.match(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g) ?? []
  if (instrs.length === 0) return false
  return instrs.every((fragment) =>
    /^\s*(EMBED|LINK)\b/.test(decodeEntities(fragment.replace(/<[^>]+>/g, ''))),
  )
}

export function onlyXeFields(xml: string): boolean {
  if (xml.includes('<w:fldSimple')) return false
  const instrs = xml.match(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g) ?? []
  if (instrs.length === 0) return false
  let checkboxInstrs = 0
  const ok = instrs.every((fragment) => {
    const text = decodeEntities(fragment.replace(/<[^>]+>/g, ''))
    if (/^\s*FORMCHECKBOX\s*$/.test(text)) {
      checkboxInstrs++
      return true
    }
    return (
      /^\s*XE[\s"]/.test(text) ||
      /^\s*REF\s/.test(text) ||
      SIMPLE_INLINE_FIELD_RE.test(text) ||
      convertibleHyperlink(text) !== null
    )
  })
  if (!ok) return false
  // a FORMCHECKBOX without its w:checkBox definition can't fold into a glyph
  // run; keep the paragraph on the byte-preserving passthrough path instead
  // of silently dropping the field on save
  const checkBoxDefs = (xml.match(/<w:checkBox[\s/>]/g) ?? []).length
  return checkboxInstrs <= checkBoxDefs
}

/** Checked state of a FORMCHECKBOX begin run (w:fldChar/w:ffData/w:checkBox):
 * w:checked wins over w:default; either element without w:val means true */
export function checkboxStateOf(beginRun: XNode | null): { checked: boolean } | null {
  if (!beginRun) return null
  const ffData = findChild(findChild(beginRun, 'w:fldChar') ?? {}, 'w:ffData')
  const box = ffData ? findChild(ffData, 'w:checkBox') : undefined
  if (!box) return null
  const state = findChild(box, 'w:checked') ?? findChild(box, 'w:default')
  if (!state) return { checked: false }
  const val = attrsOf(state)['w:val']
  return { checked: val === undefined || val === '1' || val === 'true' || val === 'on' }
}

/**
 * Character-unit attributes a w:ind specifies. Zeros are kept: Word writes
 * `w:firstLineChars="0"` next to an absolute `w:firstLine` to say the indent is
 * NOT character-based, and such a zero also cancels a *Chars inherited from
 * the style chain (probed) — see mergeCharIndents / activeCharIndents.
 */
export function charIndentsOf(ind: XNode | undefined): CharIndents | undefined {
  if (!ind) return undefined
  const a = attrsOf(ind)
  const num = (v: string | undefined): number | undefined => {
    const n = parseInt(v ?? '', 10)
    return Number.isFinite(n) ? n : undefined
  }
  const chars: CharIndents = {}
  const left = num(a['w:leftChars'] ?? a['w:startChars'])
  if (left !== undefined) chars.left = left
  const right = num(a['w:rightChars'] ?? a['w:endChars'])
  if (right !== undefined) chars.right = right
  const firstLine = num(a['w:firstLineChars'])
  if (firstLine !== undefined && firstLine >= 0) chars.firstLine = firstLine
  const hanging = num(a['w:hangingChars'])
  if (hanging !== undefined && hanging >= 0) chars.hanging = hanging
  return Object.keys(chars).length > 0 ? chars : undefined
}

/**
 * Layer one w:ind's character-unit attributes over inherited ones the way Word
 * resolves a style chain and direct formatting (Word for Mac probe, 2026-09-02):
 * a *Chars attribute — zero included — replaces the inherited value of its
 * component, a twips attribute never does (a child style's or paragraph's
 * twips-only w:firstLine / w:left leaves an inherited firstLineChars /
 * leftChars in force). firstLineChars and hangingChars are one component, the
 * special indent: setting either replaces both.
 */
export function mergeCharIndents(
  base: CharIndents | undefined,
  over: CharIndents | undefined,
): CharIndents | undefined {
  if (!base) return over
  if (!over) return base
  const merged: CharIndents = { ...base }
  if (over.left !== undefined) merged.left = over.left
  if (over.right !== undefined) merged.right = over.right
  if (over.firstLine !== undefined || over.hanging !== undefined) {
    delete merged.firstLine
    delete merged.hanging
    if (over.firstLine !== undefined) merged.firstLine = over.firstLine
    if (over.hanging !== undefined) merged.hanging = over.hanging
  }
  return merged
}

/** the character indents that lay out: an explicit zero means "absolute" and drops out */
export function activeCharIndents(chars: CharIndents | undefined): CharIndents | undefined {
  if (!chars) return undefined
  const active: CharIndents = {}
  if (chars.left) active.left = chars.left
  if (chars.right) active.right = chars.right
  if (chars.firstLine) active.firstLine = chars.firstLine
  if (chars.hanging) active.hanging = chars.hanging
  return Object.keys(active).length > 0 ? active : undefined
}

/** twips of one "character" for each family of character-unit indents */
export interface CharUnits {
  /** w:firstLineChars / w:hangingChars: the paragraph's first text run — its font
   *  size plus its letter spacing (plus the section's character-grid pitch delta) */
  run: number
  /** w:leftChars / w:rightChars: the default paragraph style's (Normal) font size —
   *  whatever the paragraph's own style or runs say (plus the grid delta) */
  normal: number
}

/**
 * Fold character-unit indents into the twips fields of a paragraph format,
 * the way Word lays them out (Word for Mac probe, 2026-09-02):
 *
 * - a nonzero *Chars attribute supersedes its twips twin (firstLineChars over
 *   firstLine, leftChars over left, rightChars over right), and a character
 *   special indent beats a twips one of the other kind (firstLineChars over
 *   w:hanging); between two character specials hanging wins, as with twips;
 * - firstLineChars/hangingChars scale with the first text run's size (a 10pt
 *   first run in a 20pt paragraph gives a 10pt unit; the paragraph mark's size
 *   never counts), leftChars/rightChars with the Normal style's size;
 * - once leftChars or hangingChars is present the left edge is in "character
 *   mode": leftChars is the position of the first line (Word's dialog "Left"),
 *   the body lines sit `hanging` further in — hangingChars, else the twips
 *   w:hanging — and a twips w:left is ignored even when leftChars is absent.
 *   firstLineChars alone keeps a twips w:left as the body indent.
 */
export function resolveCharIndents(
  format: ParaFormat | undefined,
  chars: CharIndents,
  units: CharUnits,
): ParaFormat {
  const f: ParaFormat = { ...(format ?? {}) }
  const twips = (hundredths: number, unit: number) => Math.round((hundredths / 100) * unit)
  const twipsHanging =
    f.indentFirstLine !== undefined && f.indentFirstLine < 0 ? -f.indentFirstLine : 0
  const twipsFirst =
    f.indentFirstLine !== undefined && f.indentFirstLine > 0 ? f.indentFirstLine : 0
  const hanging =
    chars.hanging !== undefined
      ? twips(chars.hanging, units.run)
      : chars.firstLine !== undefined
        ? 0
        : twipsHanging
  const firstLine =
    hanging > 0 ? 0 : chars.firstLine !== undefined ? twips(chars.firstLine, units.run) : twipsFirst
  if (chars.left !== undefined || chars.hanging !== undefined) {
    const uiLeft = chars.left !== undefined ? twips(chars.left, units.normal) : 0
    f.indentLeft = uiLeft + hanging
  }
  if (hanging > 0) f.indentFirstLine = -hanging
  else if (firstLine > 0) f.indentFirstLine = firstLine
  else delete f.indentFirstLine
  if (chars.right !== undefined) {
    const right = twips(chars.right, units.normal)
    if (right !== 0) f.indentRight = right
    else delete f.indentRight
  }
  return f
}

/**
 * w:sz governing a run-less paragraph's line height: the paragraph-mark rPr
 * (pPr/w:rPr), else the last run's rPr — those runs are all empty and get
 * dropped, but Word still sizes the empty line by them (1pt spacer lines).
 */
export function emptyParaSizeHalfPoints(pNode: XNode, pPr: XNode | undefined): number | undefined {
  let sz = pPr
    ? attrsOf(findChild(findChild(pPr, 'w:rPr') ?? {}, 'w:sz') ?? {})['w:val']
    : undefined
  if (!sz) {
    for (const r of findChildren(pNode, 'w:r')) {
      const v = attrsOf(findChild(findChild(r, 'w:rPr') ?? {}, 'w:sz') ?? {})['w:val']
      if (v) sz = v
    }
  }
  const n = sz ? parseInt(sz, 10) : NaN
  return Number.isFinite(n) && n > 0 ? n : undefined
}

/**
 * w:rFonts governing a run-less paragraph's line height: same sources as
 * emptyParaSizeHalfPoints — Word lays the empty line with the mark's face
 * metrics, not the document default face.
 */
export function emptyParaMarkFont(pNode: XNode, pPr: XNode | undefined): string | undefined {
  const pick = (rPr: XNode | undefined): string | undefined => {
    const a = attrsOf(findChild(rPr ?? {}, 'w:rFonts') ?? {})
    return a['w:ascii'] ?? a['w:hAnsi'] ?? a['w:eastAsia']
  }
  let font = pPr ? pick(findChild(pPr, 'w:rPr')) : undefined
  if (!font) {
    for (const r of findChildren(pNode, 'w:r')) {
      const v = pick(findChild(r, 'w:rPr'))
      if (v) font = v
    }
  }
  return font
}

export const IMAGE_RUN_CHILDREN = new Set([
  'w:drawing',
  'w:pict',
  'w:object',
  'mc:AlternateContent',
])

/** synthetic single-picture runs (rPr copied into each) in document order */
export function splitImageRun(rNode: XNode): XNode[] {
  const attrs = rNode[':@']
  const rPr = findChild(rNode, 'w:rPr')
  const parts: XNode[][] = []
  let current: XNode[] = rPr ? [rPr] : []
  for (const child of childrenOf(rNode)) {
    if (nameOf(child) === 'w:rPr') continue
    current.push(child)
    if (IMAGE_RUN_CHILDREN.has(nameOf(child) ?? '')) {
      parts.push(current)
      current = rPr ? [rPr] : []
    }
  }
  if (current.length > (rPr ? 1 : 0)) parts.push(current)
  return parts.map((children) => ({ 'w:r': children, ...(attrs ? { ':@': attrs } : {}) }))
}

/** exact <w:ruby> fragments in document order (w:ruby has no attributes and cannot nest) */
export function rubyFragmentsOf(xml: string): string[] {
  return xml.match(/<w:ruby>[\s\S]*?<\/w:ruby>/g) ?? []
}

/** concatenated w:t text of the runs inside a ruby part (w:rubyBase / w:rt) */
export function rubyPartText(rubyNode: XNode, part: 'w:rubyBase' | 'w:rt'): string {
  const partNode = findChild(rubyNode, part)
  if (!partNode) return ''
  let text = ''
  for (const r of childrenOf(partNode)) {
    if (nameOf(r) !== 'w:r') continue
    for (const c of childrenOf(r)) {
      if (nameOf(c) === 'w:t') text += decodeNumericCharRefs(textOf(c))
    }
  }
  return text
}

/** Word's face for an East Asian theme slot whose typeface is empty (<a:ea typeface=""/>) */
const EMPTY_EA_THEME_FONT = 'DengXian'

/** empty EA slot faces by settings.xml w:themeFontLang w:eastAsia (zh-CN / missing → DengXian) */
const EMPTY_EA_SLOT_BY_LANG: Record<string, { major: string; minor: string }> = {
  ja: { major: 'Yu Gothic', minor: 'Yu Mincho' },
  // Word probe + fontTable: ko-KR empty EA slot substitutes Malgun Gothic
  ko: { major: 'Malgun Gothic', minor: 'Malgun Gothic' },
}

/** themeFontLang w:eastAsia → theme script-table tag (<a:font script=…>) */
const EA_LANG_SCRIPT: Record<string, string> = {
  ko: 'Hang',
  ja: 'Jpan',
  zh: 'Hans',
  'zh-cn': 'Hans',
  'zh-sg': 'Hans',
  'zh-tw': 'Hant',
  'zh-hk': 'Hant',
  'zh-mo': 'Hant',
}

function emptyEaSlotFont(fonts: ThemeFonts, eaRef: string | undefined): string {
  return themeLangEaSlotFont(fonts, eaRef) ?? EMPTY_EA_THEME_FONT
}

/** empty-EA-slot face themeFontLang specifically resolves (theme script table first,
 * then the probed per-language defaults); undefined = only the generic DengXian applies */
export function themeLangEaSlotFont(
  fonts: ThemeFonts,
  eaRef: string | undefined,
): string | undefined {
  const full = fonts.eaLang?.toLowerCase()
  if (!full) return undefined
  const lang = full.split('-')[0]
  const script = EA_LANG_SCRIPT[full] ?? EA_LANG_SCRIPT[lang]
  const table = eaRef === 'majorEastAsia' ? fonts.majorScripts : fonts.minorScripts
  const fromScript = script ? table?.[script] : undefined
  if (fromScript) return fromScript
  const byLang = EMPTY_EA_SLOT_BY_LANG[lang]
  return byLang ? (eaRef === 'majorEastAsia' ? byLang.major : byLang.minor) : undefined
}

/** w:rFonts with theme references resolved: theme attrs supersede same-slot literal
 * values (ECMA-376 §17.3.2.26). Unresolvable references fall back to the literal,
 * except an empty eastAsia theme slot: Word keeps the theme's authority and renders
 * the theme language's default face, never the leftover literal name (eaSlotEmpty marks this). */
export function themedRFonts(
  attrs: Record<string, string | undefined>,
  fonts: ThemeFonts | null | undefined,
): {
  ascii?: string
  hAnsi?: string
  eastAsia?: string
  cs?: string
  eaSlotEmpty?: boolean
  themed?: { ascii?: boolean; hAnsi?: boolean; eastAsia?: boolean }
} {
  const themeVal = (ref: string | undefined): string | undefined => {
    if (!ref || !fonts) return undefined
    switch (ref) {
      case 'majorAscii':
      case 'majorHAnsi':
        return fonts.major || undefined
      case 'minorAscii':
      case 'minorHAnsi':
        return fonts.minor || undefined
      case 'majorEastAsia':
        return fonts.majorEastAsia || undefined
      case 'minorEastAsia':
        return fonts.eastAsia || undefined
      case 'majorBidi':
        return fonts.majorCs || undefined
      case 'minorBidi':
        return fonts.minorCs || undefined
      default:
        return undefined
    }
  }
  const eaRef = attrs['w:eastAsiaTheme']
  const themedEa = themeVal(eaRef)
  const eaSlotEmpty =
    !themedEa && !!fonts && (eaRef === 'majorEastAsia' || eaRef === 'minorEastAsia')
  const themedAscii = themeVal(attrs['w:asciiTheme'])
  const themedHAnsi = themeVal(attrs['w:hAnsiTheme'])
  return {
    ascii: themedAscii ?? attrs['w:ascii'],
    hAnsi: themedHAnsi ?? attrs['w:hAnsi'],
    eastAsia: themedEa ?? (eaSlotEmpty ? emptyEaSlotFont(fonts!, eaRef) : attrs['w:eastAsia']),
    cs: themeVal(attrs['w:cstheme']) ?? attrs['w:cs'],
    ...(eaSlotEmpty ? { eaSlotEmpty } : {}),
    themed: {
      ascii: themedAscii !== undefined,
      hAnsi: themedHAnsi !== undefined,
      eastAsia: themedEa !== undefined || eaSlotEmpty,
    },
  }
}

/** xml:space="preserve" declared on the part's root element — an inherited XML
 *  scope covering every w:t below it (PDF-to-DOCX converters set it once on
 *  w:document/w:hdr instead of per element; Word honors the inheritance) */
export function partXmlSpacePreserve(partXml: string, rootTag: string): boolean {
  const open = new RegExp(`<${rootTag}(\\s[^>]*)?>`).exec(partXml)?.[1] ?? ''
  return /\sxml:space="preserve"/.test(open)
}

export function mergeRuns(runs: Run[]): Run[] {
  const merged: Run[] = []
  for (const run of runs) {
    const prev = merged[merged.length - 1]
    if (prev && sameStyle(prev, run)) prev.text += run.text
    else merged.push({ ...run })
  }
  return merged
}

function sameStyle(a: Run, b: Run): boolean {
  // reference markers, index entries, cross-references and inline math are atomic; never merge
  if (a.noteRef || b.noteRef || a.xeTerm !== undefined || b.xeTerm !== undefined) return false
  if (a.refField !== undefined || b.refField !== undefined) return false
  if (a.instrField !== undefined || b.instrField !== undefined) return false
  if (a.math || b.math) return false
  if (a.ruby || b.ruby) return false
  if (a.image || b.image) return false
  return (
    (a.rawRPr ?? '') === (b.rawRPr ?? '') &&
    a.styleId === b.styleId &&
    !!a.cs === !!b.cs &&
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline &&
    !!a.strike === !!b.strike &&
    a.color === b.color &&
    a.sizeHalfPoints === b.sizeHalfPoints &&
    a.font === b.font &&
    a.fontAscii === b.fontAscii &&
    a.csFont === b.csFont &&
    a.highlight === b.highlight &&
    a.vertAlign === b.vertAlign &&
    (a.link?.href ?? '') === (b.link?.href ?? '') &&
    (a.link?.rId ?? '') === (b.link?.rId ?? '') &&
    (a.commentIds ?? []).join(',') === (b.commentIds ?? []).join(',') &&
    sameRevision(a.ins, b.ins) &&
    sameRevision(a.del, b.del)
  )
}

function sameRevision(a: RevisionInfo | undefined, b: RevisionInfo | undefined): boolean {
  if (!a || !b) return !a === !b
  return a.author === b.author && a.date === b.date && a.id === b.id
}

/** ST_Shd pct values that are not the literal digits in the name */
const SHD_PCT_EXACT: Record<string, number> = { pct12: 12.5, pct37: 37.5, pct62: 62.5, pct87: 87.5 }

function blendHex(fg: string, bg: string, ratio: number): string {
  const ch = (hex: string, i: number) => parseInt(hex.slice(i, i + 2), 16)
  const mix = (i: number) =>
    Math.round(ch(fg, i) * ratio + ch(bg, i) * (1 - ratio))
      .toString(16)
      .padStart(2, '0')
  return (mix(0) + mix(2) + mix(4)).toUpperCase()
}

/**
 * w:shd → single display color. Patterns (pctNN, stripes, crosses) approximate
 * as the pattern color blended over the fill at the pattern's ink coverage —
 * Word's actual dot/stripe raster is out of scope for cell backgrounds.
 */
export function shdDisplayFill(shd: XNode | undefined): string | undefined {
  if (!shd) return undefined
  const a = attrsOf(shd)
  const hex = (v: string | undefined) => {
    const s = v ? stripHash(v) : undefined
    return s && /^[0-9a-fA-F]{6}$/.test(s) ? s : undefined
  }
  const fill = a['w:fill'] === 'auto' ? undefined : hex(a['w:fill'])
  const val = a['w:val'] ?? 'clear'
  if (val === 'clear' || val === 'nil') return fill
  const ink = a['w:color'] === 'auto' ? undefined : hex(a['w:color'])
  if (val === 'solid') return ink ?? '000000'
  let ratio: number | undefined
  if (val.startsWith('pct')) {
    ratio = (SHD_PCT_EXACT[val] ?? Number(val.slice(3))) / 100
    if (!(ratio > 0 && ratio <= 1)) ratio = undefined
  } else if (/stripe|cross/i.test(val)) {
    ratio = val.startsWith('thin') ? 0.25 : 0.5
  }
  if (ratio === undefined) return fill
  return blendHex(ink ?? '000000', fill ?? 'FFFFFF', ratio)
}

export function tableLookOf(tblPr: XNode | undefined): NonNullable<TableModel['tableLook']> {
  const look = attrsOf(findChild(tblPr ?? {}, 'w:tblLook') ?? {})
  const bits = parseInt(look['w:val'] ?? '', 16)
  const flag = (attr: string, bit: number, dflt: boolean): boolean =>
    look[attr] !== undefined
      ? look[attr] !== '0' && look[attr] !== 'false'
      : Number.isFinite(bits)
        ? (bits & bit) !== 0
        : dflt
  return {
    firstRow: flag('w:firstRow', 0x20, true),
    lastRow: flag('w:lastRow', 0x40, false),
    firstColumn: flag('w:firstColumn', 0x80, true),
    lastColumn: flag('w:lastColumn', 0x100, false),
    bandedRows: !flag('w:noHBand', 0x200, false),
    bandedColumns: !flag('w:noVBand', 0x400, true),
  }
}

function borderLinesOf(node: XNode | undefined, withInside: true): TableBorders | undefined

function borderLinesOf(node: XNode | undefined, withInside: false): CellBorders | undefined

function borderLinesOf(node: XNode | undefined, withInside: boolean): TableBorders | undefined {
  if (!node) return undefined
  const ALIAS: Record<string, keyof TableBorders> = {
    'w:top': 'top',
    'w:left': 'left',
    'w:bottom': 'bottom',
    'w:right': 'right',
    'w:start': 'left',
    'w:end': 'right',
    ...(withInside ? { 'w:insideH': 'insideH', 'w:insideV': 'insideV' } : {}),
  }
  const borders: TableBorders = {}
  for (const [tag, side] of Object.entries(ALIAS)) {
    const child = findChild(node, tag)
    if (!child || borders[side]) continue
    const a = attrsOf(child)
    if (!a['w:val']) continue
    borders[side] = {
      style: a['w:val'],
      ...(a['w:sz'] ? { szEighths: Number(a['w:sz']) || undefined } : {}),
      ...(a['w:color'] ? { color: stripHash(a['w:color']) } : {}),
    }
  }
  return Object.keys(borders).length > 0 ? borders : undefined
}

/** Duplicated border containers (two w:tcBorders in one tcPr etc.): Word merges per side, later wins */
export function mergedBorderLinesOf(
  parent: XNode | undefined,
  tag: string,
  withInside: true,
): TableBorders | undefined

export function mergedBorderLinesOf(
  parent: XNode | undefined,
  tag: string,
  withInside: false,
): CellBorders | undefined

export function mergedBorderLinesOf(
  parent: XNode | undefined,
  tag: string,
  withInside: boolean,
): TableBorders | undefined {
  if (!parent) return undefined
  let merged: TableBorders | undefined
  for (const node of findChildren(parent, tag)) {
    const b = borderLinesOf(node, withInside as true)
    if (b) merged = { ...merged, ...b }
  }
  return merged
}

export function cellMarginsOf(node: XNode | undefined): CellMargins | undefined {
  if (!node) return undefined
  const SIDES: Array<[string, keyof CellMargins]> = [
    ['w:top', 'top'],
    ['w:left', 'left'],
    ['w:bottom', 'bottom'],
    ['w:right', 'right'],
    ['w:start', 'left'],
    ['w:end', 'right'],
  ]
  const m: CellMargins = {}
  for (const [tag, side] of SIDES) {
    const a = attrsOf(findChild(node, tag) ?? {})
    if (a['w:type'] && a['w:type'] !== 'dxa') continue
    const v = Number(a['w:w'])
    if (Number.isFinite(v) && v >= 0 && m[side] === undefined) m[side] = v
  }
  return Object.keys(m).length > 0 ? m : undefined
}

/** Word's substitution face when the East Asian font slot is empty, by w:lang w:eastAsia */
export const EA_LANG_DEFAULT_FONT: Record<string, string> = {
  // Word probe + fontTable: ko-KR empty EA slot substitutes Malgun Gothic, not Batang
  ko: 'Malgun Gothic',
  'ko-kr': 'Malgun Gothic',
  ja: 'MS Mincho',
  'ja-jp': 'MS Mincho',
  'zh-cn': 'SimSun',
  'zh-tw': 'PMingLiU',
  'zh-hk': 'PMingLiU',
}
