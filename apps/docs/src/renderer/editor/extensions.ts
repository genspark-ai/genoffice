import { Editor, Extension, Node } from '@tiptap/core'
import type { ChainedCommands, RawCommands } from '@tiptap/core'
import { UndoRedo } from '@tiptap/extensions'
import { DOMSerializer } from '@tiptap/pm/model'
import type { DOMOutputSpec, Node as PmNode } from '@tiptap/pm/model'
import {
  AllSelection,
  NodeSelection,
  Plugin,
  PluginKey,
  TextSelection,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import {
  CellSelection,
  addRowAfter,
  columnResizing,
  columnResizingPluginKey,
  deleteTable,
  goToNextCell,
  isInTable,
  tableEditing,
} from '@tiptap/pm/tables'
import {
  autospaceBoundaries,
  autospacePadBetween,
  codePointLengthAt,
  cjkDeclaredLineFactor,
  cjkScriptRanges,
  cssAutoLineMult,
  cssFontFamily,
  cssRunFontFamily,
  cssGridSpacingPt,
  cssLineHeight,
  isCjkFontName,
  lineHeightFactor,
  paraLineFactorCss,
  SIMSUN_GAP_CHAR_RE,
  simsunGapLineFactor,
  strutFontCss,
  textHasCjk,
  justifySymbolRanges,
  textHasHangul,
  textHasLatinInk,
  cssSimsunGapLineExpr,
  WORD_AUTO_SPACING_PT,
} from '../line-metrics'
import { noteMarkText } from '../note-format'
import { t } from '../i18n/locale'
import {
  ommlToMathML,
  patchMathTokens,
  type ChartDisplay,
  type DiagramDisplay,
  type DocDefaults,
  type FieldDisplay,
  type FormulaDisplay,
  type NewChart,
  type NumberingDef,
  type NumberingLevel,
  type Run,
  type StrayIndent,
  type StyleDisplay,
  type StyleInfo,
  type TableCell,
  type TableModel,
  type TextboxDisplay,
  type TextboxListMarker,
} from '@genoffice/docx-engine'
import {
  bulletMarkerScale,
  computeListMarkerInfos,
  markerTabAdvance,
  type ListItemRef,
} from './numbering'
import { symbolFontCovers } from '../font-check'
import { dropActiveSubEditor, notifySubEditorState, setActiveSubEditor } from './active-editor'
import { borderLineCss, paraBorderCss } from './hf-dom'

export { borderLineCss }
import {
  INLINE_RULE_CLASS,
  inlineRuleDecls,
  parseInlineRuleEl,
  type InlineRule,
} from './inline-rule'
import { DK_SIDE, dkBackground, dkBorder, dkColor, dkTableBorders } from './dark-page'
import { fillInk } from './shading-ink'
import {
  FloatVShiftsExtension,
  PaginationGapsExtension,
  RowFillsExtension,
} from './pagination-gaps'
import { CaretMarksMemory, FORMAT_MARKS, firstTextMarksIn, serializeMarks } from './caret-marks'
import { insertPageBreak } from './page-break'
import { ColumnLayoutExtension } from './column-layout'
import { TableHandle } from './table-handle'
import { TrackChangesExtension } from './revisions'
import { inlineToRuns, runsToInline, textboxParaSignature, type PmNode as PmJson } from './convert'
import { constrainTableWidthAtCell } from './table-sizing'

/**
 * Custom schema mirroring the docx-engine Block model 1:1.
 * Every top-level node carries `docxIndex` (patch anchor, null = new) and
 * `aiChanged` (diff highlighting for AI edits).
 */

import {
  CHART_MAX_WIDTH_PX,
  CHART_TITLE_ROW_PX,
  drawChartSvg,
  renderChartSpec,
  renderFieldSpec,
  renderFormulaSpec,
  renderTableSpec,
  renderTextboxSpec,
  runSpanSpecs,
  textboxBoxStyle,
  textboxIsFilled,
  wireChartEditing,
} from './protected-render'
import { isStraightLineKind } from './shape-svg'
import {
  BoldMark,
  CommentMark,
  DelMark,
  InsMark,
  InstrFieldMark,
  ItalicMark,
  LinkMark,
  RefFieldMark,
  RevisionOriginalExtension,
  RprChangeMark,
  StrikeMark,
  TextStyleMark,
  UnderlineMark,
} from './marks'
import {
  DropCapExtension,
  EaHintQuotesExtension,
  MoveRevisionExtension,
  PPrChangeExtension,
  ParaBorderMergeExtension,
  ParaMarkDelExtension,
  PendingCommentHighlightExtension,
  ResolvedCommentsExtension,
  SdtExtension,
  SearchHighlightExtension,
  TabStopExtension,
  WsRunLineHeightExtension,
} from './decoration-extensions'
import { JustifyShrinkExtension } from './justify-shrink'
import { CjkPunctShrinkExtension } from './cjk-punct-shrink'
import { AutoDirectionExtension } from './direction'
import { InactiveSelectionExtension } from './inactive-selection'
import { AiQueueAnchorsExtension } from './ai-queue-anchors'
import { PageGapNavExtension } from './page-gap-nav'
import { moveBlocks } from './move-block'
import {
  foldQuarterTurnMargins,
  pictureTransformFns,
  quarterTurnInsetPx,
  quarterTurnMarginCss,
} from './image-rotation'
export * from './marks'
export * from './decoration-extensions'

const anchorAttrs = {
  docxIndex: { default: null as number | null },
  styleId: { default: null as string | null },
  aiChanged: { default: false },
  /** user bookmark names starting in this paragraph */
  bookmarks: { default: null as string[] | null },
  /** Word internal bookmarks (_Ref/_Toc…): kept out of the bookmark manager, written back verbatim on paragraph rebuild so cross-references don't break */
  hiddenBookmarks: { default: null as string[] | null },
  /** Endpoints of cross-paragraph comment ranges (only one end in this paragraph); written back on paragraph rebuild to avoid orphan marks */
  commentStarts: { default: null as string[] | null },
  commentEnds: { default: null as string[] | null },
  align: { default: null as string | null },
  lineSpacing: { default: null as number | null },
  /** Word line-spacing rule (auto/atLeast/exact); null = auto */
  lineRule: { default: null as string | null },
  /** Raw w:spacing w:line twips (used by atLeast/exact) */
  lineRawTwips: { default: null as number | null },
  /** w:snapToGrid: false only when explicitly off (opts out of docGrid snapping) */
  snapToGrid: { default: null as boolean | null },
  indentLeft: { default: null as number | null },
  indentRight: { default: null as number | null },
  indentFirstLine: { default: null as number | null },
  spaceBefore: { default: null as number | null },
  spaceAfter: { default: null as number | null },
  /** w:beforeAutospacing / w:afterAutospacing: Word's HTML auto spacing (14pt) replaces
      the literal; false = explicit "0" overriding a style-chain auto */
  spaceBeforeAuto: { default: null as boolean | null },
  spaceAfterAuto: { default: null as boolean | null },
  /** w:contextualSpacing on the pPr itself; false = explicit off overriding the style */
  contextualSpacing: { default: null as boolean | null },
  // never copied to the second half of an Enter split: Word's page break is a
  // character before the paragraph content, so a newline must not clone the
  // break onto the new paragraph (alpha ledger r157)
  pageBreakBefore: { default: false, keepOnSplit: false },
  /** RTL paragraph (w:bidi); align is already the visual value */
  bidi: { default: false },
  /** render-only RTL inferred from run w:rtl / RTL script when w:bidi is absent
      (HTML-converted docs); align is the visual value as-is, never saved */
  bidiInferred: { default: false },
  /** CJK-Latin/digit auto spacing (w:autoSpaceDE/DN); null = Word default on */
  autoSpace: { default: null as boolean | null },
  /** w:wordWrap; false = hangul words may break between syllables (Word's allow-word-split option) */
  wordWrap: { default: null as boolean | null },
  /** w:overflowPunct; false = a line-end CJK stop never hangs into the margin */
  overflowPunct: { default: null as boolean | null },
  /** paragraph-style chain w:lang w:eastAsia; runs without their own value inherit it */
  eaLang: { default: null as string | null },
  shadingFill: { default: null as string | null },
  /** display-only pattern-shading blend (w:shd pctNN); never saved back */
  shadingDisplay: { default: null as string | null },
  /** sides a direct w:pBdr resets to none: cancels the style-level border (display-only) */
  borderReset: { default: null as string | null },
  /** w:sz (half-points) of the paragraph mark / dropped empty runs; sizes the line of run-less paragraphs */
  emptyRunSize: { default: null as number | null },
  /** w:rFonts of the paragraph mark / dropped empty runs; faces the line of run-less paragraphs */
  emptyRunFont: { default: null as string | null },
  /** subset of "tblr": which sides have a single-line border */
  borders: { default: null as string | null },
  /** JSON per-side {color?,szPt?} for `borders` (w:pBdr declared look) */
  borderLines: { default: null as string | null },
  /** custom tab stops JSON: Array<{pos:number,val:string,leader?:string}> */
  tabStops: { default: null as string | null },
  /** drop cap: JSON {type:'drop'|'margin',lines:number} */
  dropCap: { default: null as string | null },
  /** SDT shell: JSON SdtShell (alias, tag, controlType, openXml, closeXml) */
  sdtShell: { default: null as string | null },
  /** move revision type: 'from' (content moved away) or 'to' (content moved here) */
  moveRevision: { default: null as string | null },
  /** JSON {author,date?,id?} when the paragraph has a pPrChange tracked format change */
  pPrChange: { default: null as string | null },
  /** JSON {author,date?,id?} when the paragraph mark is a tracked deletion (w:pPr/w:rPr/w:del) */
  paraMarkDel: { default: null as string | null },
  /** top-level insertion/deletion revision ({kind,author,date?}) */
  blockRevision: { default: null as Record<string, string> | null },
  /** JSON marks snapshot for an empty block's caret (Word's pilcrow
      formatting): stamped while the caret holds stored marks in an empty
      block, restored as storedMarks when the caret re-enters bare — arrow
      navigation must not drop the pending format (alpha ledger r114) */
  caretMarks: { default: null as string | null },
}

/** Max explicit run size (half-points, plus size uniformity) when *every* text child declares one, else null.
 *  Any run inheriting the body size keeps the inherited strut (conservative: never shrinks a line Word would keep tall). */
function explicitStrutHalfPoints(node: {
  descendants?: PmNode['descendants']
}): { halfPoints: number; uniform: boolean } | null {
  if (!node.descendants) return null
  let max: number | null = null
  let min: number | null = null
  let inherited = false
  node.descendants((child) => {
    if (inherited) return false
    if (!child.isText) return true
    const sz = child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.sizeHalfPoints as
      number | null | undefined
    if (sz == null) inherited = true
    else {
      max = Math.max(max ?? 0, sz)
      min = Math.min(min ?? Infinity, sz)
    }
    return false
  })
  return inherited || max === null ? null : { halfPoints: max, uniform: max === min }
}

/**
 * Latin paragraphs: runs declaring a font override the doc-level factor with that
 * face's metric (Word sizes lines by run fonts, not the docDefaults face — a
 * Cambria-themed doc whose runs all say Calibri lays out at 1.22, not 1.17);
 * runs inheriting the body face keep the doc var via max().
 */
function latinParaFactor(
  node: { descendants?: PmNode['descendants'] },
  scriptVar: string,
  latinInkOnly = false,
): string {
  if (!node.descendants) return scriptVar
  let declaredMax = 0
  let undeclared = false
  node.descendants((child) => {
    if (!child.isText) return true
    if (latinInkOnly && !textHasLatinInk(child.text ?? '')) return false
    const attrs = child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs
    // ascii/hAnsi slot only: Word lays Latin text with the (possibly inherited)
    // ascii face, so an eastAsia-only declaration (attrs.font) must not drag
    // its factor onto a Latin line (EA "Arial Unicode MS" = 1.74, sample 13)
    const family = attrs?.fontAscii as string | null | undefined
    if (family) declaredMax = Math.max(declaredMax, lineHeightFactor(family))
    else undeclared = true
    return false
  })
  if (declaredMax <= 0) return scriptVar
  return undeclared ? `max(${scriptVar}, ${declaredMax})` : String(declaredMax)
}

/** Paragraph font-family follows the runs when every text run declares one:
 *  Chromium's line box is the union of the strut (paragraph font) and run boxes,
 *  so a paragraph face with a taller ascent than the runs' inflates every line
 *  past the computed line-height. No run inherits the face, so only the strut
 *  (and list markers without their own font) changes. */
function paraDeclaredFontFamily(node: { descendants?: PmNode['descendants'] }): string | null {
  if (!node.descendants) return null
  let first: string | null = null
  let inherited = false
  node.descendants((child) => {
    if (inherited) return false
    if (!child.isText) return true
    const attrs = child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs
    const ea = attrs?.font ? String(attrs.font) : null
    const ascii = attrs?.fontAscii ? String(attrs.fontAscii) : null
    if (!ea && !ascii) {
      inherited = true
      return false
    }
    // same chain the run's span renders, so the strut face equals the run face
    first ??= cssRunFontFamily(ascii, ea)
    return false
  })
  return inherited ? null : first
}

/** Mixed paragraph: some runs declare a CJK-named face, others inherit. The
 *  strut cannot take the declared family wholesale (inherited Latin text would
 *  change face), so typed-grid docs align its geometry through the glyphless
 *  metrics face instead (doc-style-css .doc-grid-strut; inert elsewhere). */
function paraMixedDeclaredCjk(node: { descendants?: PmNode['descendants'] }): boolean {
  if (!node.descendants) return false
  let declaredCjk = false
  let inherited = false
  node.descendants((child) => {
    if (declaredCjk && inherited) return false
    if (!child.isText) return true
    const attrs = child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs
    const family = (attrs?.font ?? attrs?.fontAscii) as string | null | undefined
    if (!family) inherited = true
    else if (isCjkFontName(family)) declaredCjk = true
    return false
  })
  return declaredCjk && inherited
}

/**
 * Per-paragraph --doc-line-factor value: CJK runs with a declared font take that
 * font's LO-metric factor (max over runs); undeclared CJK runs keep the
 * document-level CJK var; non-CJK paragraphs keep the script-based guess.
 */
function paraLineFactor(node: {
  textContent?: string
  descendants?: PmNode['descendants']
}): string {
  const scriptVar = paraLineFactorCss(node.textContent ?? '')
  if (!node.descendants) return scriptVar
  if (!textHasCjk(node.textContent ?? '')) return latinParaFactor(node, scriptVar)
  let declaredMax = 0
  let undeclaredCjk = false
  node.descendants((child) => {
    if (!child.isText) return true
    if (!textHasCjk(child.text ?? '')) return false
    const mark = child.marks.find((m) => m.type.name === 'docTextStyle')
    // empty-EA-theme-slot backfills keep the Word-look face but are not a
    // document font choice: LO cascades such runs to the document's EA default,
    // so they count as undeclared here (the doc-level var carries that factor).
    // Latin-named faces likewise don't drive CJK line height — an ascii-only
    // "Times New Roman" run still renders its CJK via the inherited EA font.
    const family =
      mark?.attrs.eaSlotEmpty === true
        ? null
        : ((mark?.attrs.font ?? mark?.attrs.fontAscii) as string | null | undefined)
    if (family && isCjkFontName(family)) {
      declaredMax = Math.max(declaredMax, cjkDeclaredLineFactor(family) ?? lineHeightFactor(family))
    } else undeclaredCjk = true
    return false
  })
  if (declaredMax <= 0) return scriptVar
  return undeclaredCjk ? `max(${scriptVar}, ${declaredMax})` : String(declaredMax)
}

type RunFactorRange = { from: number; to: number; style: string }

/**
 * Word sizes every line by the faces actually on it (probe 2026-09-06: a
 * Latin-only line inside a Malgun Gothic paragraph is the Calibri line, and
 * KR/Latin neighbours sum to one line of each). A paragraph mixing CJK and
 * Latin ink therefore lays its strut at the Latin base and lifts only its CJK
 * stretches (inline decorations); Chromium's line box is the max of the inline
 * boxes on the line, so a line grows only when a CJK stretch lands on it.
 * Pure-CJK paragraphs keep the paragraph-level factor. exact lines never lift.
 */
function perLineFactors(node: PmNode): { strut: string; runs: RunFactorRange[] } | null {
  const text = node.textContent
  if (node.attrs.lineRule === 'exact' || !textHasCjk(text) || !textHasLatinInk(text)) return null
  const scriptVar = paraLineFactorCss(text)
  const lh = cssLineHeight(
    (node.attrs.lineRule as 'auto' | 'atLeast' | 'exact' | null) ?? undefined,
    node.attrs.lineRawTwips != null ? Number(node.attrs.lineRawTwips) : undefined,
    node.attrs.lineSpacing ? Number(node.attrs.lineSpacing) : undefined,
  )
  const runs: RunFactorRange[] = []
  let offset = 0
  node.forEach((child) => {
    if (child.isText && child.text) {
      const mark = child.marks.find((m) => m.type.name === 'docTextStyle')
      const family =
        mark?.attrs.eaSlotEmpty === true
          ? null
          : ((mark?.attrs.font ?? mark?.attrs.fontAscii) as string | null | undefined)
      const factor =
        family && isCjkFontName(family)
          ? String(cjkDeclaredLineFactor(family) ?? lineHeightFactor(family))
          : scriptVar
      const style = `--doc-line-factor:${factor}${lh ? `;line-height:${lh}` : ''}`
      for (const r of cjkScriptRanges(child.text)) {
        runs.push({ from: offset + r.from, to: offset + r.to, style })
      }
    }
    offset += child.nodeSize
  })
  return { strut: latinParaFactor(node, 'var(--doc-line-factor-latin,1.2)', true), runs }
}

/**
 * Paragraph FORMATTING attrs that survive the clipboard HTML round-trip
 * (renderHTML emits them as data-para JSON; parseHTML restores them), typed so
 * a crafted/corrupt payload can't smuggle wrong-typed values into the model.
 * Identity/anchor attrs stay out on purpose: a pasted paragraph is NEW content,
 * so docxIndex (save patch anchor), bookmarks, comment endpoints, sdtShell and
 * revision metadata must not be duplicated by copy/paste (alpha ledger r117).
 */
const CLIPBOARD_PARA_ATTR_TYPES: Record<string, 'string' | 'number' | 'boolean'> = {
  styleId: 'string',
  align: 'string',
  lineSpacing: 'number',
  lineRule: 'string',
  lineRawTwips: 'number',
  snapToGrid: 'boolean',
  indentLeft: 'number',
  indentRight: 'number',
  indentFirstLine: 'number',
  spaceBefore: 'number',
  spaceAfter: 'number',
  spaceBeforeAuto: 'boolean',
  spaceAfterAuto: 'boolean',
  contextualSpacing: 'boolean',
  pageBreakBefore: 'boolean',
  bidi: 'boolean',
  bidiInferred: 'boolean',
  autoSpace: 'boolean',
  wordWrap: 'boolean',
  overflowPunct: 'boolean',
  eaLang: 'string',
  shadingFill: 'string',
  shadingDisplay: 'string',
  borderReset: 'string',
  emptyRunSize: 'number',
  emptyRunFont: 'string',
  borders: 'string',
  borderLines: 'string',
  outlineOnly: 'boolean',
  tabStops: 'string',
  dropCap: 'string',
  // docListItem identity (paragraphs never set them; ProseMirror drops attrs
  // unknown to the parsing node type). numId stays out: a pasted copy pointing
  // at another document's numbering table would dangle.
  kind: 'string',
  ilvl: 'number',
}

/** data-para payload for a block node, or null when everything is at defaults */
function clipboardParaPayload(attrs: Record<string, unknown>): string | null {
  const clip: Record<string, unknown> = {}
  for (const key of Object.keys(CLIPBOARD_PARA_ATTR_TYPES)) {
    const v = attrs[key]
    if (v == null) continue
    // false IS meaningful for snapToGrid/autoSpace (explicit opt-out); for the
    // false-by-default flags it's just the default and stays out of the payload
    if (v === false && (key === 'pageBreakBefore' || key === 'bidi' || key === 'bidiInferred')) {
      continue
    }
    clip[key] = v
  }
  return Object.keys(clip).length > 0 ? JSON.stringify(clip) : null
}

/** Inverse of clipboardParaPayload for parseHTML getAttrs: restores only
 *  whitelisted keys whose runtime type matches, null when absent/malformed. */
export function clipboardParaAttrs(el: HTMLElement): Record<string, unknown> | null {
  const raw = el.getAttribute?.('data-para')
  if (!raw) return null
  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch {
    return null
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) return null
  const attrs: Record<string, unknown> = {}
  for (const [key, type] of Object.entries(CLIPBOARD_PARA_ATTR_TYPES)) {
    const v = (parsed as Record<string, unknown>)[key]
    if (typeof v !== type) continue
    if (type === 'number' && !Number.isFinite(v)) continue
    attrs[key] = v
  }
  return Object.keys(attrs).length > 0 ? attrs : null
}

function blockAttrs(
  node: {
    attrs: Record<string, unknown>
    textContent?: string
    childCount?: number
    descendants?: PmNode['descendants']
  },
  {
    includeIndent = true,
    listGeometry = false,
  }: { includeIndent?: boolean; listGeometry?: boolean } = {},
): Record<string, string> {
  const attrs: Record<string, string> = {}
  if (node.attrs.docxIndex !== null) attrs['data-idx'] = String(node.attrs.docxIndex)
  // paragraph formatting round-trips the clipboard through this attribute:
  // the CSS below renders it, but nothing parses that CSS back (r117)
  const clip = clipboardParaPayload(node.attrs)
  if (clip) attrs['data-para'] = clip
  // per-document style CSS (generated from styles.xml) targets this attribute
  if (node.attrs.styleId) attrs['data-style'] = String(node.attrs.styleId)
  // bookmark jump targets ([data-bookmarks~="name"]; names cannot contain spaces)
  if (Array.isArray(node.attrs.bookmarks) && node.attrs.bookmarks.length > 0) {
    attrs['data-bookmarks'] = (node.attrs.bookmarks as string[]).join(' ')
  }
  const classes: string[] = []
  if (node.attrs.aiChanged) classes.push('ai-changed')
  // textless paragraph: table-cell CSS shrinks these to the Latin line height
  // (an empty cell inheriting the CJK factor would out-grow the content cells)
  if (!node.textContent) classes.push('doc-p-empty')
  if (node.attrs.pageBreakBefore) {
    classes.push('page-break-before')
    attrs['data-page-break-label'] = t('editorPageBreak')
  }
  if (classes.length > 0) attrs['class'] = classes.join(' ')
  const styles: string[] = []
  if (node.attrs.bidi || node.attrs.bidiInferred) {
    styles.push('direction:rtl', 'unicode-bidi:isolate')
  } else {
    // explicit: paragraph direction is its own w:bidi only — a bidiVisual
    // table's dir="rtl" mirrors column order but must not reorder cell text
    styles.push('direction:ltr')
  }
  // explicit autoSpaceDE/DN off also disables the browser's native 1/8em gap
  if (node.attrs.autoSpace === false) styles.push('text-autospace:no-autospace')
  if (node.attrs.align) {
    styles.push(`text-align:${node.attrs.align === 'distribute' ? 'justify' : node.attrs.align}`)
  }
  // the line-height factor follows paragraph content (approximating Word's max-of-inline-fonts
  // line height): CJK paragraphs get the CJK factor, pure-Western ones the document's
  // font-aware Latin factor (doc-style-css sets --doc-line-factor-latin per body font).
  // Runs that DECLARE a font override the script guess with that font's factor
  // (LO probe: line height follows the requested face's metrics even when CJK
  // glyphs fall through to another font — EA "Times New Roman" lays at 1.15em);
  // the paragraph takes the max over its CJK runs, like Word's tallest-run rule.
  if (node.textContent) {
    // Word breaks Korean at spaces (UAX#14 default would break between syllables);
    // scoped to Hangul-bearing paragraphs so CJ text keeps per-char breaking.
    // overflow-wrap keeps the sim's overlong-word hard-break fallback.
    if (textHasHangul(node.textContent) && node.attrs.wordWrap !== false) {
      styles.push('word-break:keep-all', 'overflow-wrap:anywhere')
    }
    styles.push(`--doc-line-factor:${paraLineFactor(node)}`)
    const fam = paraDeclaredFontFamily(node)
    if (fam) styles.push(`font-family:${fam}`)
    else if (paraMixedDeclaredCjk(node)) classes.push('doc-grid-strut')
    // Word's line strut follows run sizes; without this the paragraph inherits the
    // body size (often larger than table-cell runs) and every line box inflates.
    // Mixed sizes shrink-only, a uniform run size sizes every line (strutFontCss)
    const strut = explicitStrutHalfPoints(node)
    if (strut) styles.push(...strutFontCss(strut))
  } else {
    // Word sizes an empty line by the paragraph mark / empty run, both directions
    if (node.attrs.emptyRunSize) styles.push(`font-size:${Number(node.attrs.emptyRunSize) / 2}pt`)
    // the mark face sizes the empty line, CJK included (empty-line probe
    // 2026-08-25: SimSun 1.30 / DengXian 1.36 / Malgun 1.74 / Calibri 1.24 —
    // each face's own text factor; the earlier Western-only scoping starved
    // CJK marks down to the document factor and doubled their grid rows)
    const fam = node.attrs.emptyRunFont ? String(node.attrs.emptyRunFont) : null
    if (fam) {
      styles.push(`--doc-line-factor:${lineHeightFactor(fam)}`, `font-family:${cssFontFamily(fam)}`)
    } else if (!node.childCount) {
      // a mark without its own w:rFonts lays the line with the inherited ascii
      // face (Word probe 2026-09-05), not the document's CJK root factor
      styles.push('--doc-line-factor:var(--doc-line-factor-latin,1.2)')
    }
  }
  const lineRule = (node.attrs.lineRule as 'auto' | 'atLeast' | 'exact' | null) ?? undefined
  const lineRawTwips = node.attrs.lineRawTwips != null ? Number(node.attrs.lineRawTwips) : undefined
  const lineSpacing = node.attrs.lineSpacing ? Number(node.attrs.lineSpacing) : undefined
  const lh = cssLineHeight(lineRule, lineRawTwips, lineSpacing)
  if (lh) styles.push(`line-height:${lh}`)
  // grid-doc span snapping (doc-style-css): fixed-height lines opt out (atLeast
  // never snaps, including the line="0" opt-out form), multiples scale
  if ((lineRule === 'exact' && lineRawTwips) || lineRule === 'atLeast') {
    classes.push('doc-lh-fixed')
    attrs['class'] = classes.join(' ')
  } else {
    const mult =
      lineSpacing ?? (lineRule === 'auto' && lineRawTwips ? lineRawTwips / 240 : undefined)
    // explicit single (mult 1) still overrides an inherited style/doc multiple
    if (mult) styles.push(`--doc-line-mult:${mult}`)
  }
  // w:snapToGrid=0: opt this paragraph out of docGrid line snapping (the
  // round(up) expressions read the pitch var, so a local ~0 disables them;
  // .doc-nosnap re-declares --doc-line-max as natural x mult on the paragraph
  // AND its spans — an inline var would not reach spans, which .doc-page *
  // re-declares per element)
  if (node.attrs.snapToGrid === false) {
    styles.push('--doc-grid-pitch:0.0001px')
    classes.push('doc-nosnap')
    attrs['class'] = classes.join(' ')
  }
  // list items already indent via padding; a margin would double-shift them.
  // listGeometry: drive list geometry with w:ind (--li-left text indent, --li-hang the hanging
  // area i.e. the number-marker width); negative text-indent is expressed by the marker box, no longer emitted directly
  // Logical (inline-start/end) margins: identical to left/right in LTR, and in bidi
  // paragraphs they mirror, matching Word's quirk that w:ind left/right swap sides.
  // explicit 0 still emits: it beats the style indent, the 0.55in list default
  // and the numbering level (Word merges w:ind per attribute, direct value first)
  if (includeIndent && node.attrs.indentLeft != null) {
    styles.push(`margin-inline-start:${Number(node.attrs.indentLeft) / 20}pt`)
  } else if (listGeometry && node.attrs.indentLeft != null) {
    const left = Number(node.attrs.indentLeft) / 20
    // padding cannot go negative: a negative start indent moves the box instead
    if (left < 0) styles.push(`margin-inline-start:${left}pt`, '--li-left:0pt')
    else styles.push(`--li-left:${left}pt`)
  }
  if (node.attrs.indentRight != null)
    styles.push(`margin-inline-end:${Number(node.attrs.indentRight) / 20}pt`)
  if (node.attrs.indentFirstLine != null) {
    const firstLine = Number(node.attrs.indentFirstLine)
    if (listGeometry && firstLine < 0) styles.push(`--li-hang:${-firstLine / 20}pt`)
    else if (!listGeometry || firstLine > 0) styles.push(`text-indent:${firstLine / 20}pt`)
  }
  // explicit 0 must still emit (w:after="0" overrides the style/docDefaults margin)
  // autospacing replaces the literal with Word's HTML auto value (14pt, measured);
  // the sp-auto-* classes let CSS collapse it to 0 between two list items (Word)
  if (node.attrs.spaceBeforeAuto) {
    styles.push(`margin-top:${cssGridSpacingPt(WORD_AUTO_SPACING_PT)}`)
    classes.push('sp-auto-b')
  } else if (node.attrs.spaceBefore != null)
    styles.push(`margin-top:${cssGridSpacingPt(Number(node.attrs.spaceBefore) / 20)}`)
  if (node.attrs.spaceAfterAuto) {
    styles.push(`margin-bottom:${cssGridSpacingPt(WORD_AUTO_SPACING_PT)}`)
    classes.push('sp-auto-a')
  } else if (node.attrs.spaceAfter != null)
    styles.push(`margin-bottom:${cssGridSpacingPt(Number(node.attrs.spaceAfter) / 20)}`)
  // direct w:contextualSpacing: same-style adjacency suppression / explicit opt-out
  // (the style-level rules live in doc-style-css and honor these classes)
  if (node.attrs.contextualSpacing === true) classes.push('ctx-sp')
  else if (node.attrs.contextualSpacing === false) classes.push('ctx-sp-off')
  if (classes.length > 0) attrs['class'] = classes.join(' ')
  const shdBg = node.attrs.shadingDisplay ?? node.attrs.shadingFill
  // authored colors stay the declaration; the --dk-* twins feed the dark page (dark-page.ts)
  if (shdBg) styles.push(`background-color:#${shdBg}`, dkBackground(`#${shdBg}`))
  const ink = fillInk(shdBg ? String(shdBg) : null)
  if (ink) attrs['data-ink'] = ink
  if (node.attrs.borders) {
    const borders = String(node.attrs.borders)
    let borderLines: Partial<Record<string, { color?: string; szPt?: number }>> = {}
    if (node.attrs.borderLines) {
      try {
        borderLines = JSON.parse(String(node.attrs.borderLines))
      } catch {
        /* malformed attr: fall back to legacy line */
      }
    }
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const key = DK_SIDE[side]
      if (!borders.includes(key)) continue
      const line = paraBorderCss(borderLines[key])
      styles.push(`border-${side}:${line}`, dkBorder(key, line))
    }
    styles.push('padding:1px 4px')
  }
  if (node.attrs.borderReset) {
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      if (String(node.attrs.borderReset).includes(DK_SIDE[side])) styles.push(`border-${side}:none`)
    }
  }
  if (node.attrs.tabStops) {
    // debugging aid only; actual tab layout is measured by TabStopExtension
    attrs['data-tab-stops'] = String(node.attrs.tabStops)
  }
  if (node.attrs.dropCap) {
    attrs['data-drop-cap'] = String(node.attrs.dropCap)
  }
  if (node.attrs.sdtShell) {
    attrs['data-sdt'] = 'true'
  }
  if (node.attrs.moveRevision) {
    attrs['data-move-revision'] = String(node.attrs.moveRevision)
  }
  if (node.attrs.pPrChange) {
    attrs['data-ppr-change'] = 'true'
  }
  if (styles.length > 0) attrs['style'] = styles.join(';')
  return attrs
}

export const DocDocument = Node.create({
  name: 'doc',
  topNode: true,
  content: 'block+',
})

export const DocText = Node.create({
  name: 'text',
  group: 'inline',
})

/** Footnote / endnote reference marker: an atomic superscript number. */
export const DocNoteRef = Node.create({
  name: 'docNoteRef',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      kind: { default: 'footnote' as 'footnote' | 'endnote' },
      id: { default: '' },
      num: { default: 1 },
    }
  },
  parseHTML() {
    return [{ tag: 'sup[data-note-ref]' }]
  },
  renderHTML({ node }) {
    return [
      'sup',
      {
        'data-note-ref': String(node.attrs.id),
        'data-note-kind': String(node.attrs.kind),
        class: 'doc-note-ref',
        title: node.attrs.kind === 'footnote' ? t('editorFootnote') : t('editorEndnote'),
      },
      // bare superscript number, matching Word; hover/selection accents live in CSS
      noteMarkText(node.attrs.kind as 'footnote' | 'endnote', Number(node.attrs.num) || 1),
    ]
  },
})

/** Index entry (XE field) marker: invisible in Word, a small chip on screen. */
export const DocXeMark = Node.create({
  name: 'docXeMark',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  addAttributes() {
    return { term: { default: '' } }
  },
  parseHTML() {
    return [{ tag: 'span[data-xe-term]' }]
  },
  renderHTML({ node }) {
    return [
      'span',
      {
        'data-xe-term': String(node.attrs.term),
        class: 'doc-xe-mark',
        title: t('editorIndexEntry', { term: String(node.attrs.term) }),
      },
    ]
  },
})

/**
 * Phonetic guide (w:ruby): atomic, renders as a native <ruby> element.
 * `xml` is the exact <w:ruby> fragment that saves verbatim.
 */
export const DocRuby = Node.create({
  name: 'docRuby',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      base: { default: '' },
      rt: { default: '' },
      xml: { default: '' },
    }
  },
  parseHTML() {
    return [{ tag: 'ruby[data-doc-ruby]' }]
  },
  renderText({ node }) {
    return String(node.attrs.base ?? '')
  },
  renderHTML({ node }) {
    return [
      'ruby',
      { 'data-doc-ruby': 'true', class: 'doc-ruby' },
      String(node.attrs.base),
      ['rt', {}, String(node.attrs.rt)],
    ]
  },
})

/**
 * Inline picture in a table cell (document content: sizes stay in px, no theme
 * tokens). `xml` is the exact <w:drawing> fragment that saves verbatim.
 */
export const DocInlineImage = Node.create({
  name: 'docInlineImage',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      dataUrl: { default: '' },
      widthPx: { default: null as number | null },
      heightPx: { default: null as number | null },
      xml: { default: '' },
      /** floating (wp:anchor) picture: wrap kind + anchor offsets (display only) */
      wrap: { default: null as string | null },
      offsetXEmu: { default: null as number | null },
      offsetYEmu: { default: null as number | null },
      wrapDistTopEmu: { default: null as number | null },
      wrapDistBottomEmu: { default: null as number | null },
      wrapDistLeftEmu: { default: null as number | null },
      wrapDistRightEmu: { default: null as number | null },
      /** picture outline (pic:spPr a:ln solid fill, display-only) */
      border: { default: null as { color: string; widthPt: number } | null },
      /** positionV line/center: the picture centers on its anchor line (display only) */
      lineCenterV: { default: false },
      /** pic a:xfrm rot (deg clockwise) / mirror flips (display only; the xml saves verbatim) */
      rotDeg: { default: null as number | null },
      flipH: { default: false },
      flipV: { default: false },
      /** the run's raw <w:rPr> slice, re-emitted before the fragment on save (not rendered) */
      rawRPr: { default: null as string | null, rendered: false },
      /** VML horizontal rule (alone or sharing the paragraph): {colorHex?, thicknessPx?, sizeHalfPoints?, widthPx?, align?}; drawn as a rule on its own line */
      rule: { default: null as InlineRule | null },
      /** the rule opens a paragraph that goes on with content: compact rule line (Word) */
      leadRule: { default: false },
    }
  },
  parseHTML() {
    return [
      { tag: 'img[data-inline-image]' },
      { tag: 'span[data-inline-rule]', getAttrs: (el) => parseInlineRuleEl(el as HTMLElement) },
    ]
  },
  renderHTML({ node }) {
    const rule = node.attrs.rule as InlineRule | null
    if (rule) {
      const ruleAttrs: Record<string, string> = {
        class: INLINE_RULE_CLASS,
        'data-inline-rule': JSON.stringify(rule),
        'data-xml': String(node.attrs.xml ?? ''),
      }
      if (node.attrs.leadRule) ruleAttrs['data-lead-rule'] = '1'
      const decls = inlineRuleDecls(rule)
      if (decls.length > 0) ruleAttrs.style = decls.join(';')
      return ['span', ruleAttrs]
    }
    const attrs: Record<string, string> = {
      'data-inline-image': '1',
      class: 'doc-inline-img',
      src: String(node.attrs.dataUrl),
    }
    const w = Number(node.attrs.widthPx)
    const h = Number(node.attrs.heightPx)
    const styles: string[] = []
    if (w > 0) styles.push(`width:${w}px`, h > 0 ? `height:${h}px` : 'height:auto')
    const ib = node.attrs.border as { color: string; widthPt: number } | null
    if (ib) {
      // picture outline (document data, not chrome)
      styles.push(
        `border:${((ib.widthPt * 96) / 72).toFixed(1)}px solid #${String(ib.color).replace(/[^0-9A-Fa-f]/g, '')}`,
      )
    }
    const xf = pictureTransformFns(node.attrs.rotDeg, node.attrs.flipH, node.attrs.flipV)
    const px = (v: number) => `${v.toFixed(1)}px`
    const wrap = node.attrs.wrap as string | null
    if (wrap === 'front' || wrap === 'behind') {
      // no-wrap / behind-text anchor: absolute overlay at the anchor offset from
      // the hosting paragraph's origin (styles.css positions that paragraph;
      // zero flow footprint, like Word); behind approximated by reduced
      // opacity — the flowing canvas has no z layer under the text
      const left = Number(node.attrs.offsetXEmu ?? 0) / EMU_PER_PX
      const top = Number(node.attrs.offsetYEmu ?? 0) / EMU_PER_PX
      // positive offsets ride on transform, not left/top: an off-page anchor
      // position must not create layout overflow — print layout hoists
      // absolutely positioned boxes to the page fragmentainer where no
      // ancestor clip applies, and Chromium's fit-to-paper shrink then scales
      // the whole export down. Negative offsets stay in layout: pulling a
      // page-wide picture's static box right would create that same overflow.
      const lx = Math.min(left, 0)
      const ty = Math.min(top, 0)
      styles.push(
        `position:absolute;left:${px(lx)};top:${px(ty)}`,
        `transform:translate(${px(left - lx)}, ${px(top - ty)})${xf.length ? ` ${xf.join(' ')}` : ''}`,
        'max-width:none',
      )
      attrs.style = styles.join(';')
      if (wrap === 'behind') attrs.class += ' doc-inline-img--behind'
      return [
        'span',
        { class: 'doc-inline-img-anchor', 'data-inline-image-anchor': '1' },
        ['img', attrs],
      ]
    }
    if (xf.length) styles.push(`transform:${xf.join(' ')}`)
    // the picture keeps its extent box (turned about its centre) while the
    // flow reserves the quarter-turned bounding box; floats keep their own
    // side pinned and grow toward the text
    const qt = quarterTurnInsetPx(w, h, node.attrs.rotDeg)
    // topBottom is a centred block: its visual centre is the box centre, so
    // only the stylesheet's 2px band gap needs carrying
    const bandGap = wrap === 'topBottom' ? 2 : 0
    const margin: Record<'top' | 'right' | 'bottom' | 'left', string | null> = {
      top: qt ? px(qt + bandGap) : null,
      right: qt && !wrap ? px(-qt) : null,
      bottom: qt ? px(qt + bandGap) : null,
      left: qt && !wrap ? px(-qt) : null,
    }
    // square/tight/through wrap → real CSS float so the surrounding text wraps;
    // topBottom → block line of its own
    if (wrap) attrs.class += ` doc-inline-img--wrap-${wrap}`
    // free-position floats honor the numeric posOffset X like the block-image
    // path: X measures from the column start; right floats convert it to a
    // right-edge inset so the picture is not stuck flush against the margin
    const tx = node.attrs.offsetXEmu != null ? Number(node.attrs.offsetXEmu) / EMU_PER_PX : null
    if (tx != null && wrap) {
      if (wrap.endsWith('-right') && w > 0) {
        margin.right = `calc(100% - ${px(tx + w)})`
        styles.push('max-width:none')
      } else if (wrap.endsWith('-left')) {
        margin.left = px(tx)
        styles.push('max-width:none')
      }
    }
    let sliverGuard: string | null = null
    if (wrap) {
      const distancePx = (attr: string): number | null =>
        node.attrs[attr] != null ? Number(node.attrs[attr]) / EMU_PER_PX : null
      const top = distancePx('wrapDistTopEmu')
      const bottom = distancePx('wrapDistBottomEmu')
      const left = distancePx('wrapDistLeftEmu')
      const right = distancePx('wrapDistRightEmu')
      if (top != null) margin.top = px(top + qt)
      if (bottom != null) margin.bottom = px(bottom + qt)
      // the stylesheet's 12px text-side gap stands in for a missing distance
      if (wrap.endsWith('-right') && (left != null || qt)) margin.left = px((left ?? 12) - qt)
      if (wrap.endsWith('-left') && (right != null || qt)) margin.right = px((right ?? 12) - qt)
      if (tx != null && wrap !== 'topBottom') {
        sliverGuard = wrapSliverGuardCss(wrap, tx + qt, qt ? h : w, right, qt)
      }
      // paragraph-relative posOffset Y: the float (and its text exclusion)
      // starts that far below the anchor line; position wins over distT
      const ty = node.attrs.offsetYEmu != null ? Number(node.attrs.offsetYEmu) / EMU_PER_PX : 0
      if ((ty !== 0 || qt) && wrap !== 'topBottom') {
        // the offset places the extent box, whose unturned top is the img box
        // top; the turned visual (Word's exclusion) starts an inset above it
        margin.top = px(ty)
        // a float excludes its whole margin box: without a shape the lines
        // above the picture (Word lays them at full width) would shorten too
        if (ty - qt > 0) styles.push(`shape-outside:inset(${px(ty - qt)} 0 0 0)`)
      }
    }
    // positionV line/center: lift so the picture centers on the anchor line
    // (0.75em ≈ half a single-spaced line) instead of hanging below it
    if (node.attrs.lineCenterV && h > 0) margin.top = `calc(0.75em - ${px(h / 2)})`
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      if (margin[side] != null) styles.push(`margin-${side}:${margin[side]}`)
    }
    if (sliverGuard) styles.push(sliverGuard)
    if (styles.length) attrs.style = styles.join(';')
    return ['img', attrs]
  },
})

/**
 * Atomic inline formula flowing with the text. `omml` is the exact <m:oMath>
 * fragment that saves verbatim; `mathml` renders natively in Chromium;
 * `latex` is kept for editor-created formulas so double-click can re-edit.
 */
export const DocInlineMath = Node.create({
  name: 'docInlineMath',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      omml: { default: '' },
      mathml: { default: '' },
      latex: { default: null as string | null },
      /** flat token strip (word count / AI read fallback) */
      text: { default: '' },
    }
  },
  parseHTML() {
    return [{ tag: 'span[data-inline-math]' }]
  },
  renderText({ node }) {
    return String(node.attrs.text ?? '')
  },
  renderHTML({ node }) {
    return ['span', inlineMathDomAttrs(node), String(node.attrs.text ?? '')]
  },
  addNodeView() {
    return ({ node, getPos }) => {
      let currentNode = node
      const dom = document.createElement('span')
      const render = () => {
        for (const [key, value] of Object.entries(inlineMathDomAttrs(currentNode))) {
          dom.setAttribute(key, value)
        }
        const mathml = String(currentNode.attrs.mathml ?? '')
        if (mathml) dom.innerHTML = mathml
        else dom.textContent = String(currentNode.attrs.text ?? '')
      }
      render()
      dom.addEventListener('dblclick', () => {
        // only editor-created formulas carry LaTeX and can be re-edited
        const latex = currentNode.attrs.latex
        const pos = (getPos as () => number | undefined)()
        if (latex && typeof pos === 'number') {
          window.dispatchEvent(
            new CustomEvent('ai-docs-edit-inline-math', {
              detail: { pos, latex: String(latex), kind: 'inline' },
            }),
          )
        }
      })
      return {
        dom,
        update: (n: PmNode) => {
          if (n.type.name !== 'docInlineMath') return false
          currentNode = n
          render()
          return true
        },
      }
    }
  },
})

function inlineMathDomAttrs(node: { attrs: Record<string, unknown> }): Record<string, string> {
  return {
    'data-inline-math': '1',
    class: 'doc-inline-math',
    title: node.attrs.latex
      ? t('editorEquationEditHint', { latex: String(node.attrs.latex) })
      : t('editorEquation'),
  }
}

export const DocHardBreak = Node.create({
  name: 'hardBreak',
  inline: true,
  group: 'inline',
  selectable: false,
  linebreakReplacement: true,
  // a break-only run (page break with its own rPr) keeps its formatting as marks
  marks: '_',
  addAttributes() {
    return {
      // in-paragraph page break (w:br w:type="page"): the pagination engine turns the page at its line
      pageBreak: { default: false },
      // column break (w:br w:type="column"): next column, or next page in a single-column section
      colBreak: { default: false },
    }
  },
  parseHTML() {
    return [
      { tag: 'br.doc-page-br', attrs: { pageBreak: true } },
      { tag: 'br.doc-col-br', attrs: { colBreak: true } },
      { tag: 'br' },
    ]
  },
  renderHTML({ node }) {
    return node.attrs.pageBreak
      ? ['br', { class: 'doc-page-br' }]
      : node.attrs.colBreak
        ? ['br', { class: 'doc-col-br' }]
        : ['br']
  },
  addKeyboardShortcuts() {
    return {
      'Shift-Enter': () => this.editor.commands.insertContent({ type: 'hardBreak' }),
    }
  },
})

/**
 * Enter must replace ANY non-empty selection with a paragraph break (Word).
 * Two selection shapes broke the default chain (alpha ledger r125/r129):
 * - Ctrl+A's AllSelection: every command in the split chain declines
 *   (splitBlock needs a textblock-depth selection; AllSelection's ends sit at
 *   doc depth 0) — the press was a silent no-op.
 * - A cross-block TextSelection with BOTH ends at block starts (what
 *   Shift+Right from a paragraph end followed by Shift+Down produces):
 *   TipTap's one-shot splitBlock THROWS "Inserted content deeper than
 *   insertion position" mid-dispatch, so the press did nothing while plain
 *   typing over the same selection worked.
 * Replacing the one-shot with separate delete + split dispatches sidesteps
 * both: the delete settles the document (and block merge) first, the split
 * then runs on an ordinary caret.
 */
export const EnterReplacesSelection = Extension.create({
  name: 'enterReplacesSelection',
  addKeyboardShortcuts() {
    return {
      Enter: () => {
        const sel = this.editor.state.selection
        // Word keeps the formatting of the START of the replaced selection
        // for what is typed next; the split-off caret often sits in an
        // EMPTIED paragraph with no neighbor to inherit from, so the first
        // deleted character's marks must ride along explicitly (alpha
        // ledger r133: keyboard whole-line selections type in theme font)
        const carryMarks = () => {
          // first TEXT node in the replaced range (an AllSelection's $from
          // sits at doc depth 0, where nodeAt returns the block, not a run)
          const marks = firstTextMarksIn(sel.$from.doc, sel.from, sel.to)
          // formatting only: comment/link/ins/del must not reattach to new
          // typing (they are inclusive:false for the same reason — bugbot)
          const found = (marks ?? sel.$from.marks()).filter((mark) =>
            FORMAT_MARKS.has(mark.type.name),
          )
          if (found.length > 0) {
            const tr = this.editor.state.tr.setMeta('addToHistory', false)
            // Word keeps the pending format on BOTH empty paragraphs the
            // replace leaves behind. The caret-marks memory stamps only the
            // block the caret sits in (the second), so arrowing up to the
            // FIRST emptied line reverted typing to the theme font and the
            // font box to "(Body)" (alpha ledger r133 residual). The stamp
            // step must precede setStoredMarks: a doc-changing step resets
            // the transaction's stored marks.
            const $head = this.editor.state.selection.$from
            const boundary = $head.before($head.depth)
            const prev = this.editor.state.doc.resolve(boundary).nodeBefore
            if (prev && prev.isTextblock && prev.content.size === 0 && 'caretMarks' in prev.attrs) {
              tr.setNodeMarkup(boundary - prev.nodeSize, undefined, {
                ...prev.attrs,
                caretMarks: serializeMarks(found),
              })
            }
            tr.setStoredMarks([...found])
            this.editor.view.dispatch(tr)
          }
        }
        if (sel instanceof AllSelection) {
          // an AllSelection maps to itself through the delete (it always
          // spans the whole doc), so hand the split a real caret explicitly
          this.editor.commands.deleteSelection()
          this.editor.chain().setTextSelection(1).splitBlock().run()
          carryMarks()
          return true
        }
        if (sel.empty || !(sel instanceof TextSelection)) return false
        if (sel.$from.sameParent(sel.$to)) return false // default chain is fine
        this.editor.commands.deleteSelection()
        this.editor.commands.splitBlock()
        carryMarks()
        return true
      },
    }
  },
})

/**
 * Word's insert-and-move chords that must only fire with editor focus. Kept
 * out of the application menu on purpose: a menu accelerator would
 * swallow Cmd+Enter from the renderer inputs that submit with it (comments
 * panel, prompt modal).
 */
export const WordEditorShortcuts = Extension.create({
  name: 'wordEditorShortcuts',
  addKeyboardShortcuts() {
    return {
      // Enter inside a break paragraph: exactly ONE half keeps the break.
      // Word's page break is a character right before the paragraph content,
      // so the break stays with the half that starts with the original
      // content — at offset 0 that is the second half (the new empty line
      // above must not steal it), everywhere else the first (splitting must
      // not clone the break onto the new paragraph and turn Enter into
      // another page jump — alpha ledger r157). TipTap's keepOnSplit only
      // filters end-of-paragraph splits, so mid-splits are fixed up here.
      Enter: () => {
        const { $from, empty } = this.editor.state.selection
        if (!empty) return false
        const parent = $from.parent
        if (!parent.isTextblock || parent.attrs.pageBreakBefore !== true) return false
        // empty block / list item: lift-out and list splits have their own semantics
        if (parent.content.size === 0 || parent.type.name === 'docListItem') return false
        const atStart = $from.parentOffset === 0
        return this.editor
          .chain()
          .splitBlock()
          .command(({ state, tr, dispatch }) => {
            const caret = state.selection.$from
            const secondPos = caret.before(caret.depth)
            const firstHalf = state.doc.resolve(secondPos).nodeBefore
            if (!firstHalf || !('pageBreakBefore' in firstHalf.attrs)) return false
            if (dispatch) {
              if (atStart) {
                tr.setNodeMarkup(secondPos - firstHalf.nodeSize, undefined, {
                  ...firstHalf.attrs,
                  pageBreakBefore: false,
                })
                tr.setNodeMarkup(secondPos, undefined, {
                  ...caret.parent.attrs,
                  pageBreakBefore: true,
                })
              } else if (caret.parent.attrs.pageBreakBefore === true) {
                tr.setNodeMarkup(secondPos, undefined, {
                  ...caret.parent.attrs,
                  pageBreakBefore: false,
                })
              }
              dispatch(tr.scrollIntoView())
            }
            return true
          })
          .run()
      },
      'Mod-Enter': () => insertPageBreak(this.editor),
      'Mod-Shift-Enter': () =>
        this.editor.commands.insertContent({ type: 'hardBreak', attrs: { colBreak: true } }),
      // U+00A0 and U+2011: the characters Word inserts for these two chords
      'Mod-Shift-Space': () => this.editor.commands.insertContent('\u00a0'),
      'Mod-Shift--': () => this.editor.commands.insertContent('\u2011'),
      'Alt-Shift-ArrowUp': () => moveBlocks(this.editor, -1),
      'Alt-Shift-ArrowDown': () => moveBlocks(this.editor, 1),
    }
  },
})

export const DocParagraph = Node.create({
  name: 'docParagraph',
  group: 'block',
  content: 'inline*',
  addAttributes() {
    return { ...anchorAttrs }
  },
  parseHTML() {
    // getAttrs null = match with defaults (foreign HTML); data-para restores formatting
    return [{ tag: 'p', getAttrs: (el) => clipboardParaAttrs(el as HTMLElement) }]
  },
  renderHTML({ node }) {
    return ['p', blockAttrs(node), 0]
  },
})

export const DocHeading = Node.create({
  name: 'docHeading',
  group: 'block',
  content: 'inline*',
  addAttributes() {
    return {
      ...anchorAttrs,
      level: { default: 1 },
      /** level from a direct w:outlineLvl on a non-heading style (body formatting) */
      outlineOnly: { default: null as boolean | null },
    }
  },
  parseHTML() {
    return [1, 2, 3, 4, 5, 6].map((level) => ({
      tag: `h${level}`,
      getAttrs: (el) => ({ level, ...clipboardParaAttrs(el as HTMLElement) }),
    }))
  },
  renderHTML({ node }) {
    const level = Math.min(Math.max(Number(node.attrs.level) || 1, 1), 6)
    const attrs = blockAttrs(node)
    // heading by direct w:outlineLvl alone: the built-in h1-h6 font rules skip this class
    if (node.attrs.outlineOnly) attrs.class = `${attrs.class ?? ''} doc-outline-only`.trim()
    return [`h${level}`, attrs, 0]
  },
})

export const DocListItem = Node.create({
  name: 'docListItem',
  group: 'block',
  content: 'inline*',
  addAttributes() {
    return {
      ...anchorAttrs,
      kind: { default: 'bullet' as 'bullet' | 'ordered' },
      numId: { default: null as string | null },
      ilvl: { default: 0 },
    }
  },
  parseHTML() {
    const ilvlOf = (el: HTMLElement): number => {
      let depth = -1
      for (let node = el.parentElement; node; node = node.parentElement) {
        if (node.tagName === 'UL' || node.tagName === 'OL') depth++
      }
      return Math.max(0, Math.min(depth, 8))
    }
    return [
      {
        tag: 'li',
        getAttrs: (el) => ({
          kind: (el as HTMLElement).closest('ol') ? 'ordered' : 'bullet',
          ilvl: ilvlOf(el as HTMLElement),
          ...clipboardParaAttrs(el as HTMLElement),
        }),
      },
      // our own clipboard HTML: renderHTML emits <div class="doc-li …">, which
      // no rule matched before r117 — pasting a GenOffice list item degraded it
      // to plain text. kind/ilvl ride in data-para; classes are the fallback.
      {
        tag: 'div.doc-li',
        getAttrs: (el) => {
          const div = el as HTMLElement
          const attrs = clipboardParaAttrs(div) ?? {}
          if (attrs.kind === undefined) {
            attrs.kind = div.classList.contains('doc-li-ordered') ? 'ordered' : 'bullet'
          }
          if (attrs.ilvl === undefined) {
            const m = /(?:^|\s)ilvl-(\d)(?:\s|$)/.exec(div.className)
            attrs.ilvl = m ? Number(m[1]) : 0
          }
          return attrs
        },
      },
    ]
  },
  renderHTML({ node }) {
    const base = blockAttrs(node, { includeIndent: false, listGeometry: true })
    const cls = [
      'doc-li',
      `doc-li-${node.attrs.kind}`,
      `ilvl-${Math.min(Number(node.attrs.ilvl) || 0, 4)}`,
      base['class'] ?? '',
    ]
      .filter(Boolean)
      .join(' ')
    return ['div', { ...base, class: cls }, 0]
  },
  addCommands() {
    return {
      /**
       * Word's Enter behavior inside a list: a non-empty item splits
       * into a sibling item (same kind/numId/level, numbering follows), an empty
       * one leaves the list. ProseMirror's default splitBlock produces the
       * schema's default block — a paragraph — which broke continuous entry.
       */
      continueDocList:
        () =>
        ({ state, chain }: { state: EditorState; chain: () => ChainedCommands }) => {
          const { $from, empty } = state.selection
          if (!empty) return false
          const node = $from.parent
          if (node.type.name !== 'docListItem') return false
          if (node.content.size === 0) {
            return chain()
              .setNode('docParagraph', { ...node.attrs, docxIndex: null })
              .run()
          }
          return chain()
            .splitBlock()
            .command(({ tr, dispatch }) => {
              const pos = tr.selection.$from.before(tr.selection.$from.depth)
              const created = tr.doc.nodeAt(pos)
              if (!created) return false
              // The split half is a fresh paragraph: turn it back into a sibling
              // list item, without inheriting the original's docx anchor
              if (dispatch) {
                tr.setNodeMarkup(pos, state.schema.nodes.docListItem, {
                  ...created.attrs,
                  docxIndex: null,
                  kind: node.attrs.kind,
                  numId: node.attrs.numId,
                  ilvl: node.attrs.ilvl,
                })
              }
              return true
            })
            .run()
        },
    } as Partial<RawCommands>
  },
  addKeyboardShortcuts() {
    const changeLevel = (delta: number) => () => {
      if (!this.editor.isActive('docListItem')) return false
      const ilvl = Number(this.editor.getAttributes('docListItem').ilvl) || 0
      const next = Math.min(Math.max(ilvl + delta, 0), 8)
      if (next === ilvl) return true
      return this.editor.commands.updateAttributes('docListItem', { ilvl: next })
    }
    return {
      Tab: changeLevel(1),
      'Shift-Tab': changeLevel(-1),
      Enter: () => (this.editor.commands as unknown as DocListCommands).continueDocList(),
    }
  },
})

interface DocListCommands {
  continueDocList: () => boolean
}

export interface ListNumberingStorage {
  /** numId -> definition, from the open document's numbering.xml */
  defs: Map<string, NumberingDef>
  /** style/docDefaults display info: marker measurement reads the same font
   *  chain the ::before inherits (data-style rule -> .doc-page baseline) */
  styles?: Map<string, StyleInfo>
  docDefaults?: DocDefaults
}

declare module '@tiptap/core' {
  interface Storage {
    listNumbering: ListNumberingStorage
  }
}

/**
 * Real multilevel numbering: compute each list item's marker in document order per the
 * numbering.xml definitions (1. / a. / 1.1 / Chinese numerals, …), attach a data-marker
 * attribute via node decoration, and display with CSS.
 * defs are written into storage by App when a document is opened/re-parsed; items without a definition fall back to CSS counters.
 */
// ── Live line-height factor + strut font-size ────────────────────────────────
// blockAttrs bakes --doc-line-factor and the strut font-size into toDOM output,
// but ProseMirror reuses a block's DOM while typing (sameMarkup ignores content),
// so a block edited after creation keeps its creation-time values until
// save/reopen. These node decorations recompute both from the live content on
// every doc change; unchanged nodes are structurally shared, so the WeakMap
// makes a pass cheap. Dropping font-size from a decoration removes the baked
// value too (ProseMirror's patchAttributes calls style.removeProperty for every
// property named in the previous decoration style).

const LINE_FACTOR_BLOCKS = new Set(['docParagraph', 'docHeading', 'docListItem'])
const lineFactorCache = new WeakMap<
  PmNode,
  { style: string; cls?: string; runs?: RunFactorRange[] }
>()

const autospaceRangesCache = new WeakMap<PmNode, Array<{ from: number; to: number }>>()

const simsunGapCache = new WeakMap<PmNode, Array<{ from: number; to: number }>>()

/** ranges (relative to the block's content start) of ・/〜 in SimSun-substituted runs */
function simsunGapRanges(node: PmNode): Array<{ from: number; to: number }> {
  let ranges = simsunGapCache.get(node)
  if (ranges === undefined) {
    const found: Array<{ from: number; to: number }> = []
    let offset = 0
    node.forEach((child) => {
      if (child.isText && child.text && SIMSUN_GAP_CHAR_RE.test(child.text)) {
        const mark = child.marks.find((m) => m.type.name === 'docTextStyle')
        const family =
          mark?.attrs.eaSlotEmpty === true
            ? null
            : ((mark?.attrs.font ?? mark?.attrs.fontAscii) as string | null | undefined)
        if (family && simsunGapLineFactor(family) !== null) {
          for (let i = 0; i < child.text.length; i++) {
            if (SIMSUN_GAP_CHAR_RE.test(child.text[i])) {
              found.push({ from: offset + i, to: offset + i + 1 })
            }
          }
        }
      }
      offset += child.nodeSize
    })
    ranges = found
    simsunGapCache.set(node, ranges)
  }
  return ranges
}

/**
 * ranges (relative to the block's content start) of the character after each
 * CJK-Latin pad boundary; that character carries the pad margin (.doc-autospace-pad)
 */
function autospaceRanges(node: PmNode): Array<{ from: number; to: number }> {
  let ranges = autospaceRangesCache.get(node)
  if (ranges === undefined) {
    const found: Array<{ from: number; to: number }> = []
    let offset = 0
    let prevText = '' // reset by non-text inlines: pads need direct adjacency
    node.forEach((child) => {
      if (child.isText && child.text) {
        const text = child.text
        const cuts = autospaceBoundaries(text)
        if (autospacePadBetween(prevText, text)) cuts.unshift(0)
        for (const i of cuts) {
          found.push({ from: offset + i, to: offset + i + codePointLengthAt(text, i) })
        }
        prevText = text
      } else {
        prevText = ''
      }
      offset += child.nodeSize
    })
    ranges = found
    autospaceRangesCache.set(node, ranges)
  }
  return ranges
}

/** blocks hosting a front/behind picture are its positioning origin (styles.css .doc-anchor-origin) */
const anchorOriginCache = new WeakMap<PmNode, boolean>()
function hostsAnchoredPicture(node: PmNode): boolean {
  let hosts = anchorOriginCache.get(node)
  if (hosts === undefined) {
    hosts = false
    node.descendants((child) => {
      if (
        child.type.name === 'docInlineImage' &&
        (child.attrs.wrap === 'front' || child.attrs.wrap === 'behind')
      ) {
        hosts = true
      }
      return !hosts
    })
    anchorOriginCache.set(node, hosts)
  }
  return hosts
}

function lineFactorDecos(doc: PmNode): DecorationSet {
  const decos: Decoration[] = []
  doc.descendants((node, pos) => {
    if (!LINE_FACTOR_BLOCKS.has(node.type.name)) return true
    // a class decoration instead of a stylesheet :has(): Blink's :has()
    // invalidation crashed the renderer (OOM) on long picture-heavy documents
    if (hostsAnchoredPicture(node)) {
      decos.push(Decoration.node(pos, pos + node.nodeSize, { class: 'doc-anchor-origin' }))
    }
    if (node.textContent) {
      let cached = lineFactorCache.get(node)
      if (cached === undefined) {
        const perLine = perLineFactors(node)
        let style = `--doc-line-factor:${perLine ? perLine.strut : paraLineFactor(node)}`
        let cls: string | undefined
        const fam = paraDeclaredFontFamily(node)
        if (fam) style += `;font-family:${fam}`
        else if (paraMixedDeclaredCjk(node)) cls = 'doc-grid-strut'
        const strut = explicitStrutHalfPoints(node)
        if (strut) style += `;${strutFontCss(strut).join(';')}`
        cached = { style, ...(cls ? { cls } : {}), ...(perLine ? { runs: perLine.runs } : {}) }
        lineFactorCache.set(node, cached)
      }
      decos.push(
        Decoration.node(pos, pos + node.nodeSize, {
          style: cached.style,
          ...(cached.cls ? { class: cached.cls } : {}),
        }),
      )
      if (cached.runs) {
        for (const r of cached.runs) {
          decos.push(
            Decoration.inline(pos + 1 + r.from, pos + 1 + r.to, {
              class: 'doc-run-lf',
              style: r.style,
            }),
          )
        }
      }
      if (node.attrs.autoSpace !== false) {
        for (const r of autospaceRanges(node)) {
          decos.push(
            Decoration.inline(pos + 1 + r.from, pos + 1 + r.to, { class: 'doc-autospace-pad' }),
          )
        }
      }
      // ・/〜 in SimSun-substituted runs: Word lifts the whole line to 1.7143 ×
      // size (probe 2026-08-13); a taller inline strut reproduces the row lift.
      // exact lineRule pins the line, so no lift there.
      if (node.attrs.lineRule !== 'exact') {
        const ranges = simsunGapRanges(node)
        if (ranges.length > 0) {
          const m =
            Number(node.attrs.lineSpacing) ||
            (node.attrs.lineRule === 'auto' && node.attrs.lineRawTwips
              ? Number(node.attrs.lineRawTwips) / 240
              : 1)
          const gapStyle = `line-height:${cssSimsunGapLineExpr(m)}`
          for (const r of ranges) {
            decos.push(Decoration.inline(pos + 1 + r.from, pos + 1 + r.to, { style: gapStyle }))
          }
        }
      }
      // symbols Chromium justifies like ideographs stay unstretched in non-CJK
      // paragraphs (Word stretches only their spaces); CJK paragraphs keep
      // Chromium's inter-ideograph distribution around them
      if (
        (node.attrs.align === 'justify' || node.attrs.align === 'distribute') &&
        !textHasCjk(node.textContent)
      ) {
        for (const r of justifySymbolOffsets(node)) {
          decos.push(
            Decoration.inline(pos + 1 + r.from, pos + 1 + r.to, { class: 'doc-justify-symbol' }),
          )
        }
      }
    }
    return false
  })
  return DecorationSet.create(doc, decos)
}

/** ranges (relative to the block's content start) of Chromium's justification symbols */
function justifySymbolOffsets(node: PmNode): Array<{ from: number; to: number }> {
  const found: Array<{ from: number; to: number }> = []
  let offset = 0
  node.forEach((child) => {
    if (child.isText && child.text) {
      for (const r of justifySymbolRanges(child.text)) {
        found.push({ from: offset + r.from, to: offset + r.to })
      }
    }
    offset += child.nodeSize
  })
  return found
}

export const LineFactorExtension = Extension.create({
  name: 'lineFactorLive',
  addProseMirrorPlugins() {
    const key = new PluginKey<DecorationSet>('lineFactorLive')
    const editor = this.editor
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: (_config, state) => lineFactorDecos(state.doc),
          apply: (tr, old) => {
            if (tr.getMeta(key)) return lineFactorDecos(tr.doc)
            if (!tr.docChanged) return old
            // inserting pad widgets next to an active IME composition aborts it;
            // keep the old set mapped and refresh on compositionend
            if (editor?.view?.composing) return old.map(tr.mapping, tr.doc)
            return lineFactorDecos(tr.doc)
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
          handleDOMEvents: {
            compositionend: (view) => {
              window.setTimeout(() => {
                if (!view.isDestroyed) view.dispatch(view.state.tr.setMeta(key, true))
              })
              return false
            },
          },
        },
      }),
    ]
  },
})

function firstRunSizeHalfPoints(node: PmNode): number | null {
  let sz: number | null = null
  node.descendants((child) => {
    if (sz !== null) return false
    if (child.isText) {
      const mark = child.marks.find((m) => m.type.name === 'docTextStyle')
      sz = (mark?.attrs.sizeHalfPoints as number | null) ?? null
      return false
    }
    return true
  })
  return sz
}

/** the font the ::before marker actually inherits: the item's [data-style] rule,
 *  else the .doc-page baseline (Normal / docDefaults) — same chain as doc-style-css */
function markerParagraphFont(
  styleId: unknown,
  storage: ListNumberingStorage,
): { family: string; sizeHalf: number; bold: boolean } {
  const style = typeof styleId === 'string' ? storage.styles?.get(styleId)?.display : undefined
  let normal: StyleDisplay | undefined
  for (const info of storage.styles?.values() ?? []) {
    if (info.isDefault && info.type === 'paragraph' && info.display) {
      normal = info.display
      break
    }
  }
  const dd = storage.docDefaults
  return {
    // ascii slot first: dual-slot font-family rules put the Latin face first,
    // and markers are Latin/digit text
    family:
      style?.fontAscii ??
      style?.font ??
      normal?.fontAscii ??
      dd?.asciiFont ??
      normal?.font ??
      dd?.eastAsiaFont ??
      'Calibri',
    sizeHalf: style?.sizeHalfPoints ?? normal?.sizeHalfPoints ?? dd?.sizeHalfPoints ?? 20,
    bold: style?.bold ?? normal?.bold ?? dd?.bold ?? false,
  }
}

let markerMeasureCtx: CanvasRenderingContext2D | null | undefined
/** marker advance in twips via canvas metrics; null in DOM-less test envs */
function measureMarkerTwips(
  text: string,
  family: string,
  sizePt: number,
  bold: boolean,
): number | null {
  if (markerMeasureCtx === undefined) {
    try {
      markerMeasureCtx = document.createElement('canvas').getContext('2d')
    } catch {
      markerMeasureCtx = null
    }
  }
  if (!markerMeasureCtx) return null
  const weight = bold ? '600 ' : ''
  markerMeasureCtx.font = `${weight}${(sizePt * 4) / 3}px "${family.replaceAll('"', '')}", Calibri, sans-serif`
  const w = markerMeasureCtx.measureText(text).width
  return Number.isFinite(w) && w > 0 ? Math.round(w * 15) : null
}

/** substitute glyph for a symbol-font bullet: pin a Latin font (CJK fallback draws ・) and
 *  compensate its smaller bullet; --li-marker-lh:0 keeps the scaled em box from stretching the line */
function substituteMarkerStyles(text: string): string[] {
  const styles = [`--li-marker-font:Arial,'Helvetica Neue',sans-serif`]
  const scale = bulletMarkerScale(text)
  if (scale !== 1) styles.push(`--li-marker-scale:${scale}`, '--li-marker-lh:0')
  return styles
}

/** Word clips cell content at the cell's text area: a marker whose box ends before it
 *  (hanging deeper than the start indent) is never seen; ~0.5em stands in for the glyph advance.
 *  Centered / end-aligned lines shift the marker back inside, so only start-aligned items apply. */
function markerClippedByCell(
  doc: PmNode,
  pos: number,
  attrs: Record<string, unknown>,
  level: NumberingLevel,
  szHalfPoints: number,
): boolean {
  const align = attrs.align as string | null | undefined
  const startSide = attrs.bidi || attrs.bidiInferred ? 'right' : 'left'
  if (align && align !== startSide && align !== 'justify' && align !== 'distribute') return false
  const $pos = doc.resolve(pos)
  let inCell = false
  for (let d = $pos.depth; d > 0 && !inCell; d--) {
    const name = $pos.node(d).type.name
    inCell = name === 'docTableCell' || name === 'docTableHeader'
  }
  if (!inCell) return false
  const left =
    attrs.indentLeft != null
      ? Number(attrs.indentLeft)
      : (level.indentLeft ?? 792 + 432 * Number(attrs.ilvl || 0))
  const firstLine = specialIndentTw(attrs, level) ?? -360
  return left + firstLine + szHalfPoints * 5 <= 0
}

/**
 * Word merges w:ind per attribute: the paragraph's own firstLine/hanging (an
 * explicit firstLine="0" included) replaces the level's special indent.
 * null = the level carries no indent at all and the CSS default hang applies.
 */
function specialIndentTw(attrs: Record<string, unknown>, level: NumberingLevel): number | null {
  if (attrs.indentFirstLine != null) return Number(attrs.indentFirstLine)
  if (level.hanging) return -level.hanging
  return level.firstLine ?? (level.indentLeft ? 0 : null)
}

function cssString(text: string): string {
  return `"${text.replace(/["\\]/g, '\\$&')}"`
}

/** the stray line's list geometry in docListItem attr shape (its w:ind is the anchor paragraph's) */
function strayGeometryAttrs(node: PmNode): Record<string, unknown> {
  const ind = node.attrs.strayIndent as StrayIndent | null
  return {
    indentLeft: ind?.leftTwips ?? null,
    indentFirstLine: ind?.firstLineTwips ?? null,
    styleId: node.attrs.strayStyleId,
    ilvl: (node.attrs.strayList as ListItemRef).ilvl,
  }
}

/** the anchor paragraph's own empty line, laid out like any paragraph; the
 *  picture block ahead of it already carries the paragraph's page break */
export function anchorLineSpec(line: Record<string, unknown>): DomSpec {
  const attrs = blockAttrs({
    attrs: { ...line, docxIndex: null, pageBreakBefore: false },
    textContent: '',
  })
  delete attrs['data-para']
  attrs.class = `${attrs.class ? `${attrs.class} ` : ''}doc-anchor-line`
  return ['p', attrs, ['br']]
}

/**
 * Word keeps the anchor paragraph's own (empty) line beside a side-wrapped
 * picture. The image block is the CSS float itself, so the line is a
 * display-only block widget right after it (never serialized).
 */
export const AnchorLineExtension = Extension.create({
  name: 'anchorLine',
  addProseMirrorPlugins() {
    const build = (doc: PmNode): DecorationSet => {
      const decos: Decoration[] = []
      doc.descendants((node, pos) => {
        if (node.type.name !== 'docProtected') return true
        const line = node.attrs.anchorLine as Record<string, unknown> | null
        if (
          line &&
          !node.attrs.invisibleMarker &&
          /-(?:left|right)$/.test(String(node.attrs.imageWrap ?? ''))
        ) {
          decos.push(
            Decoration.widget(
              pos + node.nodeSize,
              () => DOMSerializer.renderSpec(document, anchorLineSpec(line) as never).dom,
              { side: -1, key: `anchor-line-${String(node.attrs.docxIndex)}` },
            ),
          )
        }
        return false
      })
      return DecorationSet.create(doc, decos)
    }
    return [
      new Plugin({
        state: {
          init: (_, state) => build(state.doc),
          apply: (tr, old) => (tr.docChanged ? build(tr.doc) : old),
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
        },
      }),
    ]
  },
})

export const ListNumberingExtension = Extension.create<object, ListNumberingStorage>({
  name: 'listNumbering',
  addStorage() {
    return { defs: new Map<string, NumberingDef>() }
  },
  addProseMirrorPlugins() {
    const storage = this.storage
    const compute = (doc: PmNode): DecorationSet | null => {
      if (storage.defs.size === 0) return null
      const refs: ListItemRef[] = []
      const nodes: Array<{ pos: number; node: PmNode }> = []
      doc.descendants((node, pos) => {
        if (node.type.name === 'docListItem') {
          refs.push({
            numId: (node.attrs.numId as string | null) ?? null,
            ilvl: Number(node.attrs.ilvl) || 0,
          })
          nodes.push({ pos, node })
          return false
        }
        // numbered anchor paragraph of a textbox block: Word keeps it in the
        // sequence, and its stray line shows the marker
        const strayList =
          node.type.name === 'docProtected' ? (node.attrs.strayList as ListItemRef | null) : null
        if (strayList) {
          refs.push({ numId: strayList.numId, ilvl: strayList.ilvl })
          nodes.push({ pos, node })
          return false
        }
        return true
      })
      if (refs.length === 0) return null
      const markers = computeListMarkerInfos(refs, storage.defs)
      const decos: Decoration[] = []
      markers.forEach((marker, i) => {
        if (marker === null) return
        const styles: string[] = []
        let text = marker.text
        if (marker.symbolFont && marker.symbolChar) {
          if (symbolFontCovers(marker.symbolFont, marker.symbolChar)) {
            text = marker.symbolChar
            styles.push(`--li-marker-font:"${marker.symbolFont}"`)
          } else styles.push(...substituteMarkerStyles(text))
        }
        // a protected node's decoration lands on the wrapper, out of attr()'s reach for
        // the inner stray line: the marker travels as an inherited custom property
        const stray = nodes[i].node.type.name === 'docProtected'
        const attrs: Record<string, string> = stray
          ? { 'data-stray-marker': '' }
          : { 'data-marker': text }
        if (stray) styles.push(`--li-marker:${cssString(text)}`)
        // geometry fallback: when the paragraph has no w:ind of its own, use the numbering.xml level's indent;
        // marker font size comes from the level's rPr, else follows the item's first text run (Word rule)
        const def = refs[i].numId !== null ? storage.defs.get(refs[i].numId as string) : undefined
        const level = def?.levels[Math.max(0, refs[i].ilvl)]
        if (level) {
          const nodeAttrs = stray ? strayGeometryAttrs(nodes[i].node) : nodes[i].node.attrs
          // an explicit w:ind left (0 included) beats the numbering level's indent
          const leftTw =
            nodeAttrs.indentLeft != null ? Number(nodeAttrs.indentLeft) : level.indentLeft || 0
          if (nodeAttrs.indentLeft == null && level.indentLeft) {
            styles.push(`--li-left:${level.indentLeft / 20}pt`)
          }
          const ownFirst =
            nodeAttrs.indentFirstLine == null ? null : Number(nodeAttrs.indentFirstLine)
          const firstTw = specialIndentTw(nodeAttrs, level)
          if (ownFirst == null && firstTw != null && firstTw < 0) {
            styles.push(`--li-hang:${-firstTw / 20}pt`)
          }
          // no hanging area: the marker sits at the first-line position instead of
          // hanging the CSS default before it; its tab then runs to the next default stop
          if (firstTw != null && firstTw >= 0) {
            styles.push('--li-hang:0pt')
            if (ownFirst == null && firstTw > 0) styles.push(`text-indent:${firstTw / 20}pt`)
          }
          const szHalf =
            level.szHalfPoints ??
            (stray
              ? (nodes[i].node.attrs.strayRuns as Run[] | null)?.[0]?.sizeHalfPoints
              : firstRunSizeHalfPoints(nodes[i].node)) ??
            undefined
          if (szHalf) styles.push(`--li-marker-size:${szHalf / 2}pt`)
          if (level.suff === 'space' || level.suff === 'nothing') attrs['data-suff'] = level.suff
          if (markerClippedByCell(doc, nodes[i].pos, nodeAttrs, level, szHalf ?? 20))
            attrs['data-marker-clip'] = ''

          // Word's default tab after the marker: a marker that escapes the hanging
          // area (or one with no hanging area at all) pushes first-line text to
          // the next default tab stop, not right up against the marker
          if (
            (level.suff === undefined || level.suff === 'tab') &&
            text &&
            firstTw != null &&
            (firstTw >= 0 || leftTw > 0)
          ) {
            // no level/run size -> the marker renders at the li's 1em; family always
            // inherits from the li (level.font only applies to symbol bullets)
            const para = markerParagraphFont(nodeAttrs.styleId, storage)
            const widthTw = measureMarkerTwips(
              text,
              para.family,
              (szHalf ?? para.sizeHalf) / 2,
              para.bold,
            )
            const adv =
              widthTw !== null ? markerTabAdvance(leftTw + firstTw, widthTw, leftTw) : null
            if (adv !== null) styles.push(`--li-tab:${adv / 20}pt`)
          }
        }
        if (styles.length > 0) attrs.style = styles.join(';')
        decos.push(Decoration.node(nodes[i].pos, nodes[i].pos + nodes[i].node.nodeSize, attrs))
      })
      return DecorationSet.create(doc, decos)
    }

    interface CachedMarkers {
      defs: Map<string, NumberingDef>
      decos: DecorationSet | null
    }
    const key = new PluginKey<CachedMarkers>('listNumbering')
    return [
      new Plugin<CachedMarkers>({
        key,
        state: {
          init: (_config, state) => ({ defs: storage.defs, decos: compute(state.doc) }),
          apply(tr, old) {
            // defs is replaced (never mutated) on open/reparse and marker overlay
            if (tr.docChanged || old.defs !== storage.defs)
              return { defs: storage.defs, decos: compute(tr.doc) }
            return old.decos ? { defs: old.defs, decos: old.decos.map(tr.mapping, tr.doc) } : old
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)?.decos ?? null
          },
        },
      }),
    ]
  },
})

const tableCellAttrs = {
  vAlign: { default: null as string | null },
  borders: { default: null as Record<string, unknown> | null },
  rawTcPr: { default: null as string | null },
  /** tcPr w:cellIns/w:cellDel cell revision ({kind, author, ...} | null) */
  cellRevision: { default: null as Record<string, string> | null },
  cellMar: { default: null as Record<string, number> | null },
  textDirection: { default: null as string | null },
  /** inner clip-box height (twips) when the row is hRule="exact" (computed in convert.ts) */
  clipHeightTwips: { default: null as number | null },
  /** display placeholder for the row's w:gridBefore/w:gridAfter columns (borderless, not saved as w:tc) */
  gridGap: { default: false },
  colspan: { default: 1 },
  rowspan: { default: 1 },
  colwidth: { default: null as number[] | null },
  fill: { default: null as string | null },
  color: { default: null as string | null },
  bold: { default: false },
  align: { default: null as string | null },
}

export function borderWidthPx(b: { style: string; szEighths?: number } | undefined | null): number {
  if (!b || b.style === 'none' || b.style === 'nil') return 0
  return Math.max(1, Math.round(((b.szEighths ?? 4) / 8 / 72) * 96))
}

/** true border width (sz/8 pt) minus the whole px we draw. Word lays rows out
 *  with the true width (probe 2026-09-04: 0.5pt adds 0.667px per row, 1pt 1.333px);
 *  the td::before spacer in styles.css gives this back to the row height */
export function borderSnapDeltaPx(
  b: { style: string; szEighths?: number } | undefined | null,
): number {
  if (!b || b.style === 'none' || b.style === 'nil') return 0
  return ((b.szEighths ?? 4) / 8 / 72) * 96 - borderWidthPx(b)
}

export const bdDeltaCss = (
  side: 'top' | 'bottom',
  b: { style: string; szEighths?: number } | undefined | null,
): string => `--cell-bd-${side === 'top' ? 't' : 'b'}:${borderSnapDeltaPx(b).toFixed(3)}px`

type BorderLine = { style: string; szEighths?: number } | undefined | null
const trueBorderPx = (b: BorderLine): number =>
  !b || b.style === 'none' || b.style === 'nil' ? 0 : ((b.szEighths ?? 4) / 8 / 72) * 96

/** Word advances a declared-height row by trHeight PLUS the true gridline width
 *  (CSS collapsed borders fit inside the tr height): the table emits --doc-row-eat
 *  from insideH, rows with explicit tcBorders resolve their own gridline (sas 035/048) */
export function tableRowEatCss(borders: TableBordersAttr | null, cellSpacing: boolean): string[] {
  // spaced cells never collapse a gridline into the row height
  if (cellSpacing) return ['--doc-row-grid:0']
  const insideH = borders?.insideH
  if (!insideH || insideH.style === 'none' || insideH.style === 'nil') return []
  return [`--doc-row-eat:${trueBorderPx(insideH).toFixed(2)}px`]
}

/** cells: gridBefore/gridAfter placeholders must be filtered out by the caller */
export function rowHeightCss(
  heightTwips: number,
  cellBorders: Array<{ top?: BorderLine; bottom?: BorderLine } | null | undefined>,
): string {
  let explicit = 0
  let inherits = false
  for (const b of cellBorders) {
    if (!b?.top || !b?.bottom) inherits = true
    explicit = Math.max(explicit, trueBorderPx(b?.top), trueBorderPx(b?.bottom))
  }
  const fixed = ((heightTwips / 1440) * 96).toFixed(1)
  if (explicit > 0) {
    const line = inherits
      ? `max(${explicit.toFixed(2)}px, var(--doc-row-eat,0px))`
      : `${explicit.toFixed(2)}px`
    return `height:calc(${fixed}px + ${line} * var(--doc-row-grid,1))`
  }
  return inherits ? `height:calc(${fixed}px + var(--doc-row-eat,0px))` : `height:${fixed}px`
}

function tableCellHtml(node: PmNode): Record<string, string> {
  const attrs: Record<string, string> = {}
  // gap placeholders never save back as w:tc, so typed text would vanish — keep
  // them inert (data-grid-gap: not in-flow evidence for the .cell-vert switch)
  if (node.attrs.gridGap) {
    attrs.contenteditable = 'false'
    attrs['data-grid-gap'] = '1'
  }
  if (node.attrs.colspan > 1) attrs.colspan = String(node.attrs.colspan)
  if (node.attrs.rowspan > 1) attrs.rowspan = String(node.attrs.rowspan)
  if (node.attrs.colwidth) attrs['data-colwidth'] = (node.attrs.colwidth as number[]).join(',')
  const cellBorders = node.attrs.borders as Record<
    string,
    { style: string; szEighths?: number; color?: string }
  > | null
  const borderCss = (side: 'top' | 'right' | 'bottom' | 'left'): string => {
    const v = borderLineCss(cellBorders?.[side])
    if (!v) return ''
    const css = `border-${side}:${v};${dkBorder(DK_SIDE[side], v)}`
    // the text inset rule (styles.css) needs the overriding side's width
    return side === 'left' || side === 'right'
      ? `${css};--cell-bw-${DK_SIDE[side]}:${borderWidthPx(cellBorders?.[side])}px`
      : `${css};${bdDeltaCss(side, cellBorders?.[side])}`
  }
  const mar = node.attrs.cellMar as Record<string, number> | null
  const styles = [
    // gridBefore/gridAfter placeholder: bare grid space (inline border beats the --doc-b-* cell rules)
    node.attrs.gridGap ? 'border:none;background:none' : '',
    // vertical-text cells: content lives in an absolute .cell-vert wrapper (see
    // tableCellSpec) anchored to this cell
    node.attrs.textDirection ? 'position:relative' : '',
    // font-weight before background: jsdom's CSSOM drops the background getter
    // when font-weight follows it (order is irrelevant to real browsers)
    node.attrs.bold ? 'font-weight:600' : '',
    // authored colors stay the declaration; the --dk-* twins feed the dark page
    // (dark-page.ts). background-color, not the shorthand: jsdom's CSSOM drops a
    // `background` shorthand that shares the attribute with a custom property
    node.attrs.color ? `color:#${node.attrs.color};${dkColor(String(node.attrs.color))}` : '',
    node.attrs.fill
      ? `background-color:#${node.attrs.fill};${dkBackground(`#${node.attrs.fill}`)}`
      : '',
    node.attrs.align ? `text-align:${node.attrs.align}` : '',
    node.attrs.vAlign && node.attrs.vAlign !== 'top'
      ? `vertical-align:${node.attrs.vAlign === 'center' ? 'middle' : 'bottom'}`
      : '',
    borderCss('top'),
    borderCss('left'),
    borderCss('bottom'),
    borderCss('right'),
    // tcMar only overrides declared sides; the rest inherit the table-level --doc-cell-pad-*
    ...(['top', 'left', 'bottom', 'right'] as const).map((side) =>
      mar?.[side] !== undefined
        ? `--doc-cell-pad-${DK_SIDE[side]}:${(mar[side] / 15).toFixed(1)}px`
        : '',
    ),
    Array.isArray(node.attrs.colwidth)
      ? `width:${(node.attrs.colwidth as number[]).reduce((sum, width) => sum + width, 0)}px`
      : '',
  ].filter(Boolean)
  if (styles.length > 0) attrs.style = styles.join(';')
  {
    const ink = fillInk(node.attrs.fill)
    if (ink) attrs['data-ink'] = ink
  }
  const cellRev = node.attrs.cellRevision as { kind?: string; author?: string } | null
  if (cellRev?.kind) {
    attrs.class = `cell-rev-${cellRev.kind}`
    if (cellRev.author) attrs.title = cellRev.author
  }
  return attrs
}

/** exact-height rows clip via a fixed-height inner box: td height is min-height
 *  semantics in CSS table layout, so the td itself can never produce overflow */
export function cellClipStyle(vAlign: string | null, clipTwips: number): string {
  return [
    `height:${(clipTwips / 15).toFixed(1)}px`,
    // replicate the td's vertical-align inside the fixed box (the box fills the cell);
    // 'safe' falls back to start when content overflows — Word keeps the top and clips
    // the bottom edge regardless of vAlign
    ...(vAlign === 'center'
      ? ['display:grid', 'align-content:safe center']
      : vAlign === 'bottom'
        ? ['display:grid', 'align-content:safe end']
        : []),
  ].join(';')
}

/** tbRl = vertical right-to-left (vertical-rl), btLr = rotated 90° counterclockwise (sideways-lr) */
export function cellWritingMode(textDirection: string | null): string | null {
  if (textDirection === 'tbRl') return 'writing-mode:vertical-rl'
  if (textDirection === 'btLr') return 'writing-mode:sideways-lr'
  return null
}

/** td vertical-align replicated inside a cell wrapper box (block-axis alignment,
 *  matching the td's own logical behavior when the writing mode sat on it) */
export function cellVAlignGridCss(vAlign: string | null): string | null {
  if (vAlign === 'center') return 'display:grid;align-content:safe center'
  if (vAlign === 'bottom') return 'display:grid;align-content:safe end'
  return null
}

function tableCellSpec(tag: 'td' | 'th', node: PmNode): DOMOutputSpec {
  const attrs = tableCellHtml(node)
  const clip = node.attrs.clipHeightTwips as number | null
  const wm = cellWritingMode(node.attrs.textDirection as string | null)
  // 0 is a real clip height (padding/borders consume the whole exact row)
  if (clip == null) {
    if (!wm) return [tag, attrs, 0]
    // Word sizes the row from horizontal cells / trHeight and wraps rotated text
    // into the resulting cell height; the out-of-flow wrapper reproduces that
    // (in-flow, an unwrappable vertical line stretches the whole row group).
    // Rows with no other height source drop the wrapper back into flow via a
    // structural CSS :has() switch keyed on cell-vert-host (a rendered tr
    // attribute would go stale: ProseMirror reuses the tr when only children
    // change), see styles.css.
    attrs.class = attrs.class ? `${attrs.class} cell-vert-host` : 'cell-vert-host'
    const av = cellVAlignGridCss(node.attrs.vAlign as string | null)
    return [tag, attrs, ['div', { class: 'cell-vert', style: av ? `${wm};${av}` : wm }, 0]]
  }
  const clipStyle = cellClipStyle(node.attrs.vAlign, clip)
  return [
    tag,
    attrs,
    ['div', { class: 'cell-clip', style: wm ? `${clipStyle};${wm}` : clipStyle }, 0],
  ]
}

export type TableBordersAttr = Partial<
  Record<
    'top' | 'left' | 'bottom' | 'right' | 'insideH' | 'insideV',
    { style: string; szEighths?: number; color?: string }
  >
>

/**
 * Table-level w:tblBorders → CSS variables consumed by edge/inside cell rules
 * (--doc-b-t/r/b/l on edge cells beat the inside lines even when the frame is
 * explicitly none). Undeclared = no borders, matching Word's printed output.
 */
export function tableBordersCss(b: TableBordersAttr | null): string[] {
  if (!b) return []
  const styles: string[] = []
  const edge = { top: 't', right: 'r', bottom: 'b', left: 'l' } as const
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    styles.push(`--doc-b-${edge[side]}:${borderLineCss(b[side]) ?? 'none'}`)
  }
  styles.push(`--doc-b-h:${borderLineCss(b.insideH) ?? 'none'}`)
  styles.push(`--doc-b-v:${borderLineCss(b.insideV) ?? 'none'}`)
  // vertical line widths feed the cell text inset (styles.css)
  styles.push(
    `--doc-bw-l:${borderWidthPx(b.left)}px`,
    `--doc-bw-r:${borderWidthPx(b.right)}px`,
    `--doc-bw-v:${borderWidthPx(b.insideV)}px`,
  )
  // horizontal line snap deltas feed the row-height spacer (styles.css)
  styles.push(
    `--doc-bd-t:${borderSnapDeltaPx(b.top).toFixed(3)}px`,
    `--doc-bd-b:${borderSnapDeltaPx(b.bottom).toFixed(3)}px`,
    `--doc-bd-h:${borderSnapDeltaPx(b.insideH).toFixed(3)}px`,
  )
  // --dk-tb-* twins: the .page-dark cell rules read these instead (dark-page.ts)
  return styles.concat(dkTableBorders(styles))
}

/** half of the outer left+right borders: the collapsed table box spans them on top of the grid width */
export function outerBorderPx(b: TableBordersAttr | null): number {
  return b ? (borderWidthPx(b.left) + borderWidthPx(b.right)) / 2 : 0
}

/** Table-level w:tblCellMar → per-side --doc-cell-pad-* declarations; undeclared sides use Word defaults (0 top/bottom, 108 twips left/right) */
export function cellPadCss(
  mar: { top?: number; right?: number; bottom?: number; left?: number } | null,
): string[] {
  if (!mar) return []
  const px = (v: number | undefined, dflt: number) => ((v ?? dflt) / 15).toFixed(1)
  return [
    `--doc-cell-pad-t:${px(mar.top, 0)}px`,
    `--doc-cell-pad-r:${px(mar.right, 108)}px`,
    `--doc-cell-pad-b:${px(mar.bottom, 0)}px`,
    `--doc-cell-pad-l:${px(mar.left, 108)}px`,
  ]
}

/**
 * Margins of a start-aligned bidiVisual table: Word hangs it from the RIGHT
 * margin and measures w:tblInd from that side (negative indents push it into
 * the margin, over-wide tables spill left). widthExpr null = auto width.
 */
export function rtlStartMarginCss(widthExpr: string | null, indentPx: number): string[] {
  const indent = `${indentPx.toFixed(1)}px`
  if (!widthExpr) return ['margin-left:auto', `margin-right:${indent}`]
  const shift = indentPx < 0 ? `+ ${(-indentPx).toFixed(1)}px` : `- ${indent}`
  return [`margin-left:calc(var(--doc-content-w,100%) - ${widthExpr} ${shift})`]
}

export const DocTable = Node.create({
  name: 'docTable',
  group: 'block',
  content: 'docTableRow+',
  isolating: true,
  selectable: true,
  draggable: true,
  addAttributes() {
    return {
      docxIndex: { default: null as number | null },
      colWidthsPct: { default: null as number[] | null },
      widthPx: { default: null as number | null },
      widthPct: { default: null as number | null },
      cellMar: { default: null as Record<string, number> | null },
      /** w:tblCellSpacing (twips, half the inter-cell gap) → CSS border-spacing */
      cellSpacingTwips: { default: null as number | null },
      /** table shading (tblPr w:shd), hex without '#' */
      tblFill: { default: null as string | null },
      cellMarEdited: { default: false },
      borders: { default: null as Record<string, unknown> | null },
      tblAlign: { default: null as string | null },
      tblFloat: { default: null as string | null },
      tblFloatSource: { default: null as string | null },
      tblFloatSuppressed: { default: false },
      tblFloatXTwips: { default: null as number | null },
      tblFloatYTwips: { default: null as number | null },
      tblFloatHorzAnchor: { default: null as string | null },
      tblFloatVertAnchor: { default: null as string | null },
      tblFloatDistance: { default: null as Record<string, number> | null },
      /** display-only measured width used to place right AutoFit floats */
      tblFloatWidthPx: { default: null as number | null },
      tblFloatEdited: { default: false },
      tblAutoFit: { default: 'fixed' as 'contents' | 'window' | 'fixed' },
      tblAutoFitEdited: { default: false },
      /** literal w:tblLayout fixed: declared widths hold, no fit-to-page narrowing */
      tblFixedLayout: { default: false },
      indentTwips: { default: null as number | null },
      tblStyleId: { default: null as string | null },
      tblLook: { default: null as Record<string, boolean> | null },
      tblLookEdited: { default: false },
      /** SDT shell JSON when the table is a content-control member (chrome hit-testing) */
      sdtShell: { default: null as string | null },
      /** RTL table (tblPr w:bidiVisual): columns right to left */
      bidiVisual: { default: false },
      originalStructure: { default: null as string | null },
      originalFormatting: { default: null as string | null },
      blockRevision: { default: null as Record<string, string> | null },
    }
  },
  parseHTML() {
    return [{ tag: 'table.doc-table' }, { tag: 'table' }]
  },
  renderHTML({ node }) {
    const attrs: Record<string, string> = { class: 'doc-table' }
    if (node.attrs.docxIndex !== null) attrs['data-idx'] = String(node.attrs.docxIndex)
    if (node.attrs.tblStyleId) attrs['data-tbl-style'] = String(node.attrs.tblStyleId)
    const autoFit = node.attrs.tblAutoFit as 'contents' | 'window' | 'fixed'
    // Imported auto-layout tables may carry a display-only expanded width from
    // the fidelity pass. Keep that measured grid until the user explicitly
    // chooses AutoFit Contents (the command clears widthPx).
    const displayAutoFit = autoFit === 'contents' && node.attrs.widthPx ? 'fixed' : autoFit
    attrs.class += ` doc-table-autofit-${displayAutoFit}`
    const tblFloated =
      !node.attrs.tblFloatSuppressed &&
      (node.attrs.tblFloat === 'left' || node.attrs.tblFloat === 'right')
    if (tblFloated) attrs.class += ` doc-table-float-${node.attrs.tblFloat}`
    if (node.attrs.bidiVisual) attrs.dir = 'rtl'
    const styles: string[] = []
    let centerMargin: string | null = null
    // --doc-content-w: per-block section content width (differing-width sections); defaults to the page content box
    const contentW = 'var(--doc-content-w,100%)'
    // bidiVisual start alignment (w:jc absent/left): the table hangs from the
    // RIGHT margin and w:tblInd mirrors to that side, spilling left when over-wide
    const rtlStart =
      node.attrs.bidiVisual === true &&
      !tblFloated &&
      node.attrs.tblAlign !== 'center' &&
      node.attrs.tblAlign !== 'right'
    const spillMargin = rtlStart ? 'var(--doc-margin-left,0px)' : 'var(--doc-margin-right,0px)'
    let widthExpr: string | null = null
    if (displayAutoFit === 'contents') styles.push('width:auto')
    // 'window' (w:tblLayout autofit with a full-width grid) and w:tblW type="pct"
    // are both percentages of the section's TEXT COLUMN, not of the canvas' padding
    // box: sections that disagree on margins/page width pad differently, so resolve
    // through --doc-content-w (set per block by sectionWidthSpecs; unset → the old %)
    else if (displayAutoFit === 'window') {
      widthExpr = contentW
      styles.push(`width:${widthExpr}`)
    } else if (node.attrs.widthPct) {
      widthExpr = `calc(${contentW} * ${Number(node.attrs.widthPct) / 100})`
      styles.push(`width:${widthExpr}`)
    }
    // Over-wide grids may spill into the page margins like Word/LO (clamping
    // them to the content box narrowed every column, wrapped cell text onto extra
    // lines and inflated PDF-converted documents by pages), but never past the paper:
    // centered tables spill both margins symmetrically (negative-margin centering —
    // auto margins resolve to 0 on overflow and would push the spill right only),
    // left-aligned ones spill right; indent comes out of the spill allowance
    else if (node.attrs.widthPx) {
      // the collapsed box also spans the outer half-borders (Word straddles the outer gridlines)
      const widthPx =
        Number(node.attrs.widthPx) +
        (node.attrs.cellSpacingTwips
          ? 0
          : outerBorderPx(node.attrs.borders as TableBordersAttr | null))
      // w:tblLayout fixed holds the declared widths even past the paper edge (Word
      // clips there); narrowing to fit would rewrap every column (prod100 sas 045)
      const holdWidth = node.attrs.tblFixedLayout === true
      if (holdWidth) {
        widthExpr = `${widthPx}px`
        styles.push(`width:${widthExpr}`, 'max-width:none')
        if (node.attrs.tblAlign === 'center' && !tblFloated)
          centerMargin = `margin-left:calc((${contentW} - ${widthPx}px)/2)`
      } else if (node.attrs.tblAlign === 'center' && !tblFloated) {
        const paper = `calc(${contentW} + var(--doc-margin-left,var(--doc-margin-right,0px)) + var(--doc-margin-right,0px))`
        styles.push(`width:min(${widthPx}px,${paper})`)
        centerMargin = `margin-left:calc((${contentW} - min(${widthPx}px,${paper}))/2)`
      } else {
        const indented =
          !tblFloated && node.attrs.tblAlign !== 'right' && Number(node.attrs.indentTwips) > 0
        const indentPx = indented ? Number(node.attrs.indentTwips) / 15 : 0
        const spill = indentPx
          ? `calc(${contentW} + ${spillMargin} - ${indentPx.toFixed(1)}px)`
          : `calc(${contentW} + ${spillMargin})`
        widthExpr = `min(${widthPx}px,${spill})`
        styles.push(`width:${widthExpr}`)
      }
    }
    styles.push(...cellPadCss(node.attrs.cellMar as Record<string, number> | null))
    // w:tblCellSpacing: each cell contributes the value on its side, so the CSS
    // gap between cells is twice it; cells render individually boxed like Word
    if (node.attrs.cellSpacingTwips) {
      const gapPx = ((Number(node.attrs.cellSpacingTwips) * 2) / 15).toFixed(1)
      // separate borders sit whole inside each cell: no half-border to absorb
      styles.push(
        'border-collapse:separate',
        `border-spacing:${gapPx}px`,
        '--doc-bw-share:0',
        '--doc-bd-share:1',
      )
    }
    if (node.attrs.tblFill) {
      styles.push(`background-color:#${node.attrs.tblFill}`, dkBackground(`#${node.attrs.tblFill}`))
      {
        const ink = fillInk(node.attrs.tblFill)
        if (ink) attrs['data-ink'] = ink
      }
    }
    styles.push(...tableBordersCss(node.attrs.borders as TableBordersAttr | null))
    styles.push(
      ...tableRowEatCss(
        node.attrs.borders as TableBordersAttr | null,
        Boolean(node.attrs.cellSpacingTwips),
      ),
    )
    // w:tblpPr positioning supersedes w:jc / w:tblInd: alignment or indent
    // margins would override the float stylesheet's wrap gaps
    if (tblFloated) {
      const distance = (node.attrs.tblFloatDistance as Record<string, number> | null) ?? {}
      const px = (twips: unknown): number => (Number(twips) || 0) / 15
      const x = px(node.attrs.tblFloatXTwips)
      const y = px(node.attrs.tblFloatYTwips)
      // w:tblpY under vertAnchor="page"/"margin" is a position on the landing
      // page, not a flow offset (Word probe prod100r2/36: a missing vertAnchor
      // is text-relative). The pagination engine resolves the page-relative
      // target and applies the shift via --tblp-dy; flow CSS gets no margin.
      const vAnchor = node.attrs.tblFloatVertAnchor
      const pageRelV =
        (vAnchor === 'page' || vAnchor === 'margin') && node.attrs.tblFloatYTwips != null
      if (pageRelV) {
        attrs['data-tblp-vy'] = y.toFixed(1)
        attrs['data-tblp-vanchor'] = String(vAnchor)
        styles.push('margin-top:var(--tblp-dy,0px)')
      }
      const top = pageRelV ? 0 : y + Math.max(0, px(distance.top))
      const bottom = Math.max(0, px(distance.bottom))
      const left = Math.max(0, px(distance.left))
      const right = Math.max(0, px(distance.right))
      if (top) styles.push(`margin-top:${top.toFixed(1)}px`)
      if (bottom) styles.push(`margin-bottom:${bottom.toFixed(1)}px`)
      // w:tblpX with horzAnchor="page" measures from the PAGE edge, not the
      // content box — subtract the left margin. And the offset is CLAMPED so
      // the table never hangs past the right content edge: unclamped
      // page-anchored deal-doc captables rendered half off-page (alpha
      // ledger, #genoffice-feedback task #6). Word keeps floats on the page.
      const fromPageEdge = node.attrs.tblFloatHorzAnchor === 'page'
      const tblWidth = Number(node.attrs.widthPx) || Number(node.attrs.tblFloatWidthPx) || 0
      const xExpr = fromPageEdge
        ? `calc(${x.toFixed(1)}px - var(--doc-margin-left,0px))`
        : `${x.toFixed(1)}px`
      if (node.attrs.tblFloat === 'left') {
        if (x) {
          const capped =
            tblWidth > 0
              ? `min(${xExpr},calc(var(--doc-content-w,100%) - ${tblWidth.toFixed(1)}px))`
              : xExpr
          styles.push(`margin-left:max(0px,${capped})`)
        }
        if (right) styles.push(`margin-right:${right.toFixed(1)}px`)
      } else {
        if (left) styles.push(`margin-left:${left.toFixed(1)}px`)
        if (node.attrs.tblFloatXTwips != null && tblWidth > 0) {
          styles.push(
            `margin-right:max(0px,calc(var(--doc-content-w,100%) - ${xExpr} - ${tblWidth.toFixed(1)}px))`,
          )
        }
      }
    } else if (node.attrs.tblAlign === 'center') {
      if (centerMargin) styles.push(centerMargin)
      else styles.push('margin-left:auto', 'margin-right:auto')
    } else if (node.attrs.tblAlign === 'right') styles.push('margin-left:auto')
    else if (rtlStart) {
      styles.push(...rtlStartMarginCss(widthExpr, Number(node.attrs.indentTwips ?? 0) / 15))
    } else if (node.attrs.indentTwips) {
      styles.push(`margin-left:${(Number(node.attrs.indentTwips) / 15).toFixed(1)}px`)
    }
    // a suppressed text-anchored float with a negative w:tblpY still hangs that far
    // above its anchor paragraph in Word (cover logo strips reach into the top
    // margin), so the inline table keeps the lift as a negative top margin
    const liftTwips =
      node.attrs.tblFloatSuppressed &&
      (node.attrs.tblFloatVertAnchor ?? 'text') === 'text' &&
      Number(node.attrs.tblFloatYTwips) < 0
        ? -Number(node.attrs.tblFloatYTwips)
        : 0
    if (liftTwips > 0) {
      const lift = (liftTwips / 15).toFixed(1)
      styles.push(`margin-top:-${lift}px`)
      attrs['data-tblp-lift'] = lift
    }
    if (styles.length > 0) attrs.style = styles.join(';')
    // A colgroup with normalized percentages defines the column grid whenever the
    // pct list matches the grid, so a table clamped to the content box compresses
    // its columns proportionally instead of overflowing via fixed td px widths.
    let firstRowCols = 0
    node.firstChild?.forEach((cell) => {
      firstRowCols += Number(cell.attrs.colspan) || 1
    })
    const rawPct = node.attrs.colWidthsPct as number[] | null
    if (displayAutoFit !== 'contents' && rawPct?.length) {
      // zero-width grid slots get a small floor, short grids pad with the average —
      // dropping the whole colgroup falls back to fixed-layout even splitting, which
      // is always worse than an approximate grid
      const pct = rawPct.map((w) => (w > 0 ? w : 0.5))
      const avg = pct.reduce((sum, w) => sum + w, 0) / pct.length
      while (pct.length < firstRowCols) pct.push(avg)
      const total = pct.reduce((sum, w) => sum + w, 0) || 100
      return [
        'table',
        attrs,
        [
          'colgroup',
          {},
          ...pct.map(
            (w) => ['col', { style: `width:${((w / total) * 100).toFixed(2)}%` }] as const,
          ),
        ],
        ['tbody', 0],
      ]
    }
    return ['table', attrs, ['tbody', 0]]
  },
})

export const DocTableRow = Node.create({
  name: 'docTableRow',
  content: '(docTableCell | docTableHeader)+',
  addAttributes() {
    return {
      heightTwips: { default: null as number | null },
      heightRule: { default: null as 'atLeast' | 'exact' | null },
      repeatHeader: { default: false },
      repeatHeaderEdited: { default: false },
      rawTrPr: { default: null as string | null },
      /** trPr w:ins/w:del row-level revision ({kind, author, ...} | null) */
      rowRevision: { default: null as Record<string, string> | null },
    }
  },
  parseHTML() {
    return [{ tag: 'tr' }]
  },
  renderHTML({ node }) {
    const h = node.attrs.heightTwips as number | null
    const rev = node.attrs.rowRevision as { kind?: string; author?: string } | null
    const attrs: Record<string, string> = {}
    const classes: string[] = []
    if (h) {
      const cellBorders: Array<Record<string, BorderLine> | null> = []
      node.forEach((cell) => {
        if (!cell.attrs.gridGap)
          cellBorders.push(cell.attrs.borders as Record<string, BorderLine> | null)
      })
      attrs.style = rowHeightCss(h, cellBorders)
      if (node.attrs.heightRule === 'exact') classes.push('row-h-exact')
    }
    attrs['data-repeat-header'] = node.attrs.repeatHeader ? '1' : '0'
    if (rev?.kind) {
      classes.push(`row-rev-${rev.kind}`)
      if (rev.author) attrs.title = rev.author
    }
    if (classes.length > 0) attrs.class = classes.join(' ')
    return ['tr', attrs, 0]
  },
})

export const DocTableCell = Node.create({
  name: 'docTableCell',
  content: '(docParagraph | docListItem | docNestedTable | docCellBoxes)+',
  isolating: true,
  addAttributes() {
    return tableCellAttrs
  },
  parseHTML() {
    return [{ tag: 'td' }]
  },
  renderHTML({ node }) {
    return tableCellSpec('td', node)
  },
})

export const DocTableHeader = Node.create({
  name: 'docTableHeader',
  content: '(docParagraph | docListItem | docNestedTable | docCellBoxes)+',
  isolating: true,
  addAttributes() {
    return tableCellAttrs
  },
  parseHTML() {
    return [{ tag: 'th' }]
  },
  renderHTML({ node }) {
    return tableCellSpec('th', node)
  },
})

/** Anchored shapes/textboxes inside a table cell (display-only): Word renders them
 *  in the cell and grows the row to hold them, so the wrapper takes the boxes'
 *  bottom extent as in-flow height (pagination then pushes the row like Word) */
export const DocCellBoxes = Node.create({
  name: 'docCellBoxes',
  atom: true,
  selectable: false,
  addAttributes() {
    return { boxes: { default: null as TextboxDisplay[] | null } }
  },
  parseHTML() {
    return []
  },
  renderHTML({ node }) {
    const boxes = (node.attrs.boxes as TextboxDisplay[] | null) ?? []
    if (boxes.length === 0) return ['div', { class: 'doc-cell-boxes' }]
    let bottom = 0
    for (const b of boxes) {
      bottom = Math.max(
        bottom,
        (b.offsetYEmu ?? 0) / EMU_PER_PX + (b.heightPx ?? b.minHeightPx ?? 0),
      )
    }
    return [
      'div',
      {
        class: 'doc-cell-boxes',
        contenteditable: 'false',
        // zero-width leading float: the cell (a BFC) grows to max(text, boxes)
        // like Word, without displacing the cell's own text
        style: bottom > 0 ? `height:${bottom.toFixed(1)}px` : '',
      },
      // floating boxes self-position from their anchor offsets; in-flow boxes get
      // the same offsets via a positioned wrapper (Word places both in the cell)
      ...boxes.map((b): DomSpec => {
        const spec = renderTextboxSpec(b)
        if (b.floating) return spec
        const left = (b.offsetXEmu ?? 0) / EMU_PER_PX
        const top = (b.offsetYEmu ?? 0) / EMU_PER_PX
        return [
          'div',
          { style: `position:absolute;left:${left.toFixed(1)}px;top:${top.toFixed(1)}px` },
          spec,
        ]
      }),
    ]
  },
})

/** Nested table inside a cell: read-only atomic child table (editing the outer cell's text
 *  doesn't affect it; saving is byte-faithful via the outer table's originalXml, dropped on structural rebuild — matching old behavior) */
export const DocNestedTable = Node.create({
  name: 'docNestedTable',
  atom: true,
  selectable: false,
  addAttributes() {
    return { model: { default: null as TableModel | null } }
  },
  parseHTML() {
    return []
  },
  renderHTML({ node }) {
    const model = node.attrs.model as TableModel | null
    if (!model?.rows?.length) return ['div', { class: 'doc-nested-table' }]
    return [
      'div',
      { class: 'doc-nested-table', contenteditable: 'false' },
      renderTableSpec(model, true),
    ]
  },
  // in-place cell editing: contenteditable island; on blur the text is committed back to the model attribute
  // (saving goes through the outer table's nested-text surgical patch; cells that themselves contain nested tables stay non-editable)
  addNodeView() {
    return ({ node, editor, getPos }) => {
      let currentNode = node
      const dom = document.createElement('div')
      dom.className = 'doc-nested-table'
      dom.setAttribute('contenteditable', 'false')
      // only cells the user typed in are read back: the rendered text is not a
      // faithful encoding of the model (trailing empty paragraphs, tabs, breaks)
      const dirty = new Set<HTMLElement>()
      dom.addEventListener('input', (e) => {
        const td = (e.target as HTMLElement | null)?.closest?.('td')
        if (td) dirty.add(td)
      })

      /** this table's own td elements (excluding deeper nested-table td), ordered like the model's non-vMerge-continue cells */
      const ownTds = (): HTMLElement[] => {
        const root = dom.querySelector('table')
        if (!root) return []
        return Array.from(dom.querySelectorAll('td')).filter((td) => td.closest('table') === root)
      }

      const flatCells = (): TableCell[] => {
        const model = currentNode.attrs.model as TableModel | null
        const flat: TableCell[] = []
        model?.rows.forEach((row) =>
          row.forEach((cell) => {
            if (cell.vMerge !== 'continue') flat.push(cell)
          }),
        )
        return flat
      }

      const applyEditable = () => {
        const cells = flatCells()
        ownTds().forEach((td, i) => {
          const editable = editor.isEditable && cells[i] && !cells[i].nestedTables?.length
          td.setAttribute('contenteditable', editable ? 'true' : 'false')
        })
      }

      const render = () => {
        dom.innerHTML = ''
        dirty.clear()
        const model = currentNode.attrs.model as TableModel | null
        if (model?.rows?.length) {
          const rendered = DOMSerializer.renderSpec(document, renderTableSpec(model, true) as never)
          dom.appendChild(rendered.dom)
        }
        applyEditable()
      }
      render()

      const commit = () => {
        const model = currentNode.attrs.model as TableModel | null
        if (!model) return
        const tds = ownTds()
        let k = 0
        let changed = false
        const rows = model.rows.map((row) =>
          row.map((cell) => {
            if (cell.vMerge === 'continue') return cell
            const td = tds[k++]
            if (!td || cell.nestedTables?.length || !dirty.has(td)) return cell
            const paras = tdParas(td)
            if (paras.join('\n') === cell.paras.join('\n')) return cell
            changed = true
            const next = { ...cell, paras }
            delete next.richParas
            return next
          }),
        )
        dirty.clear()
        if (!changed) return
        const pos = getPos()
        if (typeof pos !== 'number') return
        editor.view.dispatch(
          editor.view.state.tr.setNodeMarkup(pos, undefined, { model: { ...model, rows } }),
        )
      }

      const onFocusOut = (e: Event) => {
        const next = (e as FocusEvent).relatedTarget as HTMLElement | null
        if (next && dom.contains(next)) return
        commit()
      }
      dom.addEventListener('focusout', onFocusOut)
      window.addEventListener('ai-docs-commit-tables', commit)
      editor.on('update', applyEditable)

      return {
        dom,
        update: (n: PmNode) => {
          if (n.type.name !== 'docNestedTable') return false
          if (!n.eq(currentNode)) {
            currentNode = n
            render()
          } else {
            currentNode = n
          }
          return true
        },
        // edits stay in the DOM until the focusout commit; don't let ProseMirror re-parse them
        ignoreMutation: () => true,
        stopEvent: (event: Event) => {
          const target = event.target as HTMLElement | null
          return !!target?.closest?.('td[contenteditable="true"]')
        },
        destroy: () => {
          window.removeEventListener('ai-docs-commit-tables', commit)
          editor.off('update', applyEditable)
        },
      }
    }
  },
})

/** Delete only an explicitly selected whole table; leave cursors and partial cell selections alone. */
export function deleteSelectedWholeTable(
  state: EditorState,
  dispatch?: (transaction: Transaction) => void,
): boolean {
  const { selection } = state
  if (selection instanceof NodeSelection) {
    if (selection.node.type.spec.tableRole !== 'table') return false
    dispatch?.(state.tr.delete(selection.from, selection.to).scrollIntoView())
    return true
  }
  if (
    selection instanceof CellSelection &&
    selection.isRowSelection() &&
    selection.isColSelection()
  ) {
    return deleteTable(state, dispatch)
  }
  return false
}

export const NativeTableSupport = Extension.create({
  name: 'nativeTableSupport',
  extendNodeSchema(extension) {
    if (extension.name === 'docTable') return { tableRole: 'table' }
    if (extension.name === 'docTableRow') return { tableRole: 'row' }
    if (extension.name === 'docTableCell') return { tableRole: 'cell' }
    if (extension.name === 'docTableHeader') return { tableRole: 'header_cell' }
    return {}
  },
  addKeyboardShortcuts() {
    const deleteWholeTable = () =>
      deleteSelectedWholeTable(this.editor.state, this.editor.view.dispatch)
    return {
      Tab: () => {
        const { view } = this.editor
        if (goToNextCell(1)(this.editor.state, view.dispatch)) return true
        // last cell: Word appends a row and moves into its first cell
        if (!isInTable(this.editor.state)) return false
        if (!addRowAfter(this.editor.state, view.dispatch)) return false
        return goToNextCell(1)(this.editor.state, view.dispatch)
      },
      'Shift-Tab': () => goToNextCell(-1)(this.editor.state, this.editor.view.dispatch),
      Backspace: deleteWholeTable,
      Delete: deleteWholeTable,
    }
  },
  addProseMirrorPlugins() {
    return [
      new Plugin({
        props: {
          handleDOMEvents: {
            mouseup: (view) => {
              const resizeState = columnResizingPluginKey.getState(view.state) as
                { dragging?: unknown; activeHandle?: number } | undefined
              if (resizeState?.dragging == null) return false
              const handle = resizeState.activeHandle ?? -1
              if (handle < 0) return false
              // The resize plugin commits its final width in its window-level mouseup,
              // which runs after this handler and its microtasks; wait a macrotask so
              // the committed grid, not the previous drag frame, is constrained.
              window.setTimeout(() => {
                const raw = getComputedStyle(view.dom).getPropertyValue('--section-content-w')
                const maxWidth = Number.parseFloat(raw)
                if (Number.isFinite(maxWidth) && maxWidth > 0) {
                  constrainTableWidthAtCell(handle, maxWidth)(view.state, view.dispatch)
                }
              }, 0)
              return false
            },
          },
        },
      }),
      columnResizing({ View: null, cellMinWidth: 40, lastColumnResizable: true }),
      tableEditing({ allowTableNodeSelection: true }),
    ]
  },
})

/** Protected whole-unit blocks: images, passthrough (charts, math, ...). */
/**
 * Copying an embedded picture must put a REAL bitmap on the OS clipboard:
 * ProseMirror's HTML-only write left external apps with nothing to paste
 * (Gmail: blank) and other documents with a placeholder shell (r136). The
 * main process writes image + <img> html; in-document paste rebuilds the
 * picture through DocProtected's img parse rule.
 */
/** formats pmDocToSavePlan can rebuild into the docx (see imageFromProtectedAttrs) */
const PERSISTABLE_IMAGE_URL = /^data:image\/(?:png|jpeg|gif);base64,/

/** display attrs a copied picture needs to round-trip through clipboard HTML */
const imageMetaJson = (attrs: Record<string, unknown>): string =>
  JSON.stringify({
    imageWidthPx: attrs.imageWidthPx ?? null,
    imageHeightPx: attrs.imageHeightPx ?? null,
    imageAlign: attrs.imageAlign ?? null,
    imageWrap: attrs.imageWrap ?? null,
  })

export const ImageCopyExtension = Extension.create({
  name: 'imageClipboardCopy',
  addProseMirrorPlugins() {
    const copyImage = (view: EditorView, event: ClipboardEvent, cut: boolean): boolean => {
      const sel = view.state.selection
      if (!(sel instanceof NodeSelection)) return false
      const node = sel.node
      if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') return false
      let dataUrl = node.attrs.imageDataUrl
      if (typeof dataUrl !== 'string' || !dataUrl.startsWith('data:image/')) return false
      // the save pipeline persists only png/jpeg/gif; display-only formats
      // (bmp/webp/svg/tiff...) are transcoded to PNG from the already-decoded
      // DOM img so the copy stays saveable — undecodable ones keep the
      // default HTML copy
      if (!PERSISTABLE_IMAGE_URL.test(dataUrl)) {
        const dom = view.nodeDOM(sel.from) as HTMLElement | null
        const img = dom?.querySelector?.('img.doc-protected-img') as HTMLImageElement | null
        if (!img || !img.naturalWidth || !img.naturalHeight) return false
        const canvas = document.createElement('canvas')
        canvas.width = img.naturalWidth
        canvas.height = img.naturalHeight
        const ctx = canvas.getContext('2d')
        if (!ctx) return false
        ctx.drawImage(img, 0, 0)
        try {
          dataUrl = canvas.toDataURL('image/png')
        } catch {
          return false
        }
      }
      const write = window.desktop?.copyImageToClipboard?.(dataUrl, imageMetaJson(node.attrs))
      // no bridge: keep the default HTML-only copy instead of an empty clipboard
      if (!write) return false
      event.preventDefault()
      // cut deletes only AFTER the clipboard write is confirmed, and only if
      // the same picture is still selected — a failed write must not lose data
      void write.then((ok) => {
        if (!ok || !cut || !view.editable) return
        const cur = view.state.selection
        if (cur instanceof NodeSelection && cur.from === sel.from && cur.node === node) {
          view.dispatch(view.state.tr.deleteSelection().scrollIntoView())
        }
      })
      return true
    }
    return [
      new Plugin({
        key: new PluginKey('imageClipboardCopy'),
        props: {
          handleDOMEvents: {
            copy: (view, event) => copyImage(view, event as ClipboardEvent, false),
            cut: (view, event) => copyImage(view, event as ClipboardEvent, true),
          },
        },
      }),
    ]
  },
})

export const DocProtected = Node.create({
  name: 'docProtected',
  group: 'block',
  atom: true,
  // Editable descendants switch this off on pointer-down; the explicit handle
  // switches it back on for whole-object movement.
  draggable: true,
  selectable: true,
  addAttributes() {
    return {
      docxIndex: { default: null as number | null },
      blockRevision: { default: null as Record<string, string> | null },
      blockType: { default: 'passthrough' },
      /** w:pStyle of field/TOC paragraphs: doc style CSS (spacing/line-height) targets data-style */
      styleId: { default: null as string | null },
      label: { default: '' },
      previewText: { default: '' },
      imageDataUrl: { default: null as string | null },
      oleProgId: { default: null as string | null },
      /** display size in CSS px (blockType === 'image'), editable via drag handles */
      imageWidthPx: { default: null as number | null },
      imageHeightPx: { default: null as number | null },
      /** source crop (a:srcRect) fractions, display-only */
      imageCrop: { default: null as { l: number; t: number; r: number; b: number } | null },
      /** fill placement (a:fillRect) fractions (negative = bleed), display-only */
      imageFillRect: { default: null as { l: number; t: number; r: number; b: number } | null },
      imageLeadingText: { default: null as string | null },
      imageLeadingFont: { default: null as string | null },
      imageLeadingExplicitSpaceWidthPx: { default: null as number | null },
      imageLeadingImplicitSpaceCount: { default: null as number | null },
      imageParagraphIndentLeft: { default: null as number | null },
      imageParagraphIndentRight: { default: null as number | null },
      imageParagraphIndentFirstLine: { default: null as number | null },
      /** paragraph alignment of the image (w:jc) */
      imageAlign: { default: null as string | null },
      imageWrap: { default: null as string | null },
      /** side wrap with no room for text beside (floating table): band at the offset */
      imageBand: { default: false },
      imageWrapDistTopEmu: { default: null as number | null },
      imageWrapDistBottomEmu: { default: null as number | null },
      imageWrapDistLeftEmu: { default: null as number | null },
      imageWrapDistRightEmu: { default: null as number | null },
      /**
       * Stacking rank of a floating image among overlapping anchors
       * (bring-to-front / send-to-back). Written to the anchor's
       * relativeHeight; higher paints in front. Null = base level.
       */
      imageZOrder: { default: null as number | null },
      /**
       * Free-position offset (EMU) of a floating image with wp:posOffset.
       * Used for drag-to-reposition. Null when the image uses named alignment
       * or is inline.
       */
      imageOffsetXEmu: { default: null as number | null },
      /** wp:anchor locked="1": keep the anchor paragraph fixed while dragging */
      imageAnchorLocked: { default: false },
      /** margin-relative wp:align preset (Word position gallery) */
      imagePosH: { default: null as string | null },
      imagePosV: { default: null as string | null },
      imageOffsetYEmu: { default: null as number | null },
      /** display-only table structure (blockType === 'table') */
      table: { default: null as TableModel | null },
      /** display-only rendering for field passthrough paragraphs */
      fieldDisplay: { default: null as FieldDisplay | null },
      /** decorative rule drawing: render as a horizontal line, not a chip */
      decorative: { default: false },
      /** stroke display of a decorative rule; null = default 1px full-width line */
      ruleColorHex: { default: null as string | null },
      ruleThicknessPx: { default: null as number | null },
      ruleWidthPx: { default: null as number | null },
      /** broken picture (missing rel/media): empty frame + alt text at the declared extent */
      brokenImage: { default: false },
      /** invisible body-level marker (stray bookmarkEnd…): render nothing, keep position */
      invisibleMarker: { default: false },
      /** paragraph attrs of the anchor paragraph's own empty line (display only) */
      anchorLine: { default: null as Record<string, unknown> | null },
      /** display-only anchored textboxes (code boxes, callout cards) */
      textboxes: { default: null as TextboxDisplay[] | null },
      anchorSnapToGrid: { default: null as false | null },
      /** the anchor paragraph's own runs next to content textboxes (display-only) */
      strayRuns: { default: null as Run[] | null },
      strayStyleId: { default: null as string | null },
      strayIndent: { default: null as StrayIndent | null },
      strayAlign: { default: null as string | null },
      /** numbering reference of the anchor paragraph (ListNumberingExtension counts it and sets --li-marker) */
      strayList: { default: null as ListItemRef | null },
      /** editable OMML leaf tokens; formula structure remains protected */
      formulaDisplay: { default: null as FormulaDisplay | null },
      /** embedded chart data model; cached texts/numbers editable, structure protected */
      chartDisplay: { default: null as ChartDisplay | null },
      /** display-only SmartArt degrade (precomputed diagram drawing shapes) */
      diagramDisplay: { default: null as DiagramDisplay | null },
      /** self-contained OOXML fragment for editor-created content (new tables) */
      genXml: { default: null as string | null },
      /** new image awaiting embedding at save time */
      genImage: {
        default: null as { base64: string; mime: string; widthPx: number; heightPx: number } | null,
      },
      /** picture rotation (deg clockwise, 0-359) and mirror flips (a:xfrm rot/flipH/flipV) */
      imageRotDeg: { default: null as number | null },
      imageFlipH: { default: false },
      imageFlipV: { default: false },
      /** picture outline (pic:spPr a:ln solid fill, display-only) */
      imageBorder: { default: null as { color: string; widthPt: number } | null },
      /** replacement bytes for an original image (crop/background removal/replace):
       *  the drawing XML — and with it docxIndex, wrap and position — survives */
      imageReplace: { default: null as { base64: string; mime: string } | null },
      /** new chart awaiting embedding at save time (data snapshot; edits live in chartDisplay) */
      genChart: { default: null as NewChart | null },
    }
  },
  parseHTML() {
    /** a copied picture rebuilds from its inner img + meta payload; other
     *  protected kinds keep the default attrs (their payloads cannot travel
     *  through HTML) — r136 */
    const imageAttrsFrom = (el: HTMLElement): Record<string, unknown> | null => {
      const img = el.querySelector?.('img.doc-protected-img') as HTMLImageElement | null
      const src = img?.getAttribute('src') ?? ''
      // persistable formats only: a bmp/webp/svg picture would display and
      // then silently vanish on save — the placeholder shell is honest
      if (!PERSISTABLE_IMAGE_URL.test(src)) return null
      let meta: Record<string, unknown> = {}
      try {
        meta = JSON.parse(el.getAttribute('data-image-meta') ?? '{}') as Record<string, unknown>
      } catch {
        /* stripped payload: size falls back below */
      }
      const num = (v: unknown) => (typeof v === 'number' && Number.isFinite(v) ? v : null)
      const str = (v: unknown) => (typeof v === 'string' && v ? v : null)
      return {
        blockType: 'image',
        label: 'Image',
        imageDataUrl: src,
        imageWidthPx: num(meta.imageWidthPx) ?? num(img?.width) ?? null,
        imageHeightPx: num(meta.imageHeightPx) ?? num(img?.height) ?? null,
        imageAlign: str(meta.imageAlign),
        imageWrap: str(meta.imageWrap),
      }
    }
    return [
      {
        tag: 'div[data-doc-protected]',
        getAttrs: (el) => imageAttrsFrom(el as HTMLElement),
      },
      // bare data-URL <img> (our own image-copy clipboard html, or rich
      // sources that inline the bitmap); http images keep going through the
      // dedicated insert paths, so this rule stays data:-only
      {
        tag: 'img[src]',
        getAttrs: (el) => {
          const img = el as HTMLImageElement
          const src = img.getAttribute('src') ?? ''
          if (!PERSISTABLE_IMAGE_URL.test(src)) return false
          let meta: Record<string, unknown> = {}
          try {
            meta = JSON.parse(img.getAttribute('data-image-meta') ?? '{}') as Record<
              string,
              unknown
            >
          } catch {
            /* foreign img without our payload: presentation size below */
          }
          const num = (v: unknown) => (typeof v === 'number' && Number.isFinite(v) ? v : null)
          const str = (v: unknown) => (typeof v === 'string' && v ? v : null)
          return {
            blockType: 'image',
            label: 'Image',
            imageDataUrl: src,
            imageWidthPx: num(meta.imageWidthPx) ?? (img.width || null),
            imageHeightPx: num(meta.imageHeightPx) ?? (img.height || null),
            imageAlign: str(meta.imageAlign),
            imageWrap: str(meta.imageWrap),
          }
        },
      },
    ]
  },
  renderHTML({ node }) {
    return protectedDomSpec(node) as never
  },
  addNodeView() {
    return ({ node, editor, getPos }) => {
      let currentNode = node
      const dom = buildProtectedDom(currentNode)
      const getNode = () => currentNode
      const pos = getPos as () => number | undefined
      const textboxes = mountTextboxEditors(dom, getNode, pos, editor.view)
      const table = wireTableEditing(dom, getNode, pos, editor.view)
      const field = wireFieldEditing(dom, getNode, pos, editor.view)
      const formula = wireFormulaEditing(dom, getNode, pos, editor.view)
      const chart = wireChartEditing(dom, getNode, pos, editor.view)
      drawChartSvg(dom, currentNode.attrs.chartDisplay as ChartDisplay | null)
      const cleanups = [
        wireProtectedInteractionMode(
          dom,
          pos,
          editor.view,
          textboxes ?? table ?? field ?? formula ?? chart,
        ),
        wireFormulaLatexEdit(dom, getNode, pos),
        table?.cleanup,
        field?.cleanup,
        formula?.cleanup,
        chart?.cleanup,
        textboxes?.cleanup,
      ]
      return {
        dom,
        update: (n: PmNode) => {
          if (n.type.name !== 'docProtected') return false
          if (n.eq(currentNode)) {
            currentNode = n
            return true
          }
          // textbox commits only swap the textboxes attr; the sub-editors are
          // the source of truth, so keep the DOM (and editing session) alive
          if (textboxes && attrsEqualExcept(n.attrs, currentNode.attrs, 'textboxes')) {
            currentNode = n
            textboxes.sync(n.attrs.textboxes as TextboxDisplay[] | null)
            return true
          }
          // other attribute change (resize, table commit, ...): recreate DOM
          return false
        },
        // cell edits live in the DOM until committed on focusout; never re-parse
        ignoreMutation: () => true,
        stopEvent: (event: Event) => {
          const target = event.target as HTMLElement | null
          if (target?.closest?.('.doc-formula-edit')) return true
          // Let ProseMirror plugins receive handle presses; floating-object
          // dragging is implemented at the editor-view level.
          if (target?.closest?.('.doc-move-handle')) return false
          // Shape bodies (prst textboxes) drag-to-move on a plain press (Word
          // parity); a double-click still reaches the inner editor for the caret
          if (
            event.type === 'mousedown' &&
            (event as MouseEvent).detail < 2 &&
            target?.closest?.('.doc-textbox') &&
            !dom.classList.contains('doc-content-editing') &&
            (currentNode.attrs.textboxes as TextboxDisplay[] | null)?.[0]?.prst
          ) {
            return false
          }
          const contentTarget = target?.closest?.(EDITABLE_PROTECTED_SELECTOR)
          return (
            !!contentTarget &&
            (event.type === 'mousedown' || dom.classList.contains('doc-content-editing'))
          )
        },
        destroy: () => cleanups.forEach((c) => c?.()),
      }
    }
  },
  addProseMirrorPlugins() {
    return [imageResizePlugin(), floatingObjectDragPlugin()]
  },
})

/**
 * SmartArt / drawing-canvas display: absolutely positioned shapes (picture
 * fills, solid fills, centered texts) at the parse-resolved geometry. All
 * colors are document data (theme-resolved at parse), hence inline. Canvas
 * displays (lockedCanvas) keep raw-size text overflowing the scaled child
 * boxes instead of clipping, like LO renders them.
 */
function diagramSpecOf(diagram: DiagramDisplay): DomSpec {
  const shapeSpecs: DomSpec[] = diagram.shapes.map((s) => {
    // connectors (prst=line, zero cx or cy) render as solid rules along the axis
    if (s.lnHex && (s.wPx <= 0 || s.hPx <= 0)) {
      const w = Math.max(1, s.lnWPx ?? 1)
      const style = (
        s.hPx <= 0
          ? [
              `left:${s.xPx}px`,
              `top:${(s.yPx - w / 2).toFixed(1)}px`,
              `width:${s.wPx}px`,
              `height:${w}px`,
            ]
          : [
              `left:${(s.xPx - w / 2).toFixed(1)}px`,
              `top:${s.yPx}px`,
              `width:${w}px`,
              `height:${s.hPx}px`,
            ]
      )
        .concat(`background:#${s.lnHex}`)
        .join(';')
      return ['span', { class: 'doc-diagram-shape', style }] as DomSpec
    }
    const radius =
      s.prst === 'ellipse' || s.prst === 'circle'
        ? '50%'
        : s.prst === 'roundRect'
          ? `${Math.round(Math.min(s.wPx, s.hPx) * 0.12)}px`
          : '0'
    const style = [
      `left:${s.xPx}px`,
      `top:${s.yPx}px`,
      `width:${s.wPx}px`,
      `height:${s.hPx}px`,
      `border-radius:${radius}`,
      s.fillHex ? `background:#${s.fillHex}` : '',
      s.lnHex
        ? `border:${Math.max(1, s.lnWPx ?? 1)}px solid #${s.lnHex};box-sizing:border-box`
        : '',
      s.rotDeg ? `transform:rotate(${s.rotDeg}deg)` : '',
      s.fontSizePt ? `font-size:${s.fontSizePt}pt` : '',
      s.textColorHex ? `color:#${s.textColorHex}` : '',
    ]
      .filter(Boolean)
      .join(';')
    const kids: DomSpec[] = []
    if (s.imageDataUrl) {
      let imgStyle = 'position:absolute;left:0;top:0;width:100%;height:100%;object-fit:cover'
      if (s.fillRect) {
        const sw = s.wPx * (1 - s.fillRect.l - s.fillRect.r)
        const sh = s.hPx * (1 - s.fillRect.t - s.fillRect.b)
        imgStyle =
          `position:absolute;left:${(s.fillRect.l * s.wPx).toFixed(1)}px;` +
          `top:${(s.fillRect.t * s.hPx).toFixed(1)}px;` +
          `width:${sw.toFixed(1)}px;height:${sh.toFixed(1)}px;max-width:none`
      }
      kids.push(['img', { src: s.imageDataUrl, class: 'doc-diagram-img', style: imgStyle }])
    }
    if (s.texts?.length) {
      kids.push(['span', { class: 'doc-diagram-text' }, s.texts.join('\n')])
    }
    return ['span', { class: 'doc-diagram-shape', style }, ...kids]
  })
  const spanStyle = [
    `width:${diagram.widthPx}px`,
    `height:${diagram.heightPx}px`,
    diagram.floating
      ? `position:absolute;left:${((diagram.offsetXEmu ?? 0) / EMU_PER_PX).toFixed(1)}px;` +
        `top:${((diagram.offsetYEmu ?? 0) / EMU_PER_PX).toFixed(1)}px`
      : '',
  ]
    .filter(Boolean)
    .join(';')
  return [
    'span',
    {
      class: `doc-diagram${diagram.canvas ? ' doc-diagram-canvas' : ''}`,
      style: spanStyle,
    },
    ...shapeSpecs,
  ]
}

/** in-column wrapSquare box in a floated carrier (the box's own style is
 *  rewritten by autogrow): the column-edge margin keeps Word's horizontal
 *  position (negative when the box protrudes), the text-side margin is the
 *  anchor's clearance; a paragraph-relative offset below the anchor top starts
 *  the float (and its exclusion) there — above it Word would clamp the box
 *  onto the page, which the flow cannot express */
function sideWrappedBoxSpec(box: TextboxDisplay): DomSpec {
  const px = (v: number) => `${v.toFixed(1)}px`
  const top = px(Math.max(0, (box.offsetYEmu ?? 0) / EMU_PER_PX))
  const edge = px(box.wrapEdgePx ?? 0)
  const gap = px(box.wrapGapPx ?? 12)
  const margin = box.wrapSide === 'right' ? `${top} ${edge} 0 ${gap}` : `${top} ${gap} 0 ${edge}`
  // a float excludes its whole margin box: the lines above the offset box stay full width
  const shape = top === '0.0px' ? '' : `;shape-outside:inset(${top} 0 0 0)`
  return [
    'div',
    { class: 'doc-textbox-side', style: `float:${box.wrapSide};margin:${margin}${shape}` },
    renderTextboxSpec(box),
  ]
}

/** wrapTopAndBottom band bottom (px) for a box at the given (live) height */
export function textboxBandBottom(box: TextboxDisplay, height = box.heightPx): number {
  if (box.bandTopPx !== undefined && height !== undefined) {
    // a quarter-turned box turns about its centre: its visual band is the
    // swapped extent, sticking out (w-h)/2 above and below the layout box
    return box.bandTopPx + height + quarterTurnInsetPx(box.widthPx ?? 0, height, box.rotDeg)
  }
  return box.bandBottomPx ?? 0
}

/** wrapTopAndBottom band top (px), lifted by the quarter-turn overhang */
export function textboxBandTop(box: TextboxDisplay, height = box.heightPx): number {
  return (box.bandTopPx ?? 0) - quarterTurnInsetPx(box.widthPx ?? 0, height ?? 0, box.rotDeg)
}

/** narrowest column gap beside a float that Word still fills with text (px) */
const MIN_WRAP_SLIVER_PX = 36

/**
 * Word never wraps text into a sliver narrower than a word; CSS break-word
 * shatters it one character per line instead. When a freely positioned float
 * (posOffset X) leaves such a sliver, hand it to the float's margin so text
 * flows above/below (Word behavior). The right-side sliver is only known at
 * layout time ((100% - x - w)), hence the clamp step function: it consumes
 * the whole sliver when below the threshold, else keeps the wrap distance.
 */
function wrapSliverGuardCss(
  wrap: string,
  txPx: number,
  wPx: number,
  distRightPx: number | null,
  /** run images carry the guard on the unrotated img box; the block wrapper already shrink-wraps the turned picture */
  quarterTurnInsetPx = 0,
): string | null {
  if (!(wPx > 0)) return null
  if (wrap.endsWith('-right')) {
    // the margin pins the box, which starts the inset before the visual edge
    return txPx > 0 && txPx < MIN_WRAP_SLIVER_PX
      ? `margin-left:${(txPx - quarterTurnInsetPx).toFixed(1)}px`
      : null
  }
  if (!wrap.endsWith('-left')) return null
  const keep = ((distRightPx && distRightPx > 0 ? distRightPx : 12) - quarterTurnInsetPx).toFixed(1)
  const avail = `100% - ${(txPx + wPx).toFixed(1)}px`
  return (
    `margin-right:clamp(${keep}px, (${MIN_WRAP_SLIVER_PX}px - (${avail})) * 999, ` +
    `max(${keep}px, ${avail}))`
  )
}

/** shared DOM spec for protected blocks (renderHTML + node view) */
function protectedDomSpec(node: PmNode): DomSpec {
  const {
    blockType,
    label,
    previewText,
    imageDataUrl,
    docxIndex,
    table,
    fieldDisplay,
    decorative,
    textboxes,
    formulaDisplay,
    chartDisplay,
  } = node.attrs
  const attrs: Record<string, string> = {
    'data-doc-protected': String(blockType),
    'data-idx': docxIndex === null ? '' : String(docxIndex),
    class: `doc-protected doc-protected-${blockType}`,
  }
  // field/TOC paragraphs keep their paragraph style so document CSS
  // (TOC1 spacing etc.) reaches the wrapper like any styled paragraph
  if (node.attrs.styleId) attrs['data-style'] = String(node.attrs.styleId)
  if (node.attrs.invisibleMarker) {
    if (node.attrs.anchorLine) {
      attrs.class += ' doc-protected-anchor-line'
      return ['div', attrs, anchorLineSpec(node.attrs.anchorLine as Record<string, unknown>)]
    }
    attrs.class += ' doc-protected-invisible'
    return ['div', attrs]
  }
  if (decorative) {
    attrs.class += ' doc-protected-rule'
    const { ruleColorHex, ruleThicknessPx, ruleWidthPx } = node.attrs
    // document stroke color/size, not chrome: inline hardcoded values
    let style = ''
    if (ruleColorHex) style += `background:#${String(ruleColorHex)};`
    if (ruleThicknessPx) style += `height:${Number(ruleThicknessPx)}px;`
    if (ruleWidthPx) style += `width:${Number(ruleWidthPx)}px;max-width:100%;`
    const lineAttrs: Record<string, string> = { class: 'doc-rule-line' }
    if (style) lineAttrs.style = style
    return ['div', attrs, ['span', lineAttrs]]
  }
  if (Array.isArray(textboxes) && textboxes.length > 0) {
    attrs.class += ' doc-protected-textboxes'
    // paragraph justification centers inline shapes (WordArt) like Word
    const boxAlign = node.attrs.imageAlign
    if (boxAlign === 'center' || boxAlign === 'right') {
      attrs.class += ` doc-protected-boxes-${boxAlign}`
    }
    const boxes = textboxes as TextboxDisplay[]
    const diagram = node.attrs.diagramDisplay as DiagramDisplay | null
    // every box floats at its own anchor offset (wrapNone / multi-drawing
    // paragraphs): the wrapper leaves the flow like Word instead of stacking.
    // An inline drawing sharing the paragraph keeps its flow line while the
    // floats still leave it — stacking a page-sized floating shape into the
    // flow buried the whole page under it (prod cover sheets).
    const floatingBoxes = boxes.filter((b) => b.floating)
    const anyFloating = floatingBoxes.length > 0 && (!diagram || diagram.floating)
    const allFloating = anyFloating && floatingBoxes.length === boxes.length
    const strayRuns = node.attrs.strayRuns as Run[] | null
    let strayStyle = ''
    const sideBox = boxes.length === 1 && !diagram && boxes[0].wrapSide ? boxes[0] : null
    if (anyFloating) {
      attrs.class += ' doc-protected-floating'
      // behindDoc anchors paint under the body text (Word z-order); mirrors
      // the behind-image z band
      if (boxes.every((b) => b.behind)) attrs.class += ' doc-protected-behind'
      // page-pinned cover art: the wrapper stays un-positioned so the boxes'
      // absolute page coordinates resolve against the page box
      if (boxes.every((b) => b.pagePinned)) attrs.class += ' doc-protected-pagepinned'
      // wrapTopAndBottom band: the wrapper reserves flow height down to the
      // lowest such box bottom so following text resumes below (min-height,
      // not a sum — stray flow content on the same paragraph must not add)
      const band = Math.max(0, ...boxes.map((b) => textboxBandBottom(b)))
      if (band > 0) {
        attrs.style = `min-height:${band}px`
        // consecutive anchor paragraphs share the page in Word: expose the raw
        // band geometry so syncAnchorBands can lay a run out band-exclusively
        attrs['data-band'] = String(Math.round(band))
        // column-spanning wrapSquare band: pagination keeps the box on its
        // anchor's page and lets the band overflow the bottom margin (Word)
        if (boxes.some((b) => b.bandOverflow)) attrs['data-band-keep'] = '1'
        if (boxes.some((b) => b.bandBeside)) attrs['data-band-beside'] = '1'
        attrs['data-bands'] = boxes
          .filter((b) => textboxBandBottom(b) > 0)
          .map(
            (b) =>
              `${Math.max(0, Math.round(textboxBandTop(b)))}:${Math.round(textboxBandBottom(b))}`,
          )
          .join(' ')
      } else {
        const inlineH = Math.max(0, ...boxes.map((b) => b.inlineExtentPx ?? 0))
        if (inlineH > 0) attrs.style = `min-height:${inlineH}px`
      }
      // the wrapper inherits the page's grid-snapped line-height; snapToGrid=0
      // keeps the anchor line natural (textbox paragraphs keep their own struts)
      if (node.attrs.anchorSnapToGrid === false) {
        attrs.style =
          `${attrs.style ? `${attrs.style};` : ''}` +
          'line-height:calc(var(--doc-line-factor,1.2) * 1em * var(--doc-line-mult,1))'
      }
      // stray text keeps the anchor paragraph's flow line in Word, so the
      // wrapper must not collapse to height 0 (next block would overlap it)
      if (strayRuns?.length) attrs.class += ' doc-protected-floating-stray'
      // Word reserves the anchor paragraph's own (empty) line even when its
      // runs carry only anchored drawings — page-pinned ones too (JP flowchart
      // docs anchor paragraph-relative labels below runs of pinned shapes;
      // collapsed pinned lines pulled every later anchor up)
      else attrs.class += ' doc-protected-floating-stray'
    } else if (sideBox) {
      // body text (the anchor paragraph's own runs and the blocks after it)
      // flows beside the floated box like Word's square wrap
      attrs.class += ' doc-protected-wrapside'
    } else {
      const offsetX = node.attrs.imageOffsetXEmu
      const offsetY = node.attrs.imageOffsetYEmu
      if (offsetX != null || offsetY != null) {
        const dx = Number(offsetX ?? 0) / EMU_PER_PX
        const dy = Number(offsetY ?? 0) / EMU_PER_PX
        attrs.style = `transform:translate(${dx}px,${dy}px)`
        // the translate paints the block off its flow slot; pagination measures the slot
        attrs['data-anchor-dy'] = String(dy)
        // the drawing offset moves the boxes only; stray text stays put
        strayStyle = `transform:translate(${-dx}px,${-dy}px)`
      }
    }
    const children: DomSpec[] = sideBox
      ? [sideWrappedBoxSpec(sideBox)]
      : boxes.map(renderTextboxSpec)
    // the anchor paragraph's own text (e.g. a heading sharing its paragraph
    // with a sidebar box) renders as a display-only line before the boxes
    if (strayRuns?.length) {
      const strayAttrs: Record<string, string> = { class: 'doc-textbox-stray' }
      // a numbered anchor paragraph lays its line out like a .doc-li: the w:ind
      // feeds the marker-box geometry instead of plain margins
      const listed = node.attrs.strayList != null
      if (listed) strayAttrs.class += ' doc-li-stray'
      // the anchor paragraph's own w:ind: a large right indent carves the wrap
      // column beside the box; without it the stray line runs under the box
      const ind = node.attrs.strayIndent as StrayIndent | null
      const indCss = ind
        ? [
            ind.leftTwips ? `${listed ? '--li-left' : 'margin-left'}:${ind.leftTwips / 20}pt` : '',
            ind.rightTwips ? `margin-right:${ind.rightTwips / 20}pt` : '',
            ind.firstLineTwips && listed && ind.firstLineTwips < 0
              ? `--li-hang:${-ind.firstLineTwips / 20}pt`
              : ind.firstLineTwips
                ? `text-indent:${ind.firstLineTwips / 20}pt`
                : '',
          ]
            .filter(Boolean)
            .join(';')
        : ''
      const align = node.attrs.strayAlign as string | null
      const alignCss = align === 'center' || align === 'right' ? `text-align:${align}` : ''
      const strayCss = [strayStyle, indCss, alignCss].filter(Boolean).join(';')
      if (strayCss) strayAttrs.style = strayCss
      if (node.attrs.strayStyleId) strayAttrs['data-style'] = String(node.attrs.strayStyleId)
      const stray: DomSpec = ['div', strayAttrs, ...strayRuns.flatMap((run) => runSpanSpecs(run))]
      // a float shortens only the lines laid out after it
      if (sideBox) children.push(stray)
      else children.unshift(stray)
    } else if (
      attrs.class.includes('doc-protected-floating-stray') &&
      allFloating &&
      (fieldDisplay as FieldDisplay | null)?.kind !== 'pageBreak'
    ) {
      // the anchor paragraph's empty line (the break-chip branch below renders
      // its own stray line instead; a static inline drawing is its own line)
      children.unshift(['div', { class: 'doc-anchor-strut' }, ['br']])
    }
    if (diagram?.shapes?.length) children.push(diagramSpecOf(diagram))
    // corner resize handle; multi-box nodes keep per-box autogrow semantics only
    if (boxes.length === 1 && !diagram) {
      children.push(['span', { class: 'box-resize-handle', contenteditable: 'false' }])
    }
    // page-type w:br the anchor paragraph carries next to the boxes: Word keeps
    // the anchor's flow line, so render the break chip as a stray line (nonzero
    // height — a zero-height carrier cannot advance the pagination Y coordinate)
    if ((fieldDisplay as FieldDisplay | null)?.kind === 'pageBreak') {
      attrs.class += ' doc-protected-floating-stray'
      const spec = renderFieldSpec(fieldDisplay as FieldDisplay)
      if (spec) children.push(spec)
    }
    return ['div', attrs, moveHandleSpec(t('editorMoveTextbox')), ...children]
  }
  // empty section-break paragraphs (page-per-section converter output): Word
  // shows nothing here, so render a near-invisible strip (hover reveals it)
  if (label === 'Section break paragraph' && !previewText) {
    attrs.class += ' doc-protected-sectbreak'
    return ['div', attrs, ['span', { class: 'doc-sectbreak-label' }, t('editorSectionBreak')]]
  }
  // TOC field boundary paragraphs (fldChar begin + instruction / lone fldChar
  // end) have no visible result; Word shows nothing there either, so they get
  // the same near-invisible strip (hover/selection reveals the label)
  if (
    !previewText &&
    !fieldDisplay &&
    (label === 'Auto TOC (updates when opened in Word)' || label === 'Field end marker')
  ) {
    attrs.class += ' doc-protected-sectbreak'
    return ['div', attrs, ['span', { class: 'doc-sectbreak-label' }, String(label)]]
  }
  if (blockType === 'image' && imageDataUrl) {
    const {
      imageWidthPx,
      imageHeightPx,
      imageAlign,
      imageWrap,
      imageCrop,
      imageFillRect,
      imageRotDeg,
    } = node.attrs
    // clipboard round-trip payload (r136): parseHTML rebuilds the image from
    // the inner img src plus these display attrs; without it a copied picture
    // pasted back as an attribute-less "protected content" shell
    attrs['data-image-meta'] = imageMetaJson(node.attrs)
    if (imageAlign === 'center' || imageAlign === 'right') {
      attrs['style'] = `text-align:${imageAlign}`
    }
    // a side-wrapped picture no text fits beside floats at its offset with the
    // wrapper reserving its band (a CSS float would drop below the neighbour)
    const banded =
      node.attrs.imageBand === true &&
      /-(?:left|right)$/.test(String(imageWrap ?? '')) &&
      node.attrs.imageOffsetXEmu != null &&
      Number(imageHeightPx) > 0
    if (banded) attrs.class += ' doc-img-float img-wrap-band'
    else if (imageWrap) attrs.class += ` img-wrap-${String(imageWrap)}`
    const imageLeadingText = String(node.attrs.imageLeadingText ?? '')
    const cjkFixedLeadingSpaces =
      !imageWrap &&
      /^[ ]+$/.test(imageLeadingText) &&
      !!node.attrs.imageLeadingFont &&
      isCjkFontName(String(node.attrs.imageLeadingFont))
    // paragraph indents place the picture like its (possibly empty) first line;
    // anchored pictures position from the column instead and ignore them
    if (!imageWrap) {
      const paragraphLayout = [
        node.attrs.imageParagraphIndentLeft
          ? `margin-inline-start:${Number(node.attrs.imageParagraphIndentLeft) / 20}pt`
          : '',
        node.attrs.imageParagraphIndentRight
          ? `margin-inline-end:${Number(node.attrs.imageParagraphIndentRight) / 20}pt`
          : '',
        !cjkFixedLeadingSpaces && node.attrs.imageParagraphIndentFirstLine
          ? `text-indent:${Number(node.attrs.imageParagraphIndentFirstLine) / 20}pt`
          : '',
      ].filter(Boolean)
      if (paragraphLayout.length) {
        attrs.style = `${attrs.style ? `${attrs.style};` : ''}${paragraphLayout.join(';')}`
      }
    }
    const imageLeadingStyle = [
      node.attrs.imageLeadingFont
        ? `font-family:${cssFontFamily(String(node.attrs.imageLeadingFont))}`
        : '',
    ]
      .filter(Boolean)
      .join(';')
    const imageLeadingSpecs: DomSpec[] =
      imageLeadingText && !cjkFixedLeadingSpaces
        ? [
            [
              'span',
              {
                class: 'doc-image-leading-space',
                ...(imageLeadingStyle ? { style: imageLeadingStyle } : {}),
              },
              imageLeadingText,
            ],
          ]
        : []
    // no-wrap / behind-text anchors leave the flow like floating textboxes:
    // zero-height wrapper (doc-img-float), absolutely positioned inner wrap
    // (Word overlays them on the text instead of reserving a line)
    const explicitSpaceWidthPx = Number(node.attrs.imageLeadingExplicitSpaceWidthPx ?? 0)
    const implicitSpaceCount = Number(
      node.attrs.imageLeadingImplicitSpaceCount ??
        (explicitSpaceWidthPx > 0 ? 0 : imageLeadingText.length),
    )
    let imgWrapTransform = cjkFixedLeadingSpaces
      ? `margin-left:calc(${Number(node.attrs.imageParagraphIndentFirstLine ?? 0) / 15 + explicitSpaceWidthPx}px + ${implicitSpaceCount * 0.5}em)`
      : ''
    let imgFloatPos = ''
    // a quarter-turned picture keeps its extent box (turned about its centre)
    // while the flow reserves the swapped bounding box, so in-flow offsets
    // shift by half the side difference
    const qt = quarterTurnInsetPx(Number(imageWidthPx), Number(imageHeightPx), imageRotDeg)
    // the vertical posOffset margin must not be clobbered by a wrap-distance
    // margin-top below (distT is clearance, the offset is position — position wins)
    let hasOffsetTopMargin = false
    if (imageWrap === 'front' || imageWrap === 'behind') {
      attrs.class += ' doc-img-float'
      // z-index bands keep behind-text pictures under the body text and
      // front pictures over it, while imageZOrder ranks overlapping anchors
      // within each band (Word bring-forward / send-back).
      const z = node.attrs.imageZOrder != null ? Number(node.attrs.imageZOrder) : 0
      // behind band: negative (below text). front band: >=1 (above text; the
      // text layer sits at auto/0). Parse compresses wild relativeHeight
      // values to compact ranks, but the floor still guards the band: the
      // wrap semantics always win over the rank. No ceiling — the editor
      // content root is isolated (styles.css), so ranks can never climb
      // above editor chrome outside it (table handles, menus).
      const zi = imageWrap === 'behind' ? Math.min(-1, -1000 + z) : Math.max(1, 2 + z)
      const imgZIndexCss = `;z-index:${Math.round(zi)}`
      const posH = node.attrs.imagePosH
      const tx =
        node.attrs.imageOffsetXEmu != null ? Number(node.attrs.imageOffsetXEmu) / EMU_PER_PX : 0
      const ty =
        node.attrs.imageOffsetYEmu != null ? Number(node.attrs.imageOffsetYEmu) / EMU_PER_PX : 0
      if (posH === 'center') {
        imgFloatPos = `left:50%;top:${ty.toFixed(1)}px${imgZIndexCss}`
        imgWrapTransform = 'transform:translateX(-50%)'
      } else if (posH === 'right') {
        imgFloatPos = `right:0;top:${ty.toFixed(1)}px${imgZIndexCss}`
      } else {
        imgFloatPos = `left:${tx.toFixed(1)}px;top:${ty.toFixed(1)}px${imgZIndexCss}`
      }
    } else if (banded) {
      const tx = Number(node.attrs.imageOffsetXEmu) / EMU_PER_PX
      const ty = Number(node.attrs.imageOffsetYEmu ?? 0) / EMU_PER_PX
      const bottom =
        ty + Number(imageHeightPx) + Number(node.attrs.imageWrapDistBottomEmu ?? 0) / EMU_PER_PX
      imgFloatPos = `left:${tx.toFixed(1)}px;top:${ty.toFixed(1)}px`
      attrs.style = `${attrs.style ? `${attrs.style};` : ''}min-height:${Math.max(0, bottom).toFixed(1)}px`
    } else if (imageWrap) {
      // In-flow wraps (square/tight/through/topBottom) honor a numeric
      // posOffset so a dragged picture stays where it was dropped (Word
      // WYSIWYG): X measures from the column start, Y from the anchor
      // paragraph top — matching the saved relativeFrom column/paragraph.
      const wrapperCss: string[] = []
      if (node.attrs.imageOffsetYEmu != null || qt) {
        // the offset places the extent box; the quarter turn about its centre
        // lifts the visual (and the text exclusion) by the inset
        const ty = Number(node.attrs.imageOffsetYEmu ?? 0) / EMU_PER_PX - qt
        // a negative offset lifts via the inner wrap so the flow band keeps
        // its height (Word: logo above its anchor line must not push)
        if (ty < 0) imgWrapTransform = `margin-top:${ty.toFixed(1)}px`
        else if (ty > 0) {
          wrapperCss.push(`margin-top:${ty.toFixed(1)}px`)
          if (imageWrap !== 'topBottom') {
            wrapperCss.push(`shape-outside:inset(${ty.toFixed(1)}px 0 0 0)`)
          }
          hasOffsetTopMargin = true
        }
      }
      if (node.attrs.imageOffsetXEmu != null) {
        const tx = Number(node.attrs.imageOffsetXEmu) / EMU_PER_PX
        const w = Number(imageWidthPx ?? 0)
        if (imageWrap === 'topBottom') {
          // an explicit X replaces the centered slot
          wrapperCss.push('text-align:left')
          imgWrapTransform =
            (imgWrapTransform ? `${imgWrapTransform};` : '') +
            `margin-left:${(tx + qt).toFixed(1)}px`
        } else if (String(imageWrap).endsWith('-right') && w > 0) {
          // right floats position from the right edge: colW − x − width; the
          // 60% float clamp would shift the math, and Word never shrinks a
          // freely positioned picture
          wrapperCss.push(
            `margin-right:calc(100% - ${(tx + w - qt).toFixed(1)}px)`,
            'max-width:none',
          )
        } else {
          wrapperCss.push(`margin-left:${(tx + qt).toFixed(1)}px`, 'max-width:none')
        }
      }
      if (wrapperCss.length) {
        attrs['style'] = `${attrs['style'] ? `${attrs['style']};` : ''}${wrapperCss.join(';')}`
      }
    }
    // wrap distances apply to in-flow wraps only: wrapNone (front/behind)
    // anchors ignore them in Word, and margins on the zero-height wrapper
    // would displace the following flow content
    if (imageWrap && imageWrap !== 'front' && imageWrap !== 'behind' && !banded) {
      const distancePx = (attr: string): number | null =>
        node.attrs[attr] != null ? Number(node.attrs[attr]) / EMU_PER_PX : null
      const top = distancePx('imageWrapDistTopEmu')
      const bottom = distancePx('imageWrapDistBottomEmu')
      const left = distancePx('imageWrapDistLeftEmu')
      const right = distancePx('imageWrapDistRightEmu')
      const distances = [
        !hasOffsetTopMargin && top != null ? `margin-top:${top.toFixed(1)}px` : '',
        bottom != null ? `margin-bottom:${bottom.toFixed(1)}px` : '',
        String(imageWrap).endsWith('-right') && left != null
          ? `margin-left:${left.toFixed(1)}px`
          : '',
        String(imageWrap).endsWith('-left') && right != null
          ? `margin-right:${right.toFixed(1)}px`
          : '',
      ].filter(Boolean)
      if (distances.length) {
        attrs.style = `${attrs.style ? `${attrs.style};` : ''}${distances.join(';')}`
      }
      if (node.attrs.imageOffsetXEmu != null && imageWrap !== 'topBottom') {
        const guard = wrapSliverGuardCss(
          String(imageWrap),
          Number(node.attrs.imageOffsetXEmu) / EMU_PER_PX + qt,
          qt ? Number(imageHeightPx ?? 0) : Number(imageWidthPx ?? 0),
          node.attrs.imageWrapDistRightEmu != null
            ? Number(node.attrs.imageWrapDistRightEmu) / EMU_PER_PX
            : null,
        )
        if (guard) attrs.style = `${attrs.style ? `${attrs.style};` : ''}${guard}`
      }
    }
    const imgAttrs: Record<string, string> = {
      src: String(imageDataUrl),
      class: 'doc-protected-img',
    }
    if (imageWidthPx) {
      imgAttrs['style'] =
        `width:${Number(imageWidthPx)}px;` +
        (imageHeightPx ? `height:${Number(imageHeightPx)}px` : 'height:auto')
    }
    // picture outline (document data, not chrome): border adds outside the
    // extent, approximating Word's centered stroke
    const ib = node.attrs.imageBorder as { color: string; widthPt: number } | null
    const borderCss = ib
      ? `border:${((ib.widthPt * 96) / 72).toFixed(1)}px solid #${String(ib.color).replace(/[^0-9A-Fa-f]/g, '')}`
      : ''
    const xf = pictureTransformFns(imageRotDeg, node.attrs.imageFlipH, node.attrs.imageFlipV)
    // in-flow pictures reserve the quarter-turned bounding box (Word turns the
    // extent box about its centre); overlays keep their zero footprint
    const qtMargin = imgFloatPos ? '' : quarterTurnMarginCss(qt)
    // a:srcRect source crop / a:fillRect fill placement: an overflow-hidden
    // window at the declared extent over a scaled and offset image
    const rect = (imageCrop ?? imageFillRect) as {
      l: number
      t: number
      r: number
      b: number
    } | null
    if (rect && imageWidthPx && imageHeightPx) {
      const W = Number(imageWidthPx)
      const H = Number(imageHeightPx)
      const span = (a: number, b: number) => Math.max(0.01, 1 - a - b)
      let sw: number, sh: number, dx: number, dy: number
      if (imageCrop) {
        // crop: the window shows the (1-l-r)×(1-t-b) slice of the source
        sw = W / span(rect.l, rect.r)
        sh = H / span(rect.t, rect.b)
        dx = -rect.l * sw
        dy = -rect.t * sh
      } else {
        // fillRect: the image occupies the inset (negative = bleeding) sub-rect
        sw = W * (1 - rect.l - rect.r)
        sh = H * (1 - rect.t - rect.b)
        dx = rect.l * W
        dy = rect.t * H
      }
      imgAttrs['style'] =
        `position:absolute;left:${dx.toFixed(1)}px;top:${dy.toFixed(1)}px;` +
        `width:${sw.toFixed(1)}px;height:${sh.toFixed(1)}px;max-width:none`
      // the crop wrap carries the positioning margins itself, so the footprint
      // inset folds into them instead of replacing them
      const cropWrapMargins = foldQuarterTurnMargins(
        imgFloatPos ? 0 : qt,
        imgWrapTransform.startsWith('transform:') ? '' : imgWrapTransform,
      )
      // rot/flip turn the whole crop window, not the source inside it
      const wrapXf = [
        ...(imgWrapTransform.startsWith('transform:')
          ? [imgWrapTransform.slice('transform:'.length)]
          : []),
        ...xf,
      ]
      return [
        'div',
        attrs,
        moveHandleSpec(t('editorMoveImage')),
        ...(imageWrap ? [imageAnchorMarkerSpec()] : []),
        ...imageLeadingSpecs,
        [
          'span',
          {
            class: 'doc-img-wrap doc-img-crop',
            style: `position:${imgFloatPos ? `absolute;${imgFloatPos}` : 'relative'};display:inline-block;width:${W}px;height:${H}px${wrapXf.length ? `;transform:${wrapXf.join(' ')}` : ''}${cropWrapMargins ? `;${cropWrapMargins}` : ''}${borderCss ? `;${borderCss}` : ''}`,
          },
          [
            'span',
            {
              class: 'doc-img-crop-viewport',
              style: 'position:absolute;inset:0;overflow:hidden',
            },
            ['img', imgAttrs],
          ],
          ...imageSelectionControlsSpec(),
        ],
      ]
    }
    if (xf.length) {
      imgAttrs['style'] =
        `${imgAttrs['style'] ? `${imgAttrs['style']};` : ''}transform:${xf.join(' ')}`
    }
    if (borderCss) {
      imgAttrs['style'] = `${imgAttrs['style'] ? `${imgAttrs['style']};` : ''}${borderCss}`
    }
    if (qtMargin) {
      imgAttrs['style'] = `${imgAttrs['style'] ? `${imgAttrs['style']};` : ''}${qtMargin}`
    }
    return [
      'div',
      attrs,
      moveHandleSpec(t('editorMoveImage')),
      ...(imageWrap ? [imageAnchorMarkerSpec()] : []),
      ...imageLeadingSpecs,
      [
        'span',
        {
          class: 'doc-img-wrap',
          ...(imgFloatPos || imgWrapTransform
            ? {
                style: `${imgFloatPos ? `position:absolute;${imgFloatPos};` : ''}display:inline-block${imgWrapTransform ? `;${imgWrapTransform}` : ''}`,
              }
            : {}),
        },
        ['img', imgAttrs],
        ...imageSelectionControlsSpec(),
      ],
    ]
  }
  if (blockType === 'table' && table && (table as TableModel).rows?.length) {
    return [
      'div',
      attrs,
      moveHandleSpec(t('editorMoveTable')),
      renderTableSpec(table as TableModel),
    ]
  }
  if (fieldDisplay) {
    const spec = renderFieldSpec(fieldDisplay as FieldDisplay)
    if (spec) {
      attrs.class += ' doc-protected-field'
      const field = fieldDisplay as FieldDisplay
      if (field.deleted && field.markDeleted) attrs.class += ' doc-para-del-collapse'
      if (field.kind === 'text') attrs.class += ' doc-protected-field-text'
      return ['div', attrs, spec]
    }
  }
  if ((formulaDisplay as FormulaDisplay | null)?.tokens?.length) {
    attrs.class += ' doc-protected-formula'
    if ((formulaDisplay as FormulaDisplay).mathml) attrs.class += ' doc-protected-formula-display'
    return [
      'div',
      attrs,
      moveHandleSpec(t('editorMoveEquation')),
      renderFormulaSpec(formulaDisplay as FormulaDisplay),
    ]
  }
  if ((chartDisplay as ChartDisplay | null)?.series?.length) {
    attrs.class += ' doc-protected-chart'
    return [
      'div',
      attrs,
      moveHandleSpec(t('editorMoveChart')),
      renderChartSpec(chartDisplay as ChartDisplay),
      ['span', { class: 'box-resize-handle', contenteditable: 'false' }],
    ]
  }
  // SmartArt with a precomputed drawing part: absolutely positioned shapes
  // (picture fills, solid fills, centered texts) at Word's resolved geometry.
  // All colors are document data (theme-resolved at parse), hence inline.
  const diagram = node.attrs.diagramDisplay as DiagramDisplay | null
  if (diagram?.shapes?.length) {
    attrs.class += ' doc-protected-diagram'
    if (diagram.floating) {
      attrs.class += ' doc-protected-floating'
    } else if (diagram.offsetXEmu != null || diagram.offsetYEmu != null) {
      attrs.style =
        `transform:translate(${Number(diagram.offsetXEmu ?? 0) / EMU_PER_PX}px,` +
        `${Number(diagram.offsetYEmu ?? 0) / EMU_PER_PX}px)`
    }
    return ['div', attrs, moveHandleSpec(t('editorMoveImage')), diagramSpecOf(diagram)]
  }
  // Broken picture (missing rel/media): empty frame at the declared extent
  // with centered alt text. Frame/text colors stand in for document content
  // (must look identical in both themes), hence hardcoded inline.
  if (node.attrs.brokenImage) {
    attrs.class += ' doc-protected-broken-img'
    const { imageWidthPx, imageHeightPx, imageAlign } = node.attrs
    if (imageAlign === 'center' || imageAlign === 'right') {
      attrs.style = `text-align:${imageAlign}`
    }
    let style = 'border:1px solid #999;color:#888;'
    if (imageWidthPx) style += `width:${Number(imageWidthPx)}px;`
    if (imageHeightPx) style += `height:${Number(imageHeightPx)}px;`
    return [
      'div',
      attrs,
      ['span', { class: 'doc-broken-img-frame', style }, String(previewText || label || '')],
    ]
  }
  // OLE embed with a packaged preview picture: Word draws just the preview at
  // its declared size (Icon previews carry their own caption inside the
  // metafile) — the friendly type name moves to a hover tooltip
  const oleCaption = label === 'Embedded object' ? oleTypeLabel(node.attrs.oleProgId) : null
  if (imageDataUrl && blockType === 'passthrough') {
    attrs.class += ' doc-protected-ole'
    attrs.title = oleCaption ?? String(label)
    const { imageWidthPx, imageHeightPx, imageAlign } = node.attrs
    if (imageAlign === 'center' || imageAlign === 'right') {
      attrs.style = `text-align:${imageAlign}`
    }
    let imgStyle = ''
    if (imageWidthPx) imgStyle += `width:${Number(imageWidthPx)}px;`
    if (imageHeightPx) imgStyle += `height:${Number(imageHeightPx)}px;`
    return [
      'div',
      attrs,
      [
        'span',
        { class: 'doc-ole-wrap' },
        ['img', { src: String(imageDataUrl), class: 'doc-ole-img', style: imgStyle }],
      ],
    ]
  }
  const children: unknown[] = [
    [
      'span',
      { class: 'doc-protected-label' },
      oleCaption ?? String(label || t('editorProtectedContent')),
    ],
  ]
  if (previewText) children.push(['span', { class: 'doc-protected-preview' }, String(previewText)])
  return ['div', attrs, ...children]
}

/** o:OLEObject ProgID → localized friendly kind */
function oleTypeLabel(progId: unknown): string {
  const id = typeof progId === 'string' ? progId : ''
  if (id.startsWith('Excel.')) return t('editorOleExcel')
  if (id.startsWith('Word.')) return t('editorOleWord')
  if (id.startsWith('PowerPoint.')) return t('editorOlePpt')
  if (id.startsWith('AcroExch')) return t('editorOlePdf')
  return id ? `${t('editorOleGeneric')} (${id.split('.')[0]})` : t('editorOleGeneric')
}

function moveHandleSpec(label: string): DomSpec {
  return [
    'span',
    {
      class: 'doc-move-handle',
      title: label,
      'aria-label': label,
      contenteditable: 'false',
    },
    '↕',
  ]
}

type ImageResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'

const IMAGE_RESIZE_HANDLES: readonly ImageResizeHandle[] = [
  'nw',
  'n',
  'ne',
  'e',
  'se',
  's',
  'sw',
  'w',
]

function imageSelectionControlsSpec(): DomSpec[] {
  return IMAGE_RESIZE_HANDLES.map((direction): DomSpec => [
    'span',
    {
      class: `img-resize-handle img-resize-handle-${direction}`,
      'data-resize-handle': direction,
      contenteditable: 'false',
      draggable: 'false',
      'aria-hidden': 'true',
    },
  ])
}

function imageAnchorMarkerSpec(): DomSpec {
  return [
    'span',
    {
      class: 'doc-image-anchor-marker',
      contenteditable: 'false',
      'aria-hidden': 'true',
    },
    '⚓',
  ]
}

export function buildProtectedDom(node: PmNode): HTMLElement {
  const { dom } = DOMSerializer.renderSpec(document, protectedDomSpec(node) as never)
  const el = dom as HTMLElement
  const mathml = (node.attrs.formulaDisplay as FormulaDisplay | null)?.mathml
  const mathHost = el.querySelector?.('.doc-formula-math')
  if (mathml && mathHost) mathHost.innerHTML = mathml
  return el
}

export interface ProtectedContentEditor {
  setEditable(editable: boolean): void
  commit(): void
}

const EDITABLE_PROTECTED_SELECTOR =
  'td, .doc-textbox, .doc-toc-title, .doc-toc-page, .doc-field-text, .doc-formula-token, ' +
  '.doc-formula-math, .doc-chart-title, .doc-chart-cell'

/** Object mode (single click/drag) and text mode (double click). */
function wireProtectedInteractionMode(
  dom: HTMLElement,
  getPos: () => number | undefined,
  view: EditorView,
  contentEditor: ProtectedContentEditor | null,
): (() => void) | null {
  const handle = dom.querySelector('.doc-move-handle') as HTMLElement | null
  if (!handle && !contentEditor) return null

  const selectObject = () => {
    const pos = getPos()
    if (typeof pos !== 'number') return
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)))
  }
  const setEditing = (editing: boolean) => {
    if (!contentEditor) editing = false
    if (editing === dom.classList.contains('doc-content-editing')) {
      contentEditor?.setEditable(editing)
      dom.draggable = !!handle && !editing
      return
    }
    if (!editing) contentEditor?.commit()
    contentEditor?.setEditable(editing)
    dom.classList.toggle('doc-content-editing', editing)
    dom.draggable = !!handle && !editing
  }
  const onMouseDown = (event: MouseEvent) => {
    if (event.button !== 0) return
    const target = event.target as HTMLElement | null
    if (!target) return
    if (target.closest('.doc-move-handle')) {
      setEditing(false)
      selectObject()
      return
    }
    const content = target.closest(EDITABLE_PROTECTED_SELECTOR)
    if (content && dom.contains(content)) {
      if (!dom.classList.contains('doc-content-editing')) selectObject()
      dom.draggable = !!handle && !dom.classList.contains('doc-content-editing')
      return
    }
    if (dom.classList.contains('doc-content-editing')) setEditing(false)
    selectObject()
  }
  const onDoubleClick = (event: MouseEvent) => {
    if (!contentEditor) return
    const target = event.target as HTMLElement | null
    const content = target?.closest(EDITABLE_PROTECTED_SELECTOR)
    if (!content || !dom.contains(content)) return
    setEditing(true)
  }
  const onDocumentMouseDown = (event: MouseEvent) => {
    const target = event.target
    if (
      target instanceof HTMLElement &&
      !dom.contains(target) &&
      dom.classList.contains('doc-content-editing')
    ) {
      setEditing(false)
    }
  }
  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key !== 'Escape' || !dom.classList.contains('doc-content-editing')) return
    setEditing(false)
    selectObject()
    dom.focus()
  }

  setEditing(false)
  dom.addEventListener('mousedown', onMouseDown, true)
  dom.addEventListener('dblclick', onDoubleClick, true)
  document.addEventListener('mousedown', onDocumentMouseDown, true)
  dom.addEventListener('keydown', onKeyDown, true)
  return () => {
    dom.removeEventListener('mousedown', onMouseDown, true)
    dom.removeEventListener('dblclick', onDoubleClick, true)
    document.removeEventListener('mousedown', onDocumentMouseDown, true)
    dom.removeEventListener('keydown', onKeyDown, true)
  }
}

/** split a contenteditable cell back into paragraph strings */
function tdParas(td: HTMLElement): string[] {
  // one rendered block per paragraph (below the exact-height / vertical-text
  // wrapper): read them structurally so empty paragraphs survive (innerText
  // drops trailing empty blocks) and NBSP stays document text
  const wrapper = td.children.length === 1 ? td.children[0] : null
  const host = wrapper?.matches('.cell-clip, .cell-vert') ? wrapper : td
  const blocks = Array.from(host.childNodes)
  if (blocks.length > 0 && blocks.every((n) => n instanceof HTMLElement && n.tagName === 'DIV')) {
    return blocks.map((block) => {
      const el = block as HTMLElement
      return (el.innerText ?? el.textContent ?? '').replace(/\n$/, '')
    })
  }
  const text = (td.innerText ?? td.textContent ?? '').replace(/\n+$/, '')
  const paras = text.split('\n')
  if (paras.length === 1 && paras[0].trim() === '') return ['']
  return paras
}

/**
 * In-place cell editing: cells become contenteditable islands and
 * their text is committed back into the node's TableModel when focus leaves
 * the table (or when App broadcasts 'ai-docs-commit-tables' before saving).
 */
function wireTableEditing(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
  view: EditorView,
): (ProtectedContentEditor & { cleanup(): void }) | null {
  const node = getNode()
  if (node.attrs.blockType !== 'table' || !node.attrs.table) return null
  const setEditable = (editable: boolean) => {
    for (const td of Array.from(dom.querySelectorAll('td'))) {
      td.setAttribute('contenteditable', editable ? 'true' : 'false')
    }
  }
  const dirty = new Set<HTMLElement>()
  dom.addEventListener('input', (e) => {
    const td = (e.target as HTMLElement | null)?.closest?.('td')
    if (td) dirty.add(td)
  })

  const commit = () => {
    const current = getNode()
    const model = current.attrs.table as TableModel
    const domTds = Array.from(dom.querySelectorAll('td'))
    let k = 0
    let changed = false
    const rows = model.rows.map((row) =>
      row.map((cell) => {
        if (cell.vMerge === 'continue') return cell
        const td = domTds[k++]
        if (!td || !dirty.has(td as HTMLElement)) return cell
        const paras = tdParas(td as HTMLElement)
        if (paras.join('\n') === cell.paras.join('\n')) return cell
        changed = true
        return { ...cell, paras }
      }),
    )
    dirty.clear()
    if (!changed) return
    const pos = getPos()
    if (typeof pos !== 'number') return
    view.dispatch(
      view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, table: { ...model, rows } }),
    )
  }

  dom.addEventListener('focusout', (e) => {
    const next = (e as FocusEvent).relatedTarget as HTMLElement | null
    if (next && dom.contains(next)) return // moving between cells: not yet
    commit()
  })
  window.addEventListener('ai-docs-commit-tables', commit)
  return {
    setEditable,
    commit,
    cleanup: () => window.removeEventListener('ai-docs-commit-tables', commit),
  }
}

/** the renderer pads empty protected content with one NBSP; every other NBSP is document text */
export function protectedText(element: HTMLElement): string {
  const text = element.innerText ?? element.textContent ?? ''
  return text === '\u00a0' ? '' : text
}

export function preventProtectedLineBreak(event: KeyboardEvent) {
  if (event.key === 'Enter') event.preventDefault()
}

/** Edit cached visible field results without exposing field instructions. */
function wireFieldEditing(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
  view: EditorView,
): (ProtectedContentEditor & { cleanup(): void }) | null {
  const field = getNode().attrs.fieldDisplay as FieldDisplay | null
  if (!field || field.kind === 'pageBreak') return null
  const targets = Array.from(
    dom.querySelectorAll<HTMLElement>('.doc-toc-title, .doc-toc-page, .doc-field-text'),
  )
  if (targets.length === 0) return null

  const setEditable = (editable: boolean) => {
    for (const target of targets)
      target.setAttribute('contenteditable', editable ? 'true' : 'false')
  }
  // only user input may rewrite the cached result: re-reading untouched DOM
  // text is lossy (innerText normalisation) and would dirty every field on save
  let dirty = false
  const markDirty = () => {
    dirty = true
  }
  const commit = () => {
    if (!dirty) return
    dirty = false
    const current = getNode()
    const currentField = current.attrs.fieldDisplay as FieldDisplay | null
    if (!currentField) return
    const next: FieldDisplay = { ...currentField }
    if (currentField.kind === 'tocLine') {
      const left = dom.querySelector<HTMLElement>('.doc-toc-title')
      const right = dom.querySelector<HTMLElement>('.doc-toc-page')
      if (left) next.left = protectedText(left)
      if (right) next.right = protectedText(right)
    } else {
      const text = dom.querySelector<HTMLElement>('.doc-field-text')
      if (text) next.left = protectedText(text)
      // the formatted runs describe the pre-edit text: an edited result renders
      // as the plain string (a rebuild would otherwise redraw the stale runs)
      if (next.left !== currentField.left) delete next.runs
    }
    if (JSON.stringify(next) === JSON.stringify(currentField)) return
    const pos = getPos()
    if (typeof pos !== 'number') return
    view.dispatch(
      view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, fieldDisplay: next }),
    )
  }
  for (const target of targets) target.addEventListener('keydown', preventProtectedLineBreak)
  dom.addEventListener('input', markDirty)
  window.addEventListener('ai-docs-commit-tables', commit)
  return {
    setEditable,
    commit,
    cleanup: () => {
      for (const target of targets) target.removeEventListener('keydown', preventProtectedLineBreak)
      dom.removeEventListener('input', markDirty)
      window.removeEventListener('ai-docs-commit-tables', commit)
    },
  }
}

/** hover button on display formulas whose LaTeX was recovered: full re-edit */
function wireFormulaLatexEdit(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
): (() => void) | null {
  const formula = getNode().attrs.formulaDisplay as FormulaDisplay | null
  if (!formula?.latex) return null
  const host = dom.querySelector('.doc-formula-wrap') ?? dom
  const button = document.createElement('button')
  button.className = 'doc-formula-edit'
  button.title = t('editorEditFormulaLatex')
  button.textContent = t('editorEdit')
  button.setAttribute('contenteditable', 'false')
  button.addEventListener('mousedown', (e) => {
    e.preventDefault()
    e.stopPropagation()
  })
  button.addEventListener('click', (e) => {
    e.preventDefault()
    e.stopPropagation()
    const pos = getPos()
    const latex = (getNode().attrs.formulaDisplay as FormulaDisplay | null)?.latex
    if (typeof pos === 'number' && latex) {
      window.dispatchEvent(
        new CustomEvent('ai-docs-edit-inline-math', {
          detail: { pos, latex, kind: 'block' },
        }),
      )
    }
  })
  host.appendChild(button)
  return () => button.remove()
}

/** Edit OMML leaf tokens while keeping formula structure outside the editor. */
function wireFormulaEditing(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
  view: EditorView,
): (ProtectedContentEditor & { cleanup(): void }) | null {
  const formula = getNode().attrs.formulaDisplay as FormulaDisplay | null
  if (!formula?.tokens.length) return null
  const targets = Array.from(dom.querySelectorAll<HTMLElement>('.doc-formula-token'))
  if (targets.length !== formula.tokens.length) return null

  const setEditable = (editable: boolean) => {
    for (const target of targets)
      target.setAttribute('contenteditable', editable ? 'true' : 'false')
  }
  let dirty = false
  const markDirty = () => {
    dirty = true
  }
  const commit = () => {
    if (!dirty) return
    dirty = false
    const current = getNode()
    const currentFormula = current.attrs.formulaDisplay as FormulaDisplay | null
    if (!currentFormula || currentFormula.tokens.length !== targets.length) return
    const tokens = targets.map(protectedText)
    if (tokens.every((token, i) => token === currentFormula.tokens[i])) return
    const pos = getPos()
    if (typeof pos !== 'number') return
    // re-derive the 2D preview (and, for editor-created formulas, the OOXML
    // to be saved) from the patched OMML source
    const omml = currentFormula.omml ? patchMathTokens(currentFormula.omml, tokens) : undefined
    const nextFormula: FormulaDisplay = omml
      ? { tokens, omml, mathml: ommlToMathML(omml) }
      : { tokens }
    const attrs: Record<string, unknown> = { ...current.attrs, formulaDisplay: nextFormula }
    if (current.attrs.genXml) {
      attrs.genXml = patchMathTokens(String(current.attrs.genXml), tokens)
    }
    view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, attrs))
  }
  for (const target of targets) target.addEventListener('keydown', preventProtectedLineBreak)
  dom.addEventListener('input', markDirty)
  window.addEventListener('ai-docs-commit-tables', commit)
  return {
    setEditable,
    commit,
    cleanup: () => {
      for (const target of targets) target.removeEventListener('keydown', preventProtectedLineBreak)
      dom.removeEventListener('input', markDirty)
      window.removeEventListener('ai-docs-commit-tables', commit)
    },
  }
}

/** node attrs equality ignoring one key (identity per key is enough here) */
function attrsEqualExcept(
  a: Record<string, unknown>,
  b: Record<string, unknown>,
  skip: string,
): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)])
  for (const key of keys) {
    if (key !== skip && !Object.is(a[key], b[key])) return false
  }
  return true
}

type TextboxPara = TextboxDisplay['paras'][number]

/** ProseMirror doc JSON for one textbox's rich content */
function textboxDocJson(box: TextboxDisplay): Record<string, unknown> {
  const paras = box.paras.map((para) => ({
    type: 'docParagraph',
    attrs: {
      styleId: para.styleId ?? null,
      align: para.align ?? null,
      lineSpacing: para.lineSpacing ?? null,
      lineRule: para.lineRule ?? null,
      lineRawTwips: para.lineRawTwips ?? null,
      snapToGrid: para.snapToGrid ?? null,
      indentLeft: para.indentLeft ?? null,
      indentRight: para.indentRight ?? null,
      indentFirstLine: para.indentFirstLine ?? null,
      spaceBefore: para.spaceBefore ?? null,
      spaceAfter: para.spaceAfter ?? null,
      spaceBeforeAuto: para.spaceBeforeAuto ?? null,
      spaceAfterAuto: para.spaceAfterAuto ?? null,
      shadingFill: para.shadingFill ?? null,
      borders: para.borders ?? null,
      listMarker: para.listMarker ?? null,
    },
    content: runsToInline(para.runs),
  }))
  // a shape with no w:txbxContent yet gets Word's centered authoring default,
  // so the live preview matches the jc="center" the inject save path writes
  const fresh = box.txbxIndex === undefined && box.shapeId !== undefined && !box.readOnly
  const emptyPara = fresh
    ? { type: 'docParagraph', attrs: { align: 'center' } }
    : { type: 'docParagraph' }
  return { type: 'doc', content: paras.length > 0 ? paras : [emptyPara] }
}

/** TextboxDisplay paragraphs from a sub-editor's current doc */
function subEditorParas(sub: Editor): TextboxPara[] {
  return ((sub.getJSON().content ?? []) as PmJson[]).map((p) => {
    const para: TextboxPara = { runs: inlineToRuns(p.content ?? []) }
    const attrs = p.attrs ?? {}
    const keys = [
      'styleId',
      'align',
      'lineSpacing',
      'lineRule',
      'lineRawTwips',
      'snapToGrid',
      'indentLeft',
      'indentRight',
      'indentFirstLine',
      'spaceBefore',
      'spaceAfter',
      'spaceBeforeAuto',
      'spaceAfterAuto',
      'shadingFill',
      'borders',
      'listMarker',
    ] as const
    for (const key of keys) {
      const value = attrs[key]
      if (value !== null && value !== undefined) Object.assign(para, { [key]: value })
    }
    return para
  })
}

/**
 * Rich textbox editing: each rendered box hosts a full nested
 * Tiptap editor sharing the main schema's marks, so ribbon formatting (color,
 * bold, size, ...) works inside the box. Content commits back into the node's
 * TextboxDisplay model when focus leaves the block (or before saving); the
 * save path regenerates only the changed paragraphs inside w:txbxContent.
 */
function mountTextboxEditors(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
  view: EditorView,
):
  | (ProtectedContentEditor & {
      cleanup: () => void
      sync: (boxes: TextboxDisplay[] | null) => void
    })
  | null {
  let knownBoxes = getNode().attrs.textboxes as TextboxDisplay[] | null
  if (!knownBoxes || knownBoxes.length === 0) return null

  const editors: Editor[] = []
  const boxEls = Array.from(dom.querySelectorAll('.doc-textbox')) as HTMLElement[]
  const minHeights = knownBoxes.map((box) => box.minHeightPx ?? box.heightPx)
  const measuredHeights = knownBoxes.map((box) => box.heightPx)
  const resizeFrames: Array<number | undefined> = []
  // wrapTopAndBottom band on the wrapper must follow autogrow, or the text
  // below would overlap a box that outgrew its parse-time height
  const refreshBand = () => {
    if (!knownBoxes || !dom.classList.contains('doc-protected-floating')) return
    const bottoms = knownBoxes.map((b, i) => textboxBandBottom(b, measuredHeights[i]))
    const band = Math.max(0, ...bottoms)
    if (band > 0) {
      // the band data drives syncAnchorBands: refresh it too, or the next
      // remeasure would restore the stale parse-time band over the autogrow
      dom.dataset.band = String(Math.round(band))
      dom.dataset.bands = knownBoxes
        .map((b, i) =>
          bottoms[i] > 0
            ? `${Math.max(0, Math.round(textboxBandTop(b, measuredHeights[i])))}:${Math.round(bottoms[i])}`
            : null,
        )
        .filter(Boolean)
        .join(' ')
      // a run-adjusted wrapper keeps its syncAnchorBands layout — writing the
      // own-band value here would break the whole run mid-edit; the next
      // remeasure re-lays the run out from the refreshed band data (page
      // slicing only updates then anyway)
      if (dom.dataset.bandAdj === undefined) dom.style.minHeight = `${band}px`
    }
  }
  const measureBox = (index: number) => {
    const el = boxEls[index]
    const minHeight = minHeights[index]
    if (!el || !minHeight) return
    const borderHeight = el.offsetHeight - el.clientHeight
    el.style.height = 'auto'
    const naturalHeight = Math.ceil(el.scrollHeight + borderHeight)
    const nextHeight = Math.max(minHeight, naturalHeight)
    el.style.height = `${nextHeight}px`
    measuredHeights[index] = nextHeight
    refreshBand()
  }
  const scheduleMeasure = (index: number) => {
    if (resizeFrames[index] !== undefined) cancelAnimationFrame(resizeFrames[index]!)
    resizeFrames[index] = requestAnimationFrame(() => {
      resizeFrames[index] = undefined
      measureBox(index)
    })
  }
  boxEls.forEach((el, i) => {
    const box = knownBoxes![i]
    if (!box) return
    // boxes whose content flattens tables / content controls into display lines
    // keep the static spec: a sub-editor commit would corrupt that structure
    if (box.readOnly) return
    el.replaceChildren() // the static spec children are replaced by the live editor
    const sub: Editor = new Editor({
      element: el,
      extensions: textboxSubExtensions,
      content: textboxDocJson(box),
      editorProps: {
        attributes: { class: 'doc-textbox-editor', spellcheck: 'false' },
      },
      onFocus: () => setActiveSubEditor(sub),
      onTransaction: ({ transaction }) => {
        notifySubEditorState(transaction.docChanged)
        if (transaction.docChanged) scheduleMeasure(i)
      },
    })
    // TipTap initializes the host element and may clear attributes emitted by
    // the static DOM spec, so reapply the shape geometry after mounting.
    el.setAttribute('style', textboxBoxStyle(box))
    el.classList.toggle('doc-textbox-filled', textboxIsFilled(box))
    editors[i] = sub
  })
  if (editors.length === 0) return null

  const setEditable = (editable: boolean) => {
    for (const sub of editors) {
      if (sub && !sub.isDestroyed) sub.setEditable(editable)
    }
  }

  const commit = () => {
    const current = getNode()
    const model = current.attrs.textboxes as TextboxDisplay[] | null
    if (!model) return
    let changed = false
    const next = model.map((box, i) => {
      const sub = editors[i]
      if (!sub || sub.isDestroyed) return box
      const paras = subEditorParas(sub)
      // the sub-editor of a paras:[] ink box always holds one seeded empty
      // paragraph — without text that is still "no content", not an edit, and
      // any measured height is a transient of the emptied text: commit nothing
      if (box.paras.length === 0 && paras.every((p) => p.runs.every((r) => r.text === ''))) {
        return box
      }
      const same =
        paras.length === box.paras.length &&
        paras.every((p, j) => textboxParaSignature(p) === textboxParaSignature(box.paras[j]))
      const measuredHeight = minHeights[i] ? measuredHeights[i] : box.heightPx
      const sameHeight = measuredHeight === box.heightPx
      if (same && sameHeight) return box
      changed = true
      return { ...box, paras, heightPx: measuredHeight }
    })
    if (!changed) return
    const pos = getPos()
    if (typeof pos !== 'number') return
    knownBoxes = next
    view.dispatch(
      view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, textboxes: next }),
    )
  }

  /** external model change (undo of a commit, AI edit): re-feed the sub-editors */
  const sync = (boxes: TextboxDisplay[] | null) => {
    if (!boxes || boxes === knownBoxes) return
    knownBoxes = boxes
    boxes.forEach((box, i) => {
      const sub = editors[i]
      if (sub && !sub.isDestroyed) sub.commands.setContent(textboxDocJson(box))
      const el = boxEls[i]
      if (el) {
        el.setAttribute('style', textboxBoxStyle(box))
        el.classList.toggle('doc-textbox-filled', textboxIsFilled(box))
      }
      // corner resize only rewrites the attrs; without refreshing the minimum
      // the next autofit would shrink the box back below the resized height
      minHeights[i] = box.minHeightPx ?? box.heightPx
      measuredHeights[i] = box.heightPx
    })
    refreshBand()
  }

  dom.addEventListener('focusout', (e) => {
    const next = (e as FocusEvent).relatedTarget as HTMLElement | null
    if (next && dom.contains(next)) return // moving between boxes: not yet
    commit()
  })
  window.addEventListener('ai-docs-commit-tables', commit)
  const cleanup = () => {
    window.removeEventListener('ai-docs-commit-tables', commit)
    for (const frame of resizeFrames) {
      if (frame !== undefined) cancelAnimationFrame(frame)
    }
    for (const sub of editors) {
      if (!sub) continue
      dropActiveSubEditor(sub)
      sub.destroy()
    }
  }
  return { cleanup, sync, setEditable, commit }
}

/** drag the corner handle of a selected image to resize it */
function imageResizePlugin(): Plugin {
  return new Plugin({
    props: {
      handleDOMEvents: {
        mousedown: (view, event) => {
          const target = event.target as HTMLElement
          if (target.classList?.contains('box-resize-handle')) {
            const boxWrapper = target.closest('.doc-protected') as HTMLElement | null
            const isChart = !!boxWrapper?.classList.contains('doc-protected-chart')
            const boxEl = boxWrapper?.querySelector(
              isChart ? '.doc-chart-canvas svg' : '.doc-textbox',
            ) as HTMLElement | null
            if (!boxWrapper || !boxEl) return false
            let boxPos = -1
            view.state.doc.descendants((node, p) => {
              if (boxPos !== -1) return false
              if (node.type.name === 'docProtected' && view.nodeDOM(p) === boxWrapper) boxPos = p
              return boxPos === -1
            })
            if (boxPos === -1) return false
            const attrs = view.state.doc.nodeAt(boxPos)?.attrs
            const boxes = attrs?.textboxes as TextboxDisplay[] | null
            const chart = attrs?.chartDisplay as ChartDisplay | null
            if (!isChart && (!Array.isArray(boxes) || boxes.length !== 1)) return false
            if (isChart && !chart) return false
            event.preventDefault()
            view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, boxPos)))

            const zoomEl = document.querySelector('.doc-zoom') as HTMLElement | null
            const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
            const startRect = boxEl.getBoundingClientRect()
            const startW = isChart
              ? startRect.width / zoom
              : parseFloat(getComputedStyle(boxEl).width) || boxes![0].widthPx || 189
            const startH = isChart
              ? startRect.height / zoom
              : parseFloat(getComputedStyle(boxEl).height) || boxes![0].heightPx || 113
            const startX = event.clientX
            const startY = event.clientY

            // horizontal lines resize in length only (their saved extent is zero-height)
            const lockH = !isChart && isStraightLineKind(boxes![0].prst)
            const minW = isChart ? 120 : 24
            const minH = isChart ? 80 : 8
            // charts never draw wider than the render cap, so don't let the model exceed it
            const maxW = isChart ? CHART_MAX_WIDTH_PX : Infinity
            const sizeAt = (e: MouseEvent) => ({
              w: Math.min(maxW, Math.max(minW, startW + (e.clientX - startX) / zoom)),
              h: lockH ? startH : Math.max(minH, startH + (e.clientY - startY) / zoom),
            })
            const onMove = (e: MouseEvent) => {
              const { w, h } = sizeAt(e)
              boxEl.style.width = `${w}px`
              boxEl.style.height = `${h}px`
            }
            const onUp = (e: MouseEvent) => {
              window.removeEventListener('mousemove', onMove)
              window.removeEventListener('mouseup', onUp)
              const { w, h } = sizeAt(e)
              const node = view.state.doc.nodeAt(boxPos)
              if (!node) return
              if (isChart) {
                const display = node.attrs.chartDisplay as ChartDisplay | null
                if (!display) return
                view.dispatch(
                  view.state.tr.setNodeMarkup(boxPos, undefined, {
                    ...node.attrs,
                    chartDisplay: {
                      ...display,
                      widthPx: Math.round(w),
                      // the handle measures the plot SVG; heightPx spans title row + plot
                      heightPx:
                        Math.round(h) + (display.title !== undefined ? CHART_TITLE_ROW_PX : 0),
                    },
                  }),
                )
                return
              }
              const box = (node.attrs.textboxes as TextboxDisplay[] | null)?.[0]
              if (!box) return
              // straight lines keep their zero-height extent: never give them a heightPx
              const next = lockH
                ? { ...box, widthPx: Math.round(w) }
                : {
                    ...box,
                    widthPx: Math.round(w),
                    heightPx: Math.round(h),
                    minHeightPx: Math.round(h),
                  }
              view.dispatch(
                view.state.tr.setNodeMarkup(boxPos, undefined, {
                  ...node.attrs,
                  textboxes: [next],
                }),
              )
            }
            window.addEventListener('mousemove', onMove)
            window.addEventListener('mouseup', onUp)
            return true
          }
          const handle = target.closest('.img-resize-handle') as HTMLElement | null
          const direction = handle?.dataset.resizeHandle as ImageResizeHandle | undefined
          if (!handle || !direction || !IMAGE_RESIZE_HANDLES.includes(direction)) return false
          const wrapper = handle.closest('.doc-protected') as HTMLElement | null
          const imageBox = handle.closest('.doc-img-wrap') as HTMLElement | null
          const img = wrapper?.querySelector('img.doc-protected-img') as HTMLImageElement | null
          if (!wrapper || !imageBox || !img) return false
          event.preventDefault()
          event.stopPropagation()

          let pos = -1
          view.state.doc.descendants((node, p) => {
            if (pos !== -1) return false
            if (node.type.name === 'docProtected' && view.nodeDOM(p) === wrapper) pos = p
            return pos === -1
          })
          if (pos === -1) return false

          view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)))
          view.focus()

          // CSS `zoom` scales client coordinates; divide it back out
          const zoomEl = document.querySelector('.doc-zoom') as HTMLElement | null
          const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
          // Layout-box measurements: getBoundingClientRect would include the
          // rotation/flip transform, swapping width/height for 90°-rotated images
          const selectedNode = view.state.doc.nodeAt(pos)
          const modelW = Number(selectedNode?.attrs.imageWidthPx)
          const modelH = Number(selectedNode?.attrs.imageHeightPx)
          const startW = modelW > 0 ? modelW : imageBox.offsetWidth || img.offsetWidth
          const startH = modelH > 0 ? modelH : imageBox.offsetHeight || img.offsetHeight
          if (!(startW > 0) || !(startH > 0)) return false
          const priorBoxStyle = imageBox.getAttribute('style')
          const priorImgStyle = img.getAttribute('style')
          const cropped = imageBox.classList.contains('doc-img-crop')
          const startImgW = parseFloat(img.style.width) || img.offsetWidth || startW
          const startImgH = parseFloat(img.style.height) || img.offsetHeight || startH
          const startImgLeft = parseFloat(img.style.left) || 0
          const startImgTop = parseFloat(img.style.top) || 0
          const pxLeft = /^-?\d+(?:\.\d+)?px$/.test(imageBox.style.left)
            ? parseFloat(imageBox.style.left)
            : null
          const pxTop = /^-?\d+(?:\.\d+)?px$/.test(imageBox.style.top)
            ? parseFloat(imageBox.style.top)
            : null
          const shiftsInFlow = imageBox.style.position !== 'absolute'
          const startX = event.clientX
          const startY = event.clientY
          const west = direction.includes('w')
          const east = direction.includes('e')
          const north = direction.includes('n')
          const south = direction.includes('s')

          const geometryAt = (e: MouseEvent) => {
            const dx = (e.clientX - startX) / zoom
            const dy = (e.clientY - startY) / zoom
            let w = startW
            let h = startH
            if ((west || east) && (north || south)) {
              // Project the pointer onto the aspect-ratio diagonal. Corner
              // handles keep picture proportions, matching Word's default.
              const sx = west ? -1 : 1
              const sy = north ? -1 : 1
              const delta =
                (dx * sx * startW + dy * sy * startH) /
                Math.max(1, startW * startW + startH * startH)
              const minScale = Math.max(24 / startW, 24 / startH)
              const scale = Math.max(minScale, 1 + delta)
              w = startW * scale
              h = startH * scale
            } else if (west || east) {
              w = Math.max(24, startW + (west ? -dx : dx))
            } else if (north || south) {
              h = Math.max(24, startH + (north ? -dy : dy))
            }
            return {
              w,
              h,
              shiftX: west ? startW - w : 0,
              shiftY: north ? startH - h : 0,
            }
          }

          const restorePreview = () => {
            if (priorBoxStyle === null) imageBox.removeAttribute('style')
            else imageBox.setAttribute('style', priorBoxStyle)
            if (priorImgStyle === null) img.removeAttribute('style')
            else img.setAttribute('style', priorImgStyle)
          }
          const onMove = (e: MouseEvent) => {
            const { w, h, shiftX, shiftY } = geometryAt(e)
            imageBox.style.width = `${w}px`
            imageBox.style.height = `${h}px`
            if (pxLeft !== null) imageBox.style.left = `${pxLeft + shiftX}px`
            else if (west && shiftsInFlow) imageBox.style.left = `${shiftX}px`
            if (pxTop !== null) imageBox.style.top = `${pxTop + shiftY}px`
            else if (north && shiftsInFlow) imageBox.style.top = `${shiftY}px`
            if (cropped) {
              const sx = w / startW
              const sy = h / startH
              img.style.left = `${startImgLeft * sx}px`
              img.style.top = `${startImgTop * sy}px`
              img.style.width = `${startImgW * sx}px`
              img.style.height = `${startImgH * sy}px`
            } else {
              img.style.width = `${w}px`
              img.style.height = `${h}px`
            }
          }
          const onUp = (e: MouseEvent) => {
            window.removeEventListener('mousemove', onMove)
            window.removeEventListener('mouseup', onUp)
            // A plain click on the handle must not rewrite the stored size
            if (Math.abs(e.clientX - startX) < 2 && Math.abs(e.clientY - startY) < 2) {
              restorePreview()
              return
            }
            const node = view.state.doc.nodeAt(pos)
            if (!node) {
              restorePreview()
              return
            }
            const geometry = geometryAt(e)
            const w = Math.round(geometry.w)
            const h = Math.round(geometry.h)
            const attrs: Record<string, unknown> = {
              ...node.attrs,
              imageWidthPx: w,
              imageHeightPx: h,
            }
            if (west && node.attrs.imageOffsetXEmu != null) {
              attrs.imageOffsetXEmu =
                Number(node.attrs.imageOffsetXEmu) + Math.round((startW - w) * EMU_PER_PX)
            }
            if (north && node.attrs.imageOffsetYEmu != null) {
              attrs.imageOffsetYEmu =
                Number(node.attrs.imageOffsetYEmu) + Math.round((startH - h) * EMU_PER_PX)
            }
            view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, attrs))
          }
          window.addEventListener('mousemove', onMove)
          window.addEventListener('mouseup', onUp)
          return true
        },
      },
    },
  })
}

const EMU_PER_PX = 9525

/**
 * Drag floating images and textbox shapes (wp:anchor) to update posOffset.
 * Only handles images with numeric posOffset (imageOffsetXEmu/YEmu set).
 * Inline images (no imageWrap) are auto-converted to anchor on drag start
 * with square wrap and the initial offset derived from the drag delta.
 */
/** Is the point on an actual text glyph (not just inside a text block's box)? */
function pointOnTextGlyph(x: number, y: number): boolean {
  const range = document.caretRangeFromPoint(x, y)
  const tn = range?.startContainer
  // `Node` in this module is TipTap's node class, so use the DOM Text check
  if (!range || !(tn instanceof Text)) return false
  const probe = document.createRange()
  probe.setStart(tn, Math.max(0, range.startOffset - 1))
  probe.setEnd(tn, Math.min(tn.length, range.startOffset + 1))
  for (const r of probe.getClientRects()) {
    if (x >= r.left && x <= r.right && y >= r.top && y <= r.bottom) return true
  }
  return false
}

/**
 * Word parity: a press on empty text area falls through to the floating
 * picture painted below. Behind-text images are covered by the text layer's
 * hit box, so plain DOM hit-testing can never reach them; glyphs still win
 * (clicking on a character places the caret, like Word).
 */
export function findFloatImageAt(x: number, y: number): HTMLElement | null {
  if (pointOnTextGlyph(x, y)) return null
  for (const el of document.elementsFromPoint(x, y)) {
    if (!(el instanceof HTMLElement)) continue
    const wrap = el.classList.contains('doc-img-wrap')
      ? el
      : (el.closest('.doc-img-wrap') as HTMLElement | null)
    if (wrap && wrap.closest('.doc-img-float')) return wrap
  }
  return null
}

interface ImageParagraphAnchor {
  pos: number
  node: PmNode
  dom: HTMLElement
  rect: DOMRect
}

const IMAGE_ANCHOR_NODE_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

function imageParagraphAnchors(view: EditorView, imagePos: number): ImageParagraphAnchor[] {
  const anchors: ImageParagraphAnchor[] = []
  view.state.doc.forEach((node, pos) => {
    if (pos === imagePos || !IMAGE_ANCHOR_NODE_TYPES.has(node.type.name)) return
    const dom = view.nodeDOM(pos)
    if (!(dom instanceof HTMLElement)) return
    anchors.push({ pos, node, dom, rect: dom.getBoundingClientRect() })
  })
  return anchors
}

function pickImageParagraphAnchor(
  anchors: ImageParagraphAnchor[],
  x: number,
  y: number,
): ImageParagraphAnchor | null {
  let best: ImageParagraphAnchor | null = null
  let bestScore = Infinity
  for (const anchor of anchors) {
    const rect = anchor.dom.getBoundingClientRect()
    anchor.rect = rect
    const dy = y < rect.top ? rect.top - y : y > rect.bottom ? y - rect.bottom : 0
    const dx = x < rect.left ? rect.left - x : x > rect.right ? x - rect.right : 0
    // Vertical proximity determines the paragraph; horizontal proximity only
    // breaks ties for multi-column lines at a similar Y.
    const score = dy * 10000 + dx
    // Stacked paragraphs share a bottom/top edge. At that exact boundary,
    // prefer the paragraph beginning there instead of the one that just ended.
    if (score < bestScore || (score === bestScore && rect.top > (best?.rect.top ?? -Infinity))) {
      best = anchor
      bestScore = score
    }
  }
  return best
}

function floatingObjectDragPlugin(): Plugin {
  return new Plugin({
    props: {
      handleDOMEvents: {
        mousedown: (view, event) => {
          if (event.button !== 0) return false
          const target = event.target as HTMLElement | null
          if (!target) return false
          // Activate on the move handle of an image block, anywhere on the
          // image body (Word parity: grab the picture itself), or on a
          // textbox/shape body — its text isn't edited in place, so the
          // body gesture is unambiguous
          const handle = target.closest('.doc-move-handle') as HTMLElement | null
          const body = handle ? null : (target.closest('.doc-textbox') as HTMLElement | null)
          let imgBody: HTMLElement | null = null
          if (
            !handle &&
            !body &&
            !target.closest('.img-resize-handle') &&
            !document.body.classList.contains('docs-crop-active')
          ) {
            imgBody = target.closest('.doc-img-wrap') as HTMLElement | null
            // behind-text pictures sit under the text layer: a press that
            // hits no glyph falls through to the picture painted below
            if (!imgBody) imgBody = findFloatImageAt(event.clientX, event.clientY)
          }
          if (!handle && !body && !imgBody) return false
          const wrapper = (handle ?? body ?? imgBody)!.closest(
            '.doc-protected',
          ) as HTMLElement | null
          if (!wrapper) return false

          // Find the ProseMirror node position
          let pos = -1
          view.state.doc.descendants((node, p) => {
            if (pos !== -1) return false
            if (node.type.name === 'docProtected' && view.nodeDOM(p) === wrapper) pos = p
            return pos === -1
          })
          if (pos === -1) return false
          const node = view.state.doc.nodeAt(pos)
          if (!node) return false
          const isImage = node.attrs.blockType === 'image'
          const isTextbox = Array.isArray(node.attrs.textboxes) && node.attrs.textboxes.length > 0
          if (!isImage && !isTextbox) return false
          // Body activation is for prst shapes and pictures: plain text
          // boxes keep click-to-type
          const isShapeBody = isTextbox && !!(node.attrs.textboxes as TextboxDisplay[])[0]?.prst
          if (!handle && !isShapeBody && !(imgBody && isImage)) return false

          event.preventDefault()
          event.stopPropagation()
          view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)))
          // preventDefault suppressed the native focus: restore it, or
          // keyboard follow-ups (Delete, undo, arrows) go nowhere
          view.focus()

          const zoomEl = document.querySelector('.doc-zoom') as HTMLElement | null
          const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
          const startX = event.clientX
          const startY = event.clientY
          // Start position of the picture in its offset space (px). When the
          // attrs carry no numeric offset (fresh floats, align presets,
          // inline pictures) derive it from the rendered slot, so the drop
          // lands exactly where the user let go instead of jumping.
          const innerWrap = wrapper.querySelector('.doc-img-wrap') as HTMLElement | null
          const startVisualRect = (innerWrap ?? wrapper).getBoundingClientRect()
          const wrapperStartRect = wrapper.getBoundingClientRect()
          const anchorLocked = !!node.attrs.imageAnchorLocked
          const paragraphAnchors = isImage && !anchorLocked ? imageParagraphAnchors(view, pos) : []
          let paragraphAnchor: ImageParagraphAnchor | null = null
          const wrapMode = (node.attrs.imageWrap as string | null) ?? null
          const isSideFloat = !!wrapMode && /^(?:square|tight|through)-/.test(wrapMode)
          const imgW = innerWrap?.offsetWidth || Number(node.attrs.imageWidthPx ?? 0) || 0
          // the containing block the offsets resolve against (column width)
          let colW = wrapper.clientWidth
          if (isSideFloat && wrapper.parentElement) {
            const pcs = getComputedStyle(wrapper.parentElement)
            colW =
              wrapper.parentElement.clientWidth -
              (parseFloat(pcs.paddingLeft) || 0) -
              (parseFloat(pcs.paddingRight) || 0)
          }
          const attrX = node.attrs.imageOffsetXEmu
          const attrY = node.attrs.imageOffsetYEmu
          let startPxX: number
          if (attrX != null) {
            startPxX = Number(attrX) / EMU_PER_PX
          } else if (!wrapMode && innerWrap) {
            // inline picture: measure the visual spot inside its paragraph
            const wr = wrapper.getBoundingClientRect()
            const ir = innerWrap.getBoundingClientRect()
            startPxX = (ir.left - wr.left) / zoom
          } else {
            const posH = node.attrs.imagePosH as string | null
            const slot = wrapMode?.endsWith('-right')
              ? 'right'
              : wrapMode === 'topBottom' || posH === 'center'
                ? 'center'
                : posH === 'right'
                  ? 'right'
                  : 'left'
            startPxX =
              slot === 'center'
                ? Math.max(0, (colW - imgW) / 2)
                : slot === 'right'
                  ? Math.max(0, colW - imgW)
                  : 0
          }
          const startPxY = attrY != null ? Number(attrY) / EMU_PER_PX : 0

          // Visual feedback: apply CSS translate during drag. Cropped pictures put
          // rot/flip (and the front/behind offset) on the overflow-hidden crop
          // wrapper — translate that, or the image slides inside the fixed window
          const visual = wrapper.querySelector(
            isTextbox ? '.doc-textbox' : '.doc-img-crop, .doc-protected-img',
          ) as HTMLElement | null
          // Images may already carry a rotation/flip transform: the drag translate
          // must compose with it (prepended = applied in screen space) and the
          // original must come back on mouseup, or the orientation vanishes
          const baseTransform = visual?.style.transform ?? ''
          const anchorMarker = isImage
            ? (wrapper.querySelector('.doc-image-anchor-marker') as HTMLElement | null)
            : null
          const anchorMarkerTransform = anchorMarker?.style.transform ?? ''
          const anchorMarkerLeft = anchorMarker?.style.left ?? ''

          // 3px threshold keeps plain clicks (select, first click of a
          // double-click-to-edit) from nudging the object
          let dragging = false
          const onMove = (e: MouseEvent) => {
            const dx = (e.clientX - startX) / zoom
            const dy = (e.clientY - startY) / zoom
            if (!dragging && Math.abs(dx) < 3 && Math.abs(dy) < 3) return
            dragging = true
            if (visual) {
              visual.style.transform =
                `translate(${dx}px, ${dy}px)` + (baseTransform ? ` ${baseTransform}` : '')
            }
            if (isImage) {
              const clientDx = e.clientX - startX
              const clientDy = e.clientY - startY
              paragraphAnchor = pickImageParagraphAnchor(
                paragraphAnchors,
                startVisualRect.left + clientDx,
                startVisualRect.top + clientDy,
              )
              if (anchorMarker && paragraphAnchor) {
                const markerY = (paragraphAnchor.rect.top - wrapperStartRect.top) / zoom
                const markerX = (paragraphAnchor.rect.left - wrapperStartRect.left) / zoom - 32
                anchorMarker.style.transform = `translateY(${markerY.toFixed(1)}px)`
                anchorMarker.style.left = `${markerX.toFixed(1)}px`
                anchorMarker.dataset.anchorTargetPos = String(paragraphAnchor.pos)
              }
            }
          }

          const onUp = (e: MouseEvent) => {
            window.removeEventListener('mousemove', onMove)
            window.removeEventListener('mouseup', onUp)
            if (visual) visual.style.transform = baseTransform
            if (anchorMarker) {
              anchorMarker.style.transform = anchorMarkerTransform
              anchorMarker.style.left = anchorMarkerLeft
              delete anchorMarker.dataset.anchorTargetPos
            }

            const dx = (e.clientX - startX) / zoom
            const dy = (e.clientY - startY) / zoom
            // Only commit a real drag (past the threshold)
            if (!dragging) return

            // Word allows negative offsets (into the page margin / above the
            // anchor paragraph): no clamping
            const newX = Math.round((startPxX + dx) * EMU_PER_PX)
            if (isImage) {
              paragraphAnchor = pickImageParagraphAnchor(
                paragraphAnchors,
                startVisualRect.left + (e.clientX - startX),
                startVisualRect.top + (e.clientY - startY),
              )
            }
            const newY = Math.round(
              (paragraphAnchor
                ? (startVisualRect.top + (e.clientY - startY) - paragraphAnchor.rect.top) / zoom
                : startPxY + dy) * EMU_PER_PX,
            )

            const currentNode = view.state.doc.nodeAt(pos)
            if (!currentNode) return
            if (isTextbox) {
              view.dispatch(
                view.state.tr.setNodeMarkup(pos, undefined, {
                  ...currentNode.attrs,
                  imageWrap: currentNode.attrs.imageWrap ?? 'square-left',
                  imageOffsetXEmu: newX,
                  imageOffsetYEmu: newY,
                  imagePosH: null,
                  imagePosV: null,
                }),
              )
            } else {
              // pictures keep their wrap kind, but square/tight/through (and
              // fresh inline conversions) re-pick the wrap side from the drop
              // position — Word wraps text on the open side of the picture
              const prev = (currentNode.attrs.imageWrap as string | null) ?? null
              const kindMatch = prev ? /^(square|tight|through)-(?:left|right)$/.exec(prev) : null
              let nextWrap = prev ?? 'square-left'
              if (!prev || kindMatch) {
                const kind = kindMatch?.[1] ?? 'square'
                const centerX = startPxX + dx + imgW / 2
                nextWrap = `${kind}-${colW > 0 && centerX > colW / 2 ? 'right' : 'left'}`
              }
              const attrs = {
                ...currentNode.attrs,
                imageWrap: nextWrap,
                imageOffsetXEmu: newX,
                imageOffsetYEmu: newY,
                imagePosH: null,
                imagePosV: null,
              }
              if (
                paragraphAnchor &&
                !currentNode.attrs.imageAnchorLocked &&
                !currentNode.attrs.blockRevision
              ) {
                // A floating picture's top-level atom is its OOXML anchor
                // paragraph. Move that atom immediately before the paragraph
                // selected by the drag, preserving docxIndex as the patch
                // identity and rebasing Y to the new paragraph above.
                const movedNode = currentNode.type.create(
                  attrs,
                  currentNode.content,
                  currentNode.marks,
                )
                const tr = view.state.tr.delete(pos, pos + currentNode.nodeSize)
                const insertPos = tr.mapping.map(paragraphAnchor.pos)
                tr.insert(insertPos, movedNode)
                tr.setSelection(NodeSelection.create(tr.doc, insertPos))
                view.dispatch(tr)
              } else {
                view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, attrs))
              }
            }
          }

          window.addEventListener('mousemove', onMove)
          window.addEventListener('mouseup', onUp)
          return true
        },
      },
    },
  })
}

export type DomSpec = [string, Record<string, string>, ...unknown[]]

/**
 * Paragraph node for textbox sub-editors. Named docParagraph on purpose so
 * ribbon paragraph commands (updateAttributes('docParagraph', ...)) work
 * unchanged whether they target the main editor or a textbox.
 */
const TextboxParagraph = Node.create({
  name: 'docParagraph',
  group: 'block',
  content: 'inline*',
  addAttributes() {
    return {
      styleId: { default: null as string | null },
      align: { default: null as string | null },
      lineSpacing: { default: null as number | null },
      lineRule: { default: null as string | null },
      lineRawTwips: { default: null as number | null },
      snapToGrid: { default: null as boolean | null },
      indentLeft: { default: null as number | null },
      indentRight: { default: null as number | null },
      indentFirstLine: { default: null as number | null },
      spaceBefore: { default: null as number | null },
      spaceAfter: { default: null as number | null },
      spaceBeforeAuto: { default: null as boolean | null },
      spaceAfterAuto: { default: null as boolean | null },
      shadingFill: { default: null as string | null },
      borders: { default: null as string | null },
      listMarker: { default: null as TextboxListMarker | null },
    }
  },
  parseHTML() {
    return [{ tag: 'div.doc-textbox-para' }, { tag: 'p' }]
  },
  renderHTML({ node }) {
    // same line strut rules as the main editor's blockAttrs (Word applies the
    // typed line grid inside textboxes too; w:snapToGrid=0 opts out): without
    // them every CJK textbox line inherited the body's grid pixel value
    const lineRule = (node.attrs.lineRule as 'auto' | 'atLeast' | 'exact' | null) ?? undefined
    const lineRawTwips =
      node.attrs.lineRawTwips != null ? Number(node.attrs.lineRawTwips) : undefined
    const lineSpacing = node.attrs.lineSpacing ? Number(node.attrs.lineSpacing) : undefined
    const marker = node.attrs.listMarker as TextboxListMarker | null
    const classes = ['doc-textbox-para']
    // an empty list paragraph keeps its line box: Word still shows the marker
    if (node.content.size === 0 && !marker) classes.push('doc-textbox-para-empty')
    if ((lineRule === 'exact' && lineRawTwips) || lineRule === 'atLeast')
      classes.push('doc-lh-fixed')
    if (node.attrs.snapToGrid === false) classes.push('doc-nosnap')
    const attrs: Record<string, string> = { class: classes.join(' ') }
    if (node.attrs.styleId) attrs['data-style'] = String(node.attrs.styleId)
    const fontStyles: string[] = []
    if (node.textContent) {
      fontStyles.push(`--doc-line-factor:${paraLineFactor(node)}`)
      const fam = paraDeclaredFontFamily(node)
      if (fam) fontStyles.push(`font-family:${fam}`)
      const strut = explicitStrutHalfPoints(node)
      if (strut) fontStyles.push(...strutFontCss(strut))
    }
    const mult = cssAutoLineMult(lineRule, lineRawTwips, lineSpacing)
    if (marker) attrs['data-marker'] = marker.text
    // Word precedence: the paragraph's own w:ind (an explicit 0 included), else the numbering
    // level's, else the style's
    const indentLeft = (node.attrs.indentLeft as number | null) ?? marker?.indentLeft ?? null
    const ownFirstLine = node.attrs.indentFirstLine as number | null
    const firstLine = Number(ownFirstLine ?? marker?.firstLine ?? -(marker?.hanging ?? 0))
    const hanging = marker && firstLine < 0 ? -firstLine : 0
    // no hanging area at all: Word runs the marker to the next default tab stop (0.5in)
    const markerTab = marker && !hanging && firstLine <= 0 ? ';--li-tab:36pt' : ''
    const styles = [
      node.attrs.align
        ? `text-align:${node.attrs.align === 'distribute' ? 'justify' : node.attrs.align}`
        : '',
      ...fontStyles,
      cssLineHeight(lineRule, lineRawTwips, lineSpacing)
        ? `line-height:${cssLineHeight(lineRule, lineRawTwips, lineSpacing)}`
        : '',
      // explicit single (mult 1) still overrides an inherited style/doc multiple
      mult ? `--doc-line-mult:${mult}` : '',
      node.attrs.snapToGrid === false ? '--doc-grid-pitch:0.0001px' : '',
      // logical sides: w:ind left/right swap in bidi paragraphs (Word), and the marker box hangs
      // on the inline-start side
      indentLeft != null ? `margin-inline-start:${indentLeft / 20}pt` : '',
      node.attrs.indentRight != null
        ? `margin-inline-end:${Number(node.attrs.indentRight) / 20}pt`
        : '',
      // a list marker owns the hanging area (a positive first line instead puts the marker at
      // the first-line position); the style's first-line indent must not shift the text
      marker
        ? `--li-hang:${hanging / 20}pt;text-indent:${firstLine > 0 ? firstLine / 20 : 0}pt${markerTab}`
        : ownFirstLine != null
          ? `text-indent:${Number(ownFirstLine) / 20}pt`
          : '',
      marker?.szHalfPoints ? `--li-marker-size:${marker.szHalfPoints / 2}pt` : '',
      ...(marker?.symbol ? substituteMarkerStyles(marker.text) : []),
      node.attrs.spaceBeforeAuto
        ? `margin-top:${WORD_AUTO_SPACING_PT}pt`
        : node.attrs.spaceBefore != null
          ? `margin-top:${Number(node.attrs.spaceBefore) / 20}pt`
          : '',
      node.attrs.spaceAfterAuto
        ? `margin-bottom:${WORD_AUTO_SPACING_PT}pt`
        : node.attrs.spaceAfter != null
          ? `margin-bottom:${Number(node.attrs.spaceAfter) / 20}pt`
          : '',
      node.attrs.shadingFill
        ? `background-color:#${node.attrs.shadingFill};${dkBackground(`#${node.attrs.shadingFill}`)}`
        : '',
    ]
      .filter(Boolean)
      .join(';')
    if (styles) attrs.style = styles
    {
      const ink = fillInk(node.attrs.shadingFill)
      if (ink) attrs['data-ink'] = ink
    }
    return ['div', attrs, 0]
  },
})

/** shared with the main editor: same mark names, so ribbon commands route 1:1 */
const textboxSubExtensions = [
  Node.create({ name: 'doc', topNode: true, content: 'block+' }),
  DocText,
  DocHardBreak,
  TextboxParagraph,
  TabStopExtension,
  InactiveSelectionExtension,
  // inline pictures inside textbox paragraphs (form checkboxes): without the
  // node the sub-editor drops them on load and the first commit loses them
  DocInlineImage,
  BoldMark,
  ItalicMark,
  UnderlineMark,
  StrikeMark,
  LinkMark,
  // research-report sidebars keep PAGE/date/REF fields inside textbox tables;
  // without these marks the whole box fails to load into the sub-editor
  RefFieldMark,
  InstrFieldMark,
  TextStyleMark,
  CommentMark,
  UndoRedo,
]

// ---- find & replace highlighting ----

export interface SearchHighlight {
  ranges: Array<{ from: number; to: number }>
  activeIndex: number
}

/**
 * Word's AutoFormat-as-you-type for links (alpha ledger r151): a URL followed
 * by a space or Enter turns into a hyperlink. Runs on keydown BEFORE the key
 * itself applies (marks the URL, then lets the key proceed), matching Word's
 * behavior of linkifying the word just completed. Trailing punctuation stays
 * outside the link, and existing links are left alone.
 */
// \ufffc is textBetween's stand-in for inline leaves (breaks, images, note refs):
// a URL must stop there or the match would swallow the atom and the text after it
const AUTOLINK_URL = /(?:https?:\/\/|www\.)[^\s\ufffc]+$/i
const AUTOLINK_TRAILING = /[.,;:!?)\]}'"\u00bb\u203a]+$/

export const AutoLinkOnDelimiter = Extension.create({
  name: 'autoLinkOnDelimiter',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('autoLinkOnDelimiter'),
        props: {
          handleKeyDown(view, event) {
            if (event.key !== ' ' && event.key !== 'Enter') return false
            if (event.altKey || event.ctrlKey || event.metaKey) return false
            const { state } = view
            const { $from, empty } = state.selection
            if (!empty || !view.editable || !$from.parent.isTextblock) return false
            const linkType = state.schema.marks.link
            if (!linkType) return false
            const before = $from.parent.textBetween(0, $from.parentOffset, undefined, '\ufffc')
            const match = AUTOLINK_URL.exec(before)
            if (!match) return false
            let url = match[0]
            const trailing = AUTOLINK_TRAILING.exec(url)
            if (trailing) url = url.slice(0, -trailing[0].length)
            if (url.length < 5) return false
            const start = $from.pos - ($from.parentOffset - match.index)
            const end = start + url.length
            // already inside a link (typing after one, or re-delimiting): leave it
            if (state.doc.rangeHasMark(start, end, linkType)) return false
            const href = /^www\./i.test(url) ? `http://${url}` : url
            view.dispatch(state.tr.addMark(start, end, linkType.create({ href, rId: null })))
            return false // the space/Enter itself proceeds normally
          },
        },
      }),
    ]
  },
})

export const editorExtensions = [
  DocDocument,
  DocText,
  DocHardBreak,
  DocNoteRef,
  DocXeMark,
  DocRuby,
  DocInlineImage,
  DocInlineMath,
  DocParagraph,
  DocHeading,
  DocListItem,
  DocTable,
  DocTableRow,
  DocTableCell,
  DocTableHeader,
  DocNestedTable,
  DocCellBoxes,
  DocProtected,
  BoldMark,
  ItalicMark,
  UnderlineMark,
  StrikeMark,
  LinkMark,
  RefFieldMark,
  InstrFieldMark,
  RprChangeMark,
  TextStyleMark,
  CommentMark,
  InsMark,
  DelMark,
  UndoRedo,
  SearchHighlightExtension,
  PendingCommentHighlightExtension,
  ResolvedCommentsExtension,
  NativeTableSupport,
  TableHandle,
  TrackChangesExtension,
  LineFactorExtension,
  ListNumberingExtension,
  AnchorLineExtension,
  PaginationGapsExtension,
  RowFillsExtension,
  FloatVShiftsExtension,
  InactiveSelectionExtension,
  AiQueueAnchorsExtension,
  PageGapNavExtension,
  ImageCopyExtension,
  EnterReplacesSelection,
  AutoLinkOnDelimiter,
  WordEditorShortcuts,
  CaretMarksMemory,
  ColumnLayoutExtension,
  TabStopExtension,
  JustifyShrinkExtension,
  CjkPunctShrinkExtension,
  WsRunLineHeightExtension,
  EaHintQuotesExtension,
  DropCapExtension,
  ParaBorderMergeExtension,
  SdtExtension,
  MoveRevisionExtension,
  PPrChangeExtension,
  ParaMarkDelExtension,
  RevisionOriginalExtension,
  AutoDirectionExtension,
]
