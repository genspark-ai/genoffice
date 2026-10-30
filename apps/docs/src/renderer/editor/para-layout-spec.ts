/**
 * Paragraph line-box computation: script and declared-face line factors, strut
 * face/size, line rule, spacing and indents, over a schema-neutral LayoutPara.
 * fromPmNode serves the editor schema (docParagraph / docHeading / docListItem
 * renderHTML, image-paragraph wrappers, the live line-factor decorations, with
 * borders and frames on top); fromRuns serves the read-only cell and textbox
 * paragraphs rendered from engine Run[] (protected-render); fromHfParagraph
 * serves the header/footer strip paragraphs (hf-dom, HeaderFooterArea).
 */
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  asciiOnlyCjkFace,
  autoSpacingPt,
  cjkDeclaredLineFactor,
  cjkScriptRanges,
  cssCsFontFamily,
  cssExactLineCap,
  cssFontFamily,
  cssGridSpacingPt,
  cssLeadTop,
  cssLineHeight,
  cssRunFontFamily,
  cutAdjustedAlias,
  emptyEaSlotHangulFactor,
  facesDiffer,
  fontBoxCss,
  isCjkFontName,
  lineHeightFactor,
  nonCjkInkRanges,
  paraLineFactorCss,
  runLatinFactorFaces,
  SPACE_ONLY_RE,
  spaceBeforeProp,
  strutFontCss,
  textHasCjk,
  textHasComplexScript,
  textHasHangul,
  textHasLatinInk,
} from '../line-metrics'
import type {
  NoteRun,
  ParaFormat,
  ParaFrame,
  ParaFrameBox,
  Run,
  TextFlowDirection,
} from '@genoffice/docx-engine'
import { inferredBidi } from './direction'
import { t } from '../i18n/locale'
import { textAlignDecl } from './text-effects'
import {
  PARA_BORDER_SHADOW,
  paraBorderCss,
  paraBorderLinesWithPad,
  paraBorderPadding,
  paraBorderPaddingDecls,
  paraBorderShadowDecls,
} from './para-border-css'
import { paraFrameCss } from './para-frame'
import { DK_SIDE, dkBackground, dkBorder } from './dark-page'
import { fillInk } from './shading-ink'
import { quarterTurnInsetPx } from './image-rotation'

export type LayoutRunAttrs = {
  font?: string | null
  fontAscii?: string | null
  eastAsiaFont?: string | null
  csFont?: string | null
  cs?: boolean | null
  sizeHalfPoints?: number | null
  csSize?: number | null
  eaSlotEmpty?: boolean | null
}

/** one inline leaf of a paragraph, in document order */
export interface LayoutRun {
  kind: 'text' | 'break' | 'image' | 'other'
  text: string
  /** docTextStyle attrs of a text leaf; null when it carries no such mark */
  attrs: LayoutRunAttrs | null
  /** text leaf without any mark: it has no element of its own to restore a size */
  bare: boolean
  bold: boolean
  italic: boolean
  pageOrColBreak: boolean
}

/** what the line-box rules read of a paragraph, whichever schema it came from */
export interface LayoutPara {
  text: string
  runs: LayoutRun[]
  childCount: number
  lineRule?: 'auto' | 'atLeast' | 'exact' | null
}

export function layoutParaOf(node: {
  attrs?: Record<string, unknown>
  textContent?: string
  childCount?: number
  descendants?: PmNode['descendants']
}): LayoutPara {
  const runs: LayoutRun[] = []
  node.descendants?.((child) => {
    if (child.isText) {
      const mark = child.marks.find((m) => m.type.name === 'docTextStyle')
      runs.push({
        kind: 'text',
        text: child.text ?? '',
        attrs: mark ? (mark.attrs as LayoutRunAttrs) : null,
        bare: child.marks.length === 0,
        bold: child.marks.some((m) => m.type.name === 'bold'),
        italic: child.marks.some((m) => m.type.name === 'italic'),
        pageOrColBreak: false,
      })
      return false
    }
    const name = child.type.name
    runs.push({
      kind: name === 'hardBreak' ? 'break' : name === 'docInlineImage' ? 'image' : 'other',
      text: '',
      attrs: null,
      bare: false,
      bold: false,
      italic: false,
      pageOrColBreak: name === 'hardBreak' && !!(child.attrs.pageBreak || child.attrs.colBreak),
    })
    return true
  })
  return {
    text: node.textContent ?? '',
    runs,
    childCount: node.childCount ?? 0,
    lineRule: node.attrs?.lineRule as LayoutPara['lineRule'],
  }
}

/** Word lays a space-only paragraph out like an empty one: the mark sizes the line;
 *  page/column breaks beside the spaces keep it mark-sized */
export function isSpaceOnlyParagraph(para: LayoutPara): boolean {
  if (!para.childCount || !SPACE_ONLY_RE.test(para.text)) return false
  return para.runs.every((r) => r.kind === 'text' || r.pageOrColBreak)
}

function isBreakOnlyParagraph(para: LayoutPara): boolean {
  return para.runs.length > 0 && para.runs.every((r) => r.kind === 'break')
}

/** run size the mark paints: its own w:sz(Cs), else the Cs size an rtl run inherits */
export function markSizeHalfPoints(
  attrs: Record<string, unknown> | null | undefined,
): number | null {
  return (attrs?.sizeHalfPoints ?? attrs?.csSize ?? null) as number | null
}

/** Smallest explicit run size (half-points) carried by the strut. `mixed`: some
 *  runs inherit the body size; they render in their own elements and take it
 *  back through .doc-strut-mixed, so a 9pt sub-line under inherited 11pt text
 *  is 9pt tall like Word (production sample: Korean cells with English sub-lines
 *  paced 14.5pt for Word's 11.9). A bare inheriting text node has no element to
 *  restore its size, so such paragraphs keep the inherited strut. */
export function strutHalfPoints(para: LayoutPara): { halfPoints: number; mixed: boolean } | null {
  let min: number | null = null
  let mixed = false
  for (const r of para.runs) {
    if (r.kind !== 'text') continue
    // space-only runs never size a line (Word probe 2026-09-11)
    if (SPACE_ONLY_RE.test(r.text)) continue
    const sz = markSizeHalfPoints(r.attrs)
    if (sz != null) min = Math.min(min ?? Infinity, sz)
    else if (!r.bare) mixed = true
    else return null
  }
  return min === null ? null : { halfPoints: min, mixed }
}

export const STRUT_MIXED_CLASS = 'doc-strut-mixed'

/** Declared-face CJK paragraph whose chain holds two faces (a Latin face and
 *  an East Asian face): Chromium unions the strut's Latin-led box with the
 *  --doc-east-asian-font spans' box and small lines outgrow their line-height
 *  (8.5pt YaHei + Calibri cells: +1pt per row), so .doc-grid-strut leads every
 *  box with the metric strut face. A single declared face (Meiryo UI on every
 *  slot) has no union to tame, and the strut face would place the line by the
 *  document base face's box, 2-3px above Word, which sizes each line by its
 *  own face. */
export function declaredCjkStrutClass(para: LayoutPara): string | undefined {
  if (!textHasCjk(para.text)) return undefined
  const twoFaces = para.runs.some(
    (r) =>
      r.kind === 'text' && textHasCjk(r.text) && facesDiffer(r.attrs?.fontAscii, r.attrs?.font),
  )
  return twoFaces ? 'doc-grid-strut' : undefined
}

/** text runs that size a line (space-only runs never do, Word probe 2026-09-11) */
function inkRuns(para: LayoutPara): LayoutRun[] {
  return para.runs.filter((r) => r.kind === 'text' && !SPACE_ONLY_RE.test(r.text))
}

/** true when every ink run paints one size (uniform or all inherited) */
function oneRunSize(runs: LayoutRun[]): boolean {
  let size: number | null | undefined
  for (const r of runs) {
    const sz = markSizeHalfPoints(r.attrs)
    if (size === undefined) size = sz
    else if (size !== sz) return false
  }
  return true
}

/**
 * Cell paragraph whose declared Latin faces differ in line factor: each run
 * keeps its own line box (styles.css .doc-mixed-face) because a shorter face
 * centred in the inherited line-height lifts the line above the strut, while
 * Word lays a mixed line at max ascent + max descent (probe 2026-09-30).
 * Numeric strut factor only (every run declares a face) and one run size
 * (uniform or all inherited): then no run box can exceed the strut line, so
 * the strut keeps carrying a typed grid in adjustLineHeightInTable cells.
 */
export function paraMixedFaceClass(para: LayoutPara, factor: string): string | undefined {
  const max = Number(factor)
  if (!Number.isFinite(max) || textHasCjk(para.text)) return undefined
  const runs = inkRuns(para)
  if (!oneRunSize(runs)) return undefined
  const mixed = runs.some((r) =>
    runLatinFactorFaces(r.text, r.attrs?.fontAscii, r.attrs?.csFont, r.attrs?.cs).some(
      (family) => family && lineHeightFactor(family) < max - 1e-3,
    ),
  )
  return mixed ? 'doc-mixed-face' : undefined
}

/**
 * Mixed Latin + complex-script paragraph at one run size: every cs span
 * (marks.ts data-cs) gives up its own line box (styles.css .doc-cs-mixed) and
 * the strut, already at the tallest face's factor, sizes the line. Chromium
 * centres each face in the inherited line-height, so an Arial Arabic span under
 * a Calibri strut paced 1.31em where Word paces 1.22 (probe 2026-10-13). One
 * run size only: with mixed sizes the strut shrinks to the smallest run.
 */
export function paraCsMixedClass(para: LayoutPara): string | undefined {
  const text = para.text
  if (para.lineRule === 'exact' || para.lineRule === 'atLeast') return undefined
  if (!textHasComplexScript(text) || !/[A-Za-z]/.test(text) || textHasCjk(text)) return undefined
  const runs = inkRuns(para)
  if (!oneRunSize(runs)) return undefined
  const csRun = runs.some((r) => r.attrs?.csFont && textHasComplexScript(r.text))
  return csRun ? 'doc-cs-mixed' : undefined
}

/**
 * One-size Latin paragraph with a bold/italic run of a per-cut size-adjusted
 * alias (line-metrics cutAdjustedAlias): the span's box differs in size from
 * the regular-cut strut and the union inflates the line (Merriweather 10.5pt
 * label/value lines paced 13.5 where Word paces 13.2, prod corpus 20261020
 * sample 043). The run gives up its own box (styles.css .doc-cut-adjusted).
 */
export function paraCutAdjustedClass(para: LayoutPara): string | undefined {
  if (para.lineRule === 'exact' || para.lineRule === 'atLeast') return undefined
  if (textHasCjk(para.text)) return undefined
  const runs = inkRuns(para)
  if (!oneRunSize(runs)) return undefined
  const cutRun = runs.some(
    (r) =>
      r.attrs?.fontAscii &&
      (r.bold || r.italic) &&
      cutAdjustedAlias(cssRunFontFamily(String(r.attrs.fontAscii), null)),
  )
  return cutRun ? 'doc-cut-adjusted' : undefined
}

/**
 * Latin paragraphs: runs declaring a font override the doc-level factor with that
 * face's metric (Word sizes lines by run fonts, not the docDefaults face — a
 * Cambria-themed doc whose runs all say Calibri lays out at 1.22, not 1.17);
 * runs inheriting the body face keep the doc var via max().
 */
export function latinParaFactor(para: LayoutPara, scriptVar: string, latinInkOnly = false): string {
  let declaredMax = 0
  let undeclared = false
  for (const r of para.runs) {
    if (r.kind !== 'text') continue
    if (latinInkOnly && !textHasLatinInk(r.text)) continue
    // per-line mode lifts an ascii-only CJK face through its own stretches
    if (latinInkOnly && r.attrs && asciiOnlyCjkFace(r.attrs)) continue
    // ascii/hAnsi slot only: Word lays Latin text with the (possibly inherited)
    // ascii face, so an eastAsia-only declaration (attrs.font) must not drag
    // its factor onto a Latin line (EA "Arial Unicode MS" = 1.74, sample 13)
    for (const family of runLatinFactorFaces(
      r.text,
      r.attrs?.fontAscii,
      r.attrs?.csFont,
      r.attrs?.cs,
    )) {
      if (family) declaredMax = Math.max(declaredMax, lineHeightFactor(family))
      else undeclared = true
    }
  }
  if (declaredMax <= 0) return scriptVar
  return undeclared ? `max(${scriptVar}, ${declaredMax})` : String(declaredMax)
}

/**
 * Complex-script-only paragraph (Arabic/Hebrew text, no Latin letters): Word
 * lays every line in the w:cs face, so the strut takes that face as well. A
 * Latin strut (Calibri) under cs spans (Arial) unions two differently centred
 * boxes and paces 1.25em where Word paces Arial's 1.15 (probe 2026-10-13).
 * Null when a run lacks a cs face or the runs' cs chains differ.
 */
export function paraComplexScriptFamily(para: LayoutPara): string | null {
  const text = para.text
  if (!textHasComplexScript(text) || /[A-Za-z]/.test(text) || textHasCjk(text)) return null
  let family: string | null = null
  for (const r of inkRuns(para)) {
    const cs = r.attrs?.csFont ? String(r.attrs.csFont) : null
    if (!cs) return null
    // same chain the run's span renders (marks.ts), so the strut face equals the run face
    const chain = cssCsFontFamily(
      cs,
      r.attrs?.fontAscii ? String(r.attrs.fontAscii) : undefined,
      r.attrs?.font ? String(r.attrs.font) : undefined,
    )
    if (family === null) family = chain
    else if (family !== chain) return null
  }
  return family
}

/** Paragraph font-family follows the runs when every text run declares one:
 *  Chromium's line box is the union of the strut (paragraph font) and run boxes,
 *  so a paragraph face with a taller ascent than the runs' inflates every line
 *  past the computed line-height. No run inherits the face, so only the strut
 *  (and list markers without their own font) changes. */
export function paraRunFontFamily(para: LayoutPara): string | null {
  let first: string | null = null
  for (const r of para.runs) {
    if (r.kind !== 'text') continue
    const ea = r.attrs?.font ? String(r.attrs.font) : null
    const ascii = r.attrs?.fontAscii ? String(r.attrs.fontAscii) : null
    if (!ea && !ascii) return null
    // same chain the run's span renders, so the strut face equals the run face
    first ??= cssRunFontFamily(ascii, ea)
  }
  return first
}

export function paraDeclaredFontFamily(para: LayoutPara): string | null {
  return paraComplexScriptFamily(para) ?? paraRunFontFamily(para)
}

/** Mixed paragraph: some runs declare a CJK-named face, others inherit. The
 *  strut cannot take the declared family wholesale (inherited Latin text would
 *  change face), so typed-grid docs align its geometry through the glyphless
 *  metrics face instead (doc-style-css .doc-grid-strut; inert elsewhere). */
export function paraMixedDeclaredCjk(para: LayoutPara): boolean {
  let declaredCjk = false
  let inherited = false
  for (const r of para.runs) {
    if (r.kind !== 'text') continue
    const family = r.attrs?.font ?? r.attrs?.fontAscii
    if (!family) inherited = true
    else if (isCjkFontName(family)) declaredCjk = true
    if (declaredCjk && inherited) return true
  }
  return false
}

/** .doc-ea-strut: inheriting paragraphs whose CJK stretches render through the
 *  --doc-east-asian-font spans, a primary face other than the strut's. Only
 *  documents whose EA face carries a Word metric alias style it (doc-style-css). */
export function eaStrutClass(para: LayoutPara): string | undefined {
  if (paraMixedDeclaredCjk(para)) return 'doc-grid-strut'
  return textHasCjk(para.text) ? 'doc-ea-strut' : undefined
}

/**
 * Per-paragraph --doc-line-factor value: CJK runs with a declared font take that
 * font's LO-metric factor (max over runs); undeclared CJK runs keep the
 * document-level CJK var; non-CJK paragraphs keep the script-based guess.
 */
export function paraLineFactor(para: LayoutPara): string {
  const scriptVar = paraLineFactorCss(para.text)
  if (!textHasCjk(para.text)) return latinParaFactor(para, scriptVar)
  let declaredMax = 0
  let undeclaredCjk = false
  for (const r of para.runs) {
    if (r.kind !== 'text' || !textHasCjk(r.text)) continue
    const attrs = r.attrs
    // empty-EA-theme-slot backfills keep the Word-look face but are not a
    // document font choice: LO cascades such runs to the document's EA default,
    // so they count as undeclared here (the doc-level var carries that factor).
    // Latin-named faces likewise don't drive CJK line height — an ascii-only
    // "Times New Roman" run still renders its CJK via the inherited EA font,
    // and so does an ascii-only CJK-named face (probe 2026-09-30).
    const family =
      attrs?.eaSlotEmpty === true || (attrs && asciiOnlyCjkFace(attrs))
        ? null
        : (attrs?.font ?? attrs?.fontAscii)
    const emptySlotKr =
      attrs?.eaSlotEmpty === true ? emptyEaSlotHangulFactor(attrs.font ?? null, r.text) : null
    if (emptySlotKr !== null) declaredMax = Math.max(declaredMax, emptySlotKr)
    else if (family && isCjkFontName(family)) {
      declaredMax = Math.max(declaredMax, cjkDeclaredLineFactor(family) ?? lineHeightFactor(family))
    } else undeclaredCjk = true
  }
  if (declaredMax <= 0) return scriptVar
  return undeclaredCjk ? `max(${scriptVar}, ${declaredMax})` : String(declaredMax)
}

/** a picture paragraph's mark line follows its Latin face; the eastAsia slot
 *  only counts when it is the sole declared face */
export function pictureMarkLineFactor(
  latin: string | null,
  eastAsia: string | null,
): string | null {
  if (latin) return String(lineHeightFactor(latin))
  if (eastAsia && isCjkFontName(eastAsia))
    return String(cjkDeclaredLineFactor(eastAsia) ?? lineHeightFactor(eastAsia))
  return null
}

export type RunFactorRange = { from: number; to: number; style: string }

/**
 * Word sizes every line by the faces actually on it (probe 2026-09-06: a
 * Latin-only line inside a Malgun Gothic paragraph is the Calibri line, and
 * KR/Latin neighbours sum to one line of each). A paragraph mixing CJK and
 * Latin ink therefore lays its strut at the Latin base and lifts only its CJK
 * stretches (inline decorations); Chromium's line box is the max of the inline
 * boxes on the line, so a line grows only when a CJK stretch lands on it.
 * Pure-CJK paragraphs keep the paragraph-level factor. exact lines never lift.
 */
export function perLineFactors(
  node: PmNode,
  para: LayoutPara = layoutParaOf(node),
): { strut: string; runs: RunFactorRange[] } | null {
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
      const asciiCjk = mark ? asciiOnlyCjkFace(mark.attrs) : null
      const family =
        mark?.attrs.eaSlotEmpty === true || asciiCjk
          ? null
          : ((mark?.attrs.font ?? mark?.attrs.fontAscii) as string | null | undefined)
      const emptySlotKr =
        mark?.attrs.eaSlotEmpty === true
          ? emptyEaSlotHangulFactor(mark.attrs.font as string | null, child.text)
          : null
      const factor =
        emptySlotKr !== null
          ? String(emptySlotKr)
          : family && isCjkFontName(family)
            ? String(cjkDeclaredLineFactor(family) ?? lineHeightFactor(family))
            : scriptVar
      const style = `--doc-line-factor:${factor}${lh ? `;line-height:${lh}` : ''}`
      for (const r of cjkScriptRanges(child.text)) {
        runs.push({ from: offset + r.from, to: offset + r.to, style })
      }
      // the ascii-only CJK face draws only the non-CJK glyphs; Word lifts just their lines
      if (asciiCjk) {
        const lift = `--doc-line-factor:${lineHeightFactor(asciiCjk)}${lh ? `;line-height:${lh}` : ''}`
        for (const r of nonCjkInkRanges(child.text)) {
          runs.push({ from: offset + r.from, to: offset + r.to, style: lift })
        }
      }
    }
    offset += child.nodeSize
  })
  return { strut: latinParaFactor(para, 'var(--doc-line-factor-latin,1.2)', true), runs }
}

/**
 * Paragraph FORMATTING attrs that survive the clipboard HTML round-trip
 * (renderHTML emits them as data-para JSON; parseHTML restores them), typed so
 * a crafted/corrupt payload can't smuggle wrong-typed values into the model.
 * Identity/anchor attrs stay out on purpose: a pasted paragraph is NEW content,
 * so docxIndex (save patch anchor), bookmarks, comment endpoints, sdtShell and
 * revision metadata must not be duplicated by copy/paste.
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
  indentFirstLineChars: 'number',
  spaceBefore: 'number',
  spaceAfter: 'number',
  spaceBeforeAuto: 'boolean',
  spaceAfterAuto: 'boolean',
  contextualSpacing: 'boolean',
  keepNext: 'boolean',
  keepLines: 'boolean',
  widowControl: 'boolean',
  suppressLineNumbers: 'boolean',
  pageBreakBefore: 'boolean',
  bidi: 'boolean',
  bidiInherited: 'boolean',
  bidiInferred: 'boolean',
  autoSpace: 'boolean',
  wordWrap: 'boolean',
  overflowPunct: 'boolean',
  eaLang: 'string',
  shadingFill: 'string',
  shadingDisplay: 'string',
  shadingClear: 'boolean',
  borderReset: 'string',
  frameBox: 'string',
  emptyRunSize: 'number',
  emptyRunFont: 'string',
  markSize: 'number',
  borders: 'string',
  borderLines: 'string',
  borderPad: 'string',
  outlineOnly: 'boolean',
  tabStops: 'string',
  dropCap: 'string',
  frame: 'string',
  textDirection: 'string',
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
    if (
      v === false &&
      (key === 'pageBreakBefore' ||
        key === 'bidi' ||
        key === 'bidiInherited' ||
        key === 'bidiInferred')
    ) {
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
    if (typeof v !== type && !(key === 'autoSpace' && (v === 'latin' || v === 'digits'))) continue
    if (type === 'number' && !Number.isFinite(v)) continue
    attrs[key] = v
  }
  return Object.keys(attrs).length > 0 ? attrs : null
}

/** every inline child is an in-flow picture (no wrap, no rule): the paragraph's
 *  line is the picture extent (CSS .doc-p-pic) */
function isPictureOnlyContent(node: {
  childCount?: number
  descendants?: PmNode['descendants']
}): boolean {
  if (!node.descendants) return false
  let ok = true
  node.descendants((child) => {
    if (child.type.name !== 'docInlineImage' || child.attrs.wrap || child.attrs.rule) ok = false
    return false
  })
  return ok
}

/** wrapTopAndBottom below its anchor paragraph top: rendered as a column-wide
 *  float the host block contains (doc-band-host) */
export function isBandOffsetPicture(attrs: Record<string, unknown>): boolean {
  if (attrs.wrap !== 'topBottom' || attrs.offsetYEmu == null) return false
  if (attrs.relV === 'page' || attrs.relV === 'margin') return false
  if (Number(attrs.offsetYEmu) <= 0) return false
  return !quarterTurnInsetPx(
    Number(attrs.widthPx),
    Number(attrs.heightPx),
    attrs.rotDeg as number | null,
  )
}

function hostsBandOffsetPicture(node: { descendants?: PmNode['descendants'] }): boolean {
  if (!node.descendants) return false
  let found = false
  node.descendants((child) => {
    if (child.type.name === 'docInlineImage' && isBandOffsetPicture(child.attrs)) found = true
    return !found
  })
  return found
}

/** hosts an anchored run picture whose X measures from the column (not page / margin) */
function hostsColumnAnchoredPicture(node: { descendants?: PmNode['descendants'] }): boolean {
  if (!node.descendants) return false
  let found = false
  node.descendants((child) => {
    if (
      child.type.name === 'docInlineImage' &&
      child.attrs.wrap &&
      child.attrs.relH !== 'page' &&
      child.attrs.relH !== 'margin'
    )
      found = true
    return !found
  })
  return found
}

interface TblParaProps {
  align?: string
  spaceBefore?: number
  spaceAfter?: number
  line?: { rule: 'auto' | 'atLeast' | 'exact'; raw: number }
}

function tblParaOf(attrs: Record<string, unknown>): TblParaProps | undefined {
  if (typeof attrs.tblPara !== 'string') return undefined
  try {
    return JSON.parse(attrs.tblPara) as TblParaProps
  } catch {
    return undefined
  }
}

export interface ParaLayoutNode {
  attrs: Record<string, unknown>
  textContent?: string
  childCount?: number
  firstChild?: PmNode | null
  descendants?: PmNode['descendants']
}

export interface ParaLayoutOptions {
  includeIndent?: boolean
  listGeometry?: boolean
}

/**
 * Paragraph line-box spec: classes and inline declarations in emission order,
 * plus the other rendered attributes. `attrs` may already hold a `class`
 * entry: its key position (hence HTML attribute order) is part of the output.
 */
export interface ParaLayoutSpec {
  classes: string[]
  styles: string[]
  attrs: Record<string, string>
}

export function toStyleString(spec: ParaLayoutSpec): string {
  return spec.styles.join(';')
}

export function toAttrs(spec: ParaLayoutSpec): Record<string, string> {
  const attrs = { ...spec.attrs }
  if (spec.classes.length > 0) attrs['class'] = spec.classes.join(' ')
  if (spec.styles.length > 0) attrs['style'] = toStyleString(spec)
  return attrs
}

function numOrNull(v: unknown): number | null {
  return v == null ? null : Number(v)
}

export type ParaLayoutKind = 'body' | 'cell' | 'textbox' | 'hf' | 'note'

export interface ParaLayoutContext {
  kind: ParaLayoutKind
  /** body: emit margin-inline-start for indentLeft (list items indent through padding) */
  includeIndent?: boolean
  listGeometry?: boolean
  /** body paragraph hosting a column-anchored run picture: --anchor-ind / --anchor-pad
   *  carry the margin and padding parts of the start indent the picture subtracts */
  anchorHost?: boolean
  /** hf cell: the paragraph / table style's run size that runs without their own w:sz inherit */
  styleSizeHalfPoints?: number | null
}

/** paragraph formatting the line-box rules read, table-style layer already merged */
export interface LayoutParaFormat {
  align?: string | null
  bidi?: boolean | null
  bidiInferred?: boolean | null
  autoSpace?: unknown
  wordWrap?: boolean | null
  lineRule?: 'auto' | 'atLeast' | 'exact' | null
  lineRawTwips?: number | null
  lineSpacing?: number | null
  /** table-style line rule, used when the paragraph declares no line of its own */
  tableStyleLine?: { rule: 'auto' | 'atLeast' | 'exact'; raw: number } | null
  snapToGrid?: boolean | null
  indentLeft?: number | null
  indentRight?: number | null
  indentFirstLine?: number | null
  indentFirstLineChars?: number | null
  spaceBefore?: number | null
  spaceAfter?: number | null
  spaceBeforeAuto?: boolean | null
  spaceAfterAuto?: boolean | null
  contextualSpacing?: boolean | null
  emptyRunSize?: number | null
  emptyRunFont?: string | null
  /** in-flow pictures and nothing else */
  picOnly: boolean
}

export interface LineBoxSpec {
  classes: string[]
  styles: string[]
}

/**
 * The line box of one paragraph: direction, line factor, strut face and size,
 * line-height rule, grid snapping, indents and spacing, as classes and inline
 * declarations in emission order. Every surface (body, read-only cell,
 * read-only textbox) goes through here; where a surface differs the
 * difference is an explicit `kind` branch with its reason, never an omission.
 */
export function lineBoxSpec(
  para: LayoutPara,
  fmt: LayoutParaFormat,
  ctx: ParaLayoutContext,
): LineBoxSpec {
  const { kind } = ctx
  const hf = kind === 'hf'
  // footnote / endnote rows: the reservation model pins the row's line height
  // (px) and base size so the reserved and rendered heights agree, the spec
  // supplies direction, wrapping, the strut face and the face classes
  const note = kind === 'note'
  const classes: string[] = []
  const styles: string[] = []
  const text = para.text

  if (fmt.bidi || fmt.bidiInferred) {
    styles.push('direction:rtl', 'unicode-bidi:isolate')
  } else {
    // explicit: paragraph direction is its own w:bidi only — a bidiVisual
    // table's dir="rtl" mirrors column order but must not reorder cell text
    styles.push('direction:ltr')
  }
  // explicit autoSpaceDE/DN off also disables the browser's native 1/8em gap;
  // with only one switch off the pads supply the remaining kind of gap alone.
  // Read-only surfaces set it on every run span instead (textSpanSpec)
  if (kind === 'body' && fmt.autoSpace != null && fmt.autoSpace !== true) {
    styles.push('text-autospace:no-autospace')
  }
  // cell alignment is painted on the td (renderTableSpec); a strip paragraph's
  // jc yields to its w:framePr placement, so hf-dom paints it
  if (kind !== 'cell' && !hf && !note && fmt.align) styles.push(textAlignDecl(String(fmt.align)))
  // the line-height factor follows paragraph content (approximating Word's max-of-inline-fonts
  // line height): CJK paragraphs get the CJK factor, pure-Western ones the document's
  // font-aware Latin factor (doc-style-css sets --doc-line-factor-latin per body font).
  // Runs that DECLARE a font override the script guess with that font's factor
  // (LO probe: line height follows the requested face's metrics even when CJK
  // glyphs fall through to another font — EA "Times New Roman" lays at 1.15em);
  // the paragraph takes the max over its CJK runs, like Word's tallest-run rule.
  const spaceOnly = isSpaceOnlyParagraph(para)
  if (text && !spaceOnly) {
    // Word breaks Korean at spaces (UAX#14 default would break between syllables);
    // scoped to Hangul-bearing paragraphs so CJ text keeps per-char breaking.
    // overflow-wrap keeps the sim's overlong-word hard-break fallback.
    if (textHasHangul(text) && fmt.wordWrap !== false) {
      styles.push('word-break:keep-all', 'overflow-wrap:anywhere')
    }
    const factor = paraLineFactor(para)
    styles.push(`--doc-line-factor:${factor}`)
    // strip paragraph: runs without their own w:sz inherit the paragraph /
    // table style's size (hf cells), which .doc-strut-mixed restores through
    // --doc-base-fs; the strip base itself is the Normal size already
    const strut = strutHalfPoints(para)
    if (hf && ctx.styleSizeHalfPoints && (!strut || strut.mixed)) {
      styles.push(`--doc-base-fs:${ctx.styleSizeHalfPoints / 2}pt`)
      if (!strut) styles.push(`font-size:${ctx.styleSizeHalfPoints / 2}pt`)
    }
    // CJK lines follow the East Asian factor rule, their glyph rest is unprobed
    if (textHasCjk(text)) styles.push('--doc-lead-gap:0px')
    // the strut face follows the runs on every surface: a letterhead of Times
    // New Roman 8pt lines pitches 9.2 in Word (probe 2026-10-29) and a strip
    // strut left on the inherited face added ~0.5pt per line through the
    // strut/run baseline mismatch
    const fam = paraDeclaredFontFamily(para)
    const cls = fam
      ? (paraMixedFaceClass(para, factor) ?? declaredCjkStrutClass(para))
      : eaStrutClass(para)
    if (fam) styles.push(`font-family:${fam}`, `--doc-grid-strut-tail:${fam}`, ...fontBoxCss(fam))
    if (cls) classes.push(cls)
    const csMixed = paraCsMixedClass(para)
    if (csMixed) classes.push(csMixed)
    const cutAdjusted = paraCutAdjustedClass(para)
    if (cutAdjusted) classes.push(cutAdjusted)
    // Word's line strut follows run sizes; without this the paragraph inherits the
    // body size (often larger than table-cell runs) and every line box inflates.
    // the smallest declared run size carries the strut; larger runs lay their own boxes (strutFontCss)
    if (strut && !note) {
      styles.push(...strutFontCss(strut.halfPoints, strut.mixed))
      if (strut.mixed) classes.push(STRUT_MIXED_CLASS)
    }
  } else {
    // Word sizes an empty line by the paragraph mark / empty run, both directions
    if (fmt.emptyRunSize) styles.push(`font-size:${fmt.emptyRunSize / 2}pt`)
    else if (hf && ctx.styleSizeHalfPoints)
      styles.push(`font-size:${ctx.styleSizeHalfPoints / 2}pt`)
    // the mark face sizes the empty line, CJK included (empty-line probe
    // 2026-08-25: SimSun 1.30 / DengXian 1.36 / Malgun 1.74 / Calibri 1.24 —
    // each face's own text factor; the earlier Western-only scoping starved
    // CJK marks down to the document factor and doubled their grid rows)
    const fam = fmt.emptyRunFont ?? null
    if (fam) {
      const chain = cssFontFamily(fam)
      styles.push(
        `--doc-line-factor:${lineHeightFactor(fam)}`,
        `font-family:${chain}`,
        ...fontBoxCss(chain),
      )
      if (isCjkFontName(fam)) styles.push('--doc-lead-gap:0px')
    } else if (kind !== 'cell' && (!para.childCount || spaceOnly || isBreakOnlyParagraph(para))) {
      // a mark without its own w:rFonts lays the line with the inherited ascii
      // face (Word probe 2026-09-05), not the document's CJK root factor; a
      // break-only paragraph lays its break line the same way (Word probe
      // 2026-10-20: ko-KR doc, Malgun eastAsia default, Calibri 11 break line
      // still fits at 14pt remaining). Empty cell lines take it from the
      // .doc-table .doc-p-empty rule (styles.css)
      styles.push('--doc-line-factor:var(--doc-line-factor-latin,1.2)')
    }
  }
  if (note) return { classes, styles }
  if (hf) {
    // strips sit outside the body grid (.page-hf kills the pitch) and inherit
    // .doc-page's computed line height, so a multiple stays a plain unitless
    // factor x m and every factor re-declares the line height on the paragraph
    // (Word sizes a strip line by its runs' size and face, probe 2026-10-29);
    // the paragraph box (indents, spacing, jc) is laid by hf-dom: tab stops
    // measure from the column edge and strip spacing merges per part
    const lh = hfParaLineHeightCss(fmt)
    if (lh) styles.push(`line-height:${lh}`)
    else if (styles.some((d) => d.startsWith('--doc-line-factor:')))
      styles.push(`line-height:${HF_FACE_LINE_HEIGHT}`)
  }
  const ownLine = fmt.lineRule != null || fmt.lineRawTwips != null || !!fmt.lineSpacing
  const tblLine = !ownLine ? fmt.tableStyleLine : undefined
  const lineRule = fmt.lineRule ?? tblLine?.rule ?? undefined
  const lineRawTwips =
    fmt.lineRawTwips != null
      ? fmt.lineRawTwips
      : tblLine && tblLine.rule !== 'auto'
        ? tblLine.raw
        : undefined
  const lineSpacing = fmt.lineSpacing
    ? fmt.lineSpacing
    : tblLine?.rule === 'auto'
      ? Math.round((tblLine.raw / 240) * 100) / 100
      : undefined
  const lh = cssLineHeight(lineRule, lineRawTwips, lineSpacing)
  // a picture-only line keeps its zero strut under an auto/atLeast rule (the
  // multiple's extra is padding, .doc-p-pic); an exact rule still pins it.
  // Read-only cells pin it inline (.doc-p-picture has no stylesheet twin)
  const picLine = fmt.picOnly && lineRule !== 'exact' && lineRule !== 'atLeast'
  if (kind === 'cell' && fmt.picOnly) styles.push('line-height:0')
  else if (lh && !picLine && !hf) styles.push(`line-height:${lh}`)
  const lhCap = cssExactLineCap(lineRule, lineRawTwips)
  if (lhCap) styles.push(`--doc-lh-cap:${lhCap}`)
  // a direct atLeast must not cap its runs at an inherited style-level exact rule
  else if (lineRule === 'atLeast') styles.push('--doc-lh-cap:initial')
  const leadTop = cssLeadTop(lineRule, lineRawTwips, lineSpacing)
  if (leadTop) styles.push(`--doc-lead-top:${leadTop}`)
  // grid-doc span snapping (doc-style-css): fixed-height lines opt out (atLeast
  // never snaps, including the line="0" opt-out form), multiples scale.
  // Cells stay off the typed grid (gridBlocks excludes .doc-cell-para)
  const fixedLh = (lineRule === 'exact' && lineRawTwips) || lineRule === 'atLeast'
  if (fixedLh) {
    if (kind !== 'cell') classes.push('doc-lh-fixed')
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
  if (fmt.snapToGrid === false && kind !== 'cell' && !hf) {
    styles.push('--doc-grid-pitch:0.0001px')
    classes.push('doc-nosnap')
  }
  // list items already indent via padding; a margin would double-shift them.
  // listGeometry: drive list geometry with w:ind (--li-left text indent, --li-hang the hanging
  // area i.e. the number-marker width); negative text-indent is expressed by the marker box, no longer emitted directly
  // Logical (inline-start/end) margins: identical to left/right in LTR, and in bidi
  // paragraphs they mirror, matching Word's quirk that w:ind left/right swap sides.
  // explicit 0 still emits: it beats the style indent, the 0.55in list default
  // and the numbering level (Word merges w:ind per attribute, direct value first)
  const includeIndent = ctx.includeIndent ?? true
  const listGeometry = ctx.listGeometry ?? false
  if (includeIndent && fmt.indentLeft != null) {
    styles.push(`margin-inline-start:${fmt.indentLeft / 20}pt`)
    if (ctx.anchorHost) styles.push(`--anchor-ind:${fmt.indentLeft / 20}pt`)
  } else if (listGeometry && fmt.indentLeft != null) {
    const left = fmt.indentLeft / 20
    // padding cannot go negative: a negative start indent moves the box instead
    if (left < 0) styles.push(`margin-inline-start:${left}pt`, '--li-left:0pt')
    else styles.push(`--li-left:${left}pt`)
    if (ctx.anchorHost) styles.push(`--anchor-${left < 0 ? 'ind' : 'pad'}:${left}pt`)
  }
  if (fmt.indentRight != null) styles.push(`margin-inline-end:${fmt.indentRight / 20}pt`)
  if (fmt.indentFirstLine != null) {
    const firstLine = fmt.indentFirstLine
    // the character-unit indent renders in em so it rescales with the first
    // run's font size (w:firstLineChars semantics), not frozen at the twips twin
    const chars = Number(fmt.indentFirstLineChars)
    const indent =
      chars > 0 ? `text-indent:calc(${chars / 100}em)` : `text-indent:${firstLine / 20}pt`
    if (listGeometry && firstLine < 0) styles.push(`--li-hang:${-firstLine / 20}pt`)
    else if (!listGeometry || firstLine > 0) styles.push(indent)
  }
  // explicit 0 must still emit (w:after="0" overrides the style/docDefaults margin)
  // autospacing replaces the literal with Word's HTML auto value (14pt, measured);
  // the sp-auto-* classes let CSS collapse it to 0 between two list items (Word)
  if (fmt.spaceBeforeAuto) {
    styles.push(`${spaceBeforeProp()}:${cssGridSpacingPt(autoSpacingPt())}`)
    classes.push('sp-auto-b')
  } else if (fmt.spaceBefore != null)
    styles.push(`${spaceBeforeProp()}:${cssGridSpacingPt(fmt.spaceBefore / 20)}`)
  if (fmt.spaceAfterAuto) {
    styles.push(`margin-bottom:${cssGridSpacingPt(autoSpacingPt())}`)
    classes.push('sp-auto-a')
  } else if (fmt.spaceAfter != null)
    styles.push(`margin-bottom:${cssGridSpacingPt(fmt.spaceAfter / 20)}`)
  // direct w:contextualSpacing: same-style adjacency suppression / explicit opt-out
  // (the style-level rules live in doc-style-css and honor these classes; they
  // target body p/.doc-li only)
  if (kind === 'body') {
    if (fmt.contextualSpacing === true) classes.push('ctx-sp')
    else if (fmt.contextualSpacing === false) classes.push('ctx-sp-off')
  }
  return { classes, styles }
}

/** inline leaves of an engine paragraph: a run's text, then its picture (the
 *  order runSpansWithPads renders); `null` runs = a plain-string cell paragraph */
export function layoutParaOfRuns(
  runs: Run[] | null,
  text: string,
  lineRule: LayoutPara['lineRule'],
): LayoutPara {
  if (!runs) {
    const run: LayoutRun = {
      kind: 'text',
      text,
      attrs: null,
      bare: true,
      bold: false,
      italic: false,
      pageOrColBreak: false,
    }
    return { text, runs: text ? [run] : [], childCount: text ? 1 : 0, lineRule }
  }
  const out: LayoutRun[] = []
  for (const run of runs) {
    if (run.text !== '') {
      out.push({
        kind: 'text',
        text: run.text,
        attrs: {
          font: run.font ?? null,
          fontAscii: run.fontAscii ?? null,
          eastAsiaFont: run.eastAsiaFont ?? null,
          // the cs chain sizes complex-script text, and every character of an rtl run (convert.ts)
          csFont: run.csFont && (run.cs || textHasComplexScript(run.text)) ? run.csFont : null,
          cs: run.cs ?? null,
          sizeHalfPoints: run.sizeHalfPoints ?? null,
          eaSlotEmpty: run.eaSlotEmpty ?? null,
        },
        bare: false,
        bold: !!run.bold,
        italic: !!run.italic,
        pageOrColBreak: false,
      })
    }
    if (run.image) {
      out.push({
        kind: 'image',
        text: '',
        attrs: null,
        bare: false,
        bold: false,
        italic: false,
        pageOrColBreak: false,
      })
    }
  }
  return { text, runs: out, childCount: out.length, lineRule }
}

/** RTL paragraph: own or style-inherited w:bidi, else inferred from the runs (convert.ts bidiInferred) */
export function paraRtl(fmt: ParaFormat | undefined, runs: Run[] | null): boolean {
  return !!(fmt?.bidi || fmt?.bidiInherited || inferredBidi(fmt, runs ?? undefined))
}

/**
 * Word orders a run without w:rtl left-to-right even inside a w:bidi paragraph
 * (probe 2026-09-16): an RTL paragraph whose runs are none of them rtl/cs
 * isolates its text as one LTR item (.doc-ltr-runs), like the editor's
 * paraContentSpec.
 */
export function wrapsLtrRuns(fmt: ParaFormat | undefined, runs: Run[] | null): boolean {
  return paraRtl(fmt, runs) && !!runs?.length && !runs.some((r) => r.cs || r.rtl)
}

/**
 * Read-only cell / textbox paragraph (engine ParaFormat + Run[]): the same line
 * box as the editor's blockAttrs, through lineBoxSpec. The table-style layer
 * (effective.source) sits between the stylesheet and direct formatting and is
 * merged here like the editor's tblPara attr.
 */
export function fromRuns(
  fmt: ParaFormat | undefined,
  runs: Run[] | null,
  text: string,
  kind: 'cell' | 'textbox',
): LineBoxSpec {
  const eff = kind === 'cell' ? fmt?.effective : undefined
  const tableStyle = (k: 'spaceBefore' | 'spaceAfter') =>
    eff?.source[k] === 'tableStyle'
      ? eff.spacing[k === 'spaceBefore' ? 'before' : 'after']
      : undefined
  const picOnly = !text && !!runs?.some((r) => r.image && !r.image.rule)
  return lineBoxSpec(
    layoutParaOfRuns(runs, text, fmt?.lineRule ?? null),
    {
      align: fmt?.align ?? null,
      bidi: paraRtl(fmt, runs),
      autoSpace: fmt?.autoSpace,
      wordWrap: fmt?.wordWrap ?? null,
      lineRule: fmt?.lineRule ?? null,
      lineRawTwips: fmt?.lineRawTwips ?? null,
      lineSpacing: fmt?.lineSpacing ?? null,
      tableStyleLine: eff?.source.line === 'tableStyle' ? (eff.spacing.line ?? null) : null,
      snapToGrid: fmt?.snapToGrid ?? null,
      indentLeft: fmt?.indentLeft ?? null,
      indentRight: fmt?.indentRight ?? null,
      indentFirstLine: fmt?.indentFirstLine ?? null,
      indentFirstLineChars: fmt?.directCharIndents?.firstLine ?? null,
      spaceBefore: fmt?.spaceBefore ?? tableStyle('spaceBefore') ?? null,
      spaceAfter: fmt?.spaceAfter ?? tableStyle('spaceAfter') ?? null,
      spaceBeforeAuto: fmt?.spaceBeforeAuto ?? false,
      spaceAfterAuto: fmt?.spaceAfterAuto ?? false,
      contextualSpacing: fmt?.contextualSpacing ?? null,
      emptyRunSize: fmt?.emptyRunSizeHalfPoints ?? null,
      emptyRunFont: fmt?.emptyRunFontFamily ?? null,
      picOnly,
    },
    { kind },
  )
}

export const HF_FACE_LINE_HEIGHT = 'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))'

/** strip paragraph line-height from its (style-resolved) w:spacing; strips sit
 *  outside the body grid, so multiples stay plain unitless factors */
export function hfParaLineHeightCss(para: {
  lineRule?: 'auto' | 'atLeast' | 'exact' | null
  lineRawTwips?: number | null
  lineSpacing?: number | null
}): string | null {
  if (para.lineRule === 'exact' || para.lineRule === 'atLeast') {
    return cssLineHeight(
      para.lineRule,
      para.lineRawTwips ?? undefined,
      para.lineSpacing ?? undefined,
    )
  }
  const m =
    para.lineSpacing ??
    (para.lineRule === 'auto' && para.lineRawTwips ? para.lineRawTwips / 240 : undefined)
  return m ? `calc(var(--doc-line-factor,1.2) * ${m})` : null
}

/** Display face of a strip run: its direct w:rPr, else what the paragraph /
 *  table style settles (a footer style declaring Malgun Gothic 10pt sizes runs
 *  without their own w:sz; the strip default is only for the style-less rest).
 *  Direct fields stay untouched: they are what saves. */
export function hfRunFace(run: Run): {
  sizeHalfPoints?: number
  fontAscii?: string
  font?: string
  eastAsiaFont?: string
  bold?: boolean
  italic?: boolean
} {
  const eff = run.effective
  // a w:rtl run of only neutral characters (page number, dashes, spaces) is
  // not sized by the style's szCs: Word draws a footer "- 9 -" under Normal
  // sz=32 szCs=22 at 16pt (measured 2026-10-28), so the strip default stands
  const neutralCs = eff?.cs === true && !/\p{L}/u.test(run.text ?? '')
  const fromStyle = (key: string) =>
    !neutralCs && (eff?.source[key] === 'paraStyle' || eff?.source[key] === 'tableStyle')
  const sizeHalfPoints =
    run.sizeHalfPoints ?? (fromStyle('sizeHalfPoints') ? eff?.sizeHalfPoints : undefined)
  const fontAscii = run.fontAscii ?? (fromStyle('ascii') ? eff?.fonts.ascii : undefined)
  const font =
    run.font ??
    (run.fontAscii ? undefined : fromStyle('eastAsia') ? eff?.fonts.eastAsia : undefined)
  // a style declaring both slots is not an ascii-only CJK face: without the
  // East Asian slot the line factor falls back to the script default
  const eastAsiaFont = run.eastAsiaFont ?? (fromStyle('eastAsia') ? eff?.fonts.eastAsia : undefined)
  const bold = run.bold ?? (fromStyle('bold') ? eff?.bold : undefined)
  const italic = run.italic ?? (fromStyle('italic') ? eff?.italic : undefined)
  return {
    ...(sizeHalfPoints ? { sizeHalfPoints } : {}),
    ...(fontAscii ? { fontAscii } : {}),
    ...(font ? { font } : {}),
    ...(eastAsiaFont ? { eastAsiaFont } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
  }
}

/** strip paragraph fields the line box reads: a stacked HfParagraph, or a
 *  layout-table cell paragraph's HfCellParaProps (whose sizeHalfPoints is the style size) */
export interface HfLayoutParagraph {
  runs?: Run[]
  lineRule?: 'auto' | 'atLeast' | 'exact'
  lineRawTwips?: number
  lineSpacing?: number
  bidi?: boolean
  wordWrap?: boolean
  emptyRunFont?: string
  emptyRunSizeHalfPoints?: number
  sizeHalfPoints?: number
}

/** the strip runs with their style-resolved display face (hfRunFace) */
export function hfFacedRuns(runs: readonly Run[]): Run[] {
  return runs.map((r) => ({ ...r, ...hfRunFace(r) }))
}

/**
 * Header/footer strip paragraph (stacked or layout-table cell line): the same
 * line box as the body, through lineBoxSpec with kind 'hf'. hfRunFace is
 * applied once here, so the strut and factor read the style-settled faces.
 */
export function fromHfParagraph(
  para: HfLayoutParagraph,
  runs: readonly Run[] | undefined = para.runs,
): LineBoxSpec {
  const faced = hfFacedRuns(runs ?? [])
  const text = faced.map((r) => r.text).join('')
  return lineBoxSpec(
    layoutParaOfRuns(faced, text, para.lineRule ?? null),
    {
      bidi: !!para.bidi,
      wordWrap: para.wordWrap ?? null,
      lineRule: para.lineRule ?? null,
      lineRawTwips: para.lineRawTwips ?? null,
      lineSpacing: para.lineSpacing ?? null,
      emptyRunSize: para.emptyRunSizeHalfPoints ?? null,
      emptyRunFont: para.emptyRunFont ?? null,
      picOnly: false,
    },
    { kind: 'hf', styleSizeHalfPoints: para.sizeHalfPoints ?? null },
  )
}

/**
 * Footnote / endnote entry (NoteRun[] of its first paragraph): direction,
 * Korean wrapping, strut face and face classes through lineBoxSpec with kind
 * 'note'; the row's line height and base size stay with the reservation
 * model (note-dom noteRowMetrics), which the row declares itself.
 */
export function fromNoteRuns(runs: readonly NoteRun[]): LineBoxSpec {
  const text = runs.map((r) => r.text).join('')
  return lineBoxSpec(
    layoutParaOfRuns(runs as Run[], text, null),
    { picOnly: false },
    { kind: 'note' },
  )
}

/** the strip paragraph's line factor as a number when every ink run declares
 *  its face (tab-laid lines floor at size x factor), null when a run inherits */
export function hfLineFactorNumber(runs: readonly Run[]): number | null {
  const faced = hfFacedRuns(runs).filter((r) => r.text.trim() !== '')
  const text = faced.map((r) => r.text).join('')
  const n = Number(paraLineFactor(layoutParaOfRuns(faced, text, null)))
  return Number.isFinite(n) && n > 0 ? n : null
}

/** camelCase style record (React style objects, hf-dom assignStyle); custom properties keep their name */
export function toStyleRecord(spec: LineBoxSpec): Record<string, string> {
  const out: Record<string, string> = {}
  for (const decl of spec.styles) {
    const i = decl.indexOf(':')
    const prop = decl.slice(0, i)
    out[
      prop.startsWith('--') ? prop : prop.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase())
    ] = decl.slice(i + 1)
  }
  return out
}

export function fromPmNode(
  node: ParaLayoutNode,
  { includeIndent = true, listGeometry = false }: ParaLayoutOptions = {},
): ParaLayoutSpec {
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
  // in-line pictures and nothing else: Word sizes the line to the picture
  // extent, no strut or descent under it (probe 2026-10-13: a 19.5pt icon in a
  // tcMar 40 cell makes a 23.52pt row; a text run beside it adds its descent)
  const picOnly = !node.textContent && (node.childCount ?? 0) > 0 && isPictureOnlyContent(node)
  if (picOnly) classes.push('doc-p-pic')
  // the picture's containing block is already inset by the start indent, which
  // a column-relative X subtracts again (--anchor-ind / --anchor-pad)
  const anchorHost = hostsColumnAnchoredPicture(node)
  if (anchorHost) classes.push('doc-anchor-host')
  // a band-offset float needs its host to contain it (styles.css flow-root)
  if (hostsBandOffsetPicture(node)) classes.push('doc-band-host')
  // a column break leading a paragraph with text has no line of its own: Word
  // opens the next column with the first text line (CSS hides the break)
  const lead = node.firstChild
  if (lead?.type.name === 'hardBreak' && lead.attrs.colBreak && node.textContent)
    classes.push('doc-lead-col-br')
  if (node.attrs.pageBreakBefore) {
    classes.push('page-break-before')
    attrs['data-page-break-label'] = t('editorPageBreak')
  }
  if (classes.length > 0) attrs['class'] = classes.join(' ')
  const tbl = tblParaOf(node.attrs)
  const box = lineBoxSpec(
    layoutParaOf(node),
    {
      align: (node.attrs.align as string | null) ?? tbl?.align ?? null,
      bidi: !!node.attrs.bidi,
      bidiInferred: !!node.attrs.bidiInferred,
      autoSpace: node.attrs.autoSpace,
      wordWrap: node.attrs.wordWrap as boolean | null,
      lineRule: node.attrs.lineRule as LayoutParaFormat['lineRule'],
      lineRawTwips: node.attrs.lineRawTwips != null ? Number(node.attrs.lineRawTwips) : null,
      lineSpacing: node.attrs.lineSpacing ? Number(node.attrs.lineSpacing) : null,
      tableStyleLine: tbl?.line ?? null,
      snapToGrid: node.attrs.snapToGrid as boolean | null,
      indentLeft: numOrNull(node.attrs.indentLeft),
      indentRight: numOrNull(node.attrs.indentRight),
      indentFirstLine: numOrNull(node.attrs.indentFirstLine),
      indentFirstLineChars: numOrNull(node.attrs.indentFirstLineChars),
      spaceBefore: numOrNull(node.attrs.spaceBefore ?? tbl?.spaceBefore),
      spaceAfter: numOrNull(node.attrs.spaceAfter ?? tbl?.spaceAfter),
      spaceBeforeAuto: !!node.attrs.spaceBeforeAuto,
      spaceAfterAuto: !!node.attrs.spaceAfterAuto,
      contextualSpacing: node.attrs.contextualSpacing as boolean | null,
      emptyRunSize: numOrNull(node.attrs.emptyRunSize),
      emptyRunFont: node.attrs.emptyRunFont ? String(node.attrs.emptyRunFont) : null,
      picOnly,
    },
    { kind: 'body', includeIndent, listGeometry, anchorHost },
  )
  classes.push(...box.classes)
  const styles = box.styles
  if (classes.length > 0) attrs['class'] = classes.join(' ')
  const shdBg = node.attrs.shadingDisplay ?? node.attrs.shadingFill
  // authored colors stay the declaration; the --dk-* twins feed the dark page (dark-page.ts)
  if (shdBg)
    styles.push(
      `background-color:#${shdBg}`,
      dkBackground(`#${shdBg}`),
      `--pbdr-fill:#${shdBg}`,
      `box-shadow:${PARA_BORDER_SHADOW}`,
    )
  else if (node.attrs.shadingClear)
    styles.push('background-color:transparent', '--pbdr-fill:transparent')
  const ink = fillInk(shdBg ? String(shdBg) : null)
  if (ink) attrs['data-ink'] = ink
  const borderPad = parseJsonAttr<Partial<Record<'t' | 'b' | 'l' | 'r', number>>>(
    node.attrs.borderPad,
  )
  if (node.attrs.borderReset) {
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const key = DK_SIDE[side]
      if (!String(node.attrs.borderReset).includes(key)) continue
      styles.push(`border-${side}:none`)
      // list items indent through padding-inline-start, so their reset l/r sides keep it
      if ((!listGeometry || side === 'top' || side === 'bottom') && !borderPad?.[key])
        styles.push(`padding-${side}:0`)
    }
  }
  if (node.attrs.borders || borderPad) {
    const borders = String(node.attrs.borders ?? '')
    const borderLines = paraBorderLinesWithPad(
      parseJsonAttr<Partial<Record<'t' | 'b' | 'l' | 'r', { color?: string; szPt?: number }>>>(
        node.attrs.borderLines,
      ),
      borderPad,
    )
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const key = DK_SIDE[side]
      if (!borders.includes(key)) continue
      const line = paraBorderCss(borderLines[key] ?? undefined)
      styles.push(`border-${side}:${line}`, dkBorder(key, line))
    }
    const padding = paraBorderPadding(borders, borderLines)
    // list items indent through padding-inline-start; a border must not replace it
    if (listGeometry) {
      delete padding.paddingLeft
      delete padding.paddingRight
    }
    styles.push(...paraBorderPaddingDecls(padding), ...paraBorderShadowDecls(borderLines))
  }
  // a positioned frame (below) models the same w:framePr; frameBox only covers frames without a width
  if (node.attrs.frameBox && !node.attrs.frame)
    styles.push(...frameBoxCss(String(node.attrs.frameBox)))
  if (node.attrs.tabStops) {
    // debugging aid only; actual tab layout is measured by TabStopExtension
    attrs['data-tab-stops'] = String(node.attrs.tabStops)
  }
  if (node.attrs.dropCap) {
    attrs['data-drop-cap'] = String(node.attrs.dropCap)
  }
  if (node.attrs.frame) {
    try {
      const frame = JSON.parse(String(node.attrs.frame)) as ParaFrame
      if (frame && frame.wTwips > 0) {
        const fc = paraFrameCss(frame, node.attrs.textDirection as TextFlowDirection | null)
        classes.push(...fc.classes)
        attrs['class'] = classes.join(' ')
        styles.push(...fc.styles)
      }
    } catch {
      /* ignore malformed */
    }
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
  return { classes, styles, attrs }
}

function parseJsonAttr<T>(raw: unknown): T | undefined {
  if (!raw) return undefined
  try {
    return JSON.parse(String(raw)) as T
  } catch {
    return undefined
  }
}

/** w:framePr box: the paragraph lays out at the frame size, floating beside the text when it wraps */
function frameBoxCss(json: string): string[] {
  let box: ParaFrameBox
  try {
    box = JSON.parse(json) as ParaFrameBox
  } catch {
    return []
  }
  const px = (twips: number): string => `${(twips / 15).toFixed(1)}px`
  const out: string[] = []
  if (box.wTwips) out.push(`width:${px(box.wTwips)}`, 'box-sizing:border-box')
  if (box.hTwips) {
    if (box.hRule === 'exact') out.push(`height:${px(box.hTwips)}`, 'overflow:hidden')
    else out.push(`min-height:${px(box.hTwips)}`)
  }
  if (box.floatSide) {
    out.push(`float:${box.floatSide}`)
    if (box.hSpaceTwips)
      out.push(`margin-${box.floatSide === 'left' ? 'right' : 'left'}:${px(box.hSpaceTwips)}`)
    if (box.vSpaceTwips) out.push(`margin-bottom:${px(box.vSpaceTwips)}`)
  }
  return out
}
