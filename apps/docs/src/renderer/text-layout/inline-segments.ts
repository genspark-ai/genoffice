/**
 * Character-level inline layout decisions of one paragraph, computed once and
 * consumed by every emitter (live ProseMirror decorations, the static DOM of
 * protected tables / textboxes, the snapshot tool). Each rule is a pure
 * function over the paragraph's runs; the emitters only map a Segment to a
 * class / style span and a joiner to a word-joiner widget.
 *
 * Rules here (Word truths in each function's comment):
 *   autospace pads, balanced spaces, Korean no-break seams, SimSun-gap line
 *   lifts, justification symbols, hinted East Asian curly quotes.
 */
import { textHasEastAsian } from '@genoffice/docx-engine'
import {
  asciiOnlyCjkFace,
  autospaceLatinFirst,
  autospacePadBetween,
  autospacePadClass,
  autospacePadScale,
  autospaceSeams,
  balancedSpaceOffsets,
  codePointLengthAt,
  hasBalancedSpaceContext,
  justifySymbolRanges,
  krNoBreakOffsets,
  latinAutospaceGapEm,
  SIMSUN_GAP_CHAR_RE,
  simsunGapLineFactor,
  textHasHangul,
  type AutospaceMode,
} from '../line-metrics'

export interface InlineRunAttrs {
  font?: string | null
  fontAscii?: string | null
  eastAsiaFont?: string | null
  eaSlotEmpty?: boolean | null
  rawRPr?: string | null
}

export interface InlineRun {
  /** '' for a non-text inline (picture, break, field): it breaks adjacency */
  text: string
  /** content offset of the run start */
  offset: number
  /** content size (text length, 1 for an inline node) */
  size: number
  attrs: InlineRunAttrs | null
}

export interface InlineParagraph {
  runs: InlineRun[]
  /** autoSpaceDE/DN pad mode; null = pads off */
  autospace: AutospaceMode | null
  /** ascii face a run without its own font inherits */
  inheritedAscii: string | null
  /** East Asian face a run without its own font inherits */
  inheritedEa: string | null
  /** w:wordWrap (Korean word-wrap keeps hangul/punctuation seams together) */
  wordWrap: boolean
  /** balanceSingleByteDoubleByteWidth widening is on for the document */
  balancedSpaces: boolean
  /** justified / distributed paragraph (the symbol stretch rule applies when it holds no CJK text) */
  justified: boolean
  /** line-height expression the SimSun-gap lift applies, null when the lift is off (exact lines) */
  simsunGapLineHeight: string | null
}

export interface Segment {
  from: number
  to: number
  cls?: string
  style?: string
}

export interface InlineSegments {
  segments: Segment[]
  /** content offsets of U+2060 word-joiner widgets */
  joiners: number[]
}

const textOf = (run: InlineRun | undefined) => run?.text ?? ''

/** Word's autoSpaceDE/DN gap: ~1/4em between a CJK letter and a Latin letter/digit, never
 *  beside spaces or punctuation; the Latin → CJK side scales by the Latin face's xAvgCharWidth
 *  (probes 2026-10-14, 2026-10-20) */
function autospacePads(p: InlineParagraph, mode: AutospaceMode): Segment[] {
  const out: Segment[] = []
  const cls = autospacePadClass(mode)
  const faceOf = (run: InlineRun): string | null =>
    (run.attrs?.fontAscii ?? run.attrs?.font) || p.inheritedAscii
  const push = (from: number, to: number, latinFirst: boolean, face: string | null) => {
    const scale = latinFirst ? autospacePadScale(latinAutospaceGapEm(face), mode) : 1
    out.push(
      scale === 1
        ? { from, to, cls }
        : {
            from,
            to,
            cls: `${cls} doc-autospace-pad-le`,
            style: `--doc-autospace-pad-le:${scale}`,
          },
    )
  }
  let prevText = ''
  let prevFace: string | null = null
  for (const run of p.runs) {
    if (!run.text) {
      prevText = ''
      continue
    }
    const face = faceOf(run)
    if (autospacePadBetween(prevText, run.text, mode)) {
      push(
        run.offset,
        run.offset + codePointLengthAt(run.text, 0),
        autospaceLatinFirst(prevText),
        prevFace,
      )
    }
    for (const seam of autospaceSeams(run.text, mode)) {
      push(
        run.offset + seam.at,
        run.offset + seam.at + codePointLengthAt(run.text, seam.at),
        seam.latinFirst,
        face,
      )
    }
    prevText = run.text
    prevFace = face
  }
  return out
}

/** balanceSingleByteDoubleByteWidth: a space beside an East Asian character or another
 *  space advances 0.5em (probe 2026-11-04) */
function balancedSpaces(p: InlineParagraph): Segment[] {
  const out: Segment[] = []
  p.runs.forEach((run, k) => {
    if (!run.text) return
    for (const i of balancedSpaceOffsets(run.text, textOf(p.runs[k - 1]), textOf(p.runs[k + 1]))) {
      out.push({ from: run.offset + i, to: run.offset + i + 1, cls: 'doc-balanced-space' })
    }
  })
  return out
}

/** Korean word-wrap breaks only at spaces; Chromium's keep-all still splits hangul from
 *  adjoining punctuation, so those seams get a word joiner (probe 2026-10-20) */
function krNoBreakJoiners(p: InlineParagraph): number[] {
  const out: number[] = []
  let prevText = ''
  for (const run of p.runs) {
    if (run.text) {
      for (const i of krNoBreakOffsets(run.text, prevText)) out.push(run.offset + i)
      prevText = run.text
    } else prevText = ''
  }
  return out
}

/** ・ U+30FB / 〜 U+301C in a SimSun-substituted run lift the whole line to the YaHei
 *  factor (probe 2026-08-13); the glyphs come from the eastAsia slot, so an empty theme
 *  slot or an ascii-only CJK face does not count (same slot rule as paraLineFactor) */
function simsunGapLifts(p: InlineParagraph, lineHeight: string): Segment[] {
  const out: Segment[] = []
  for (const run of p.runs) {
    if (!run.text || !SIMSUN_GAP_CHAR_RE.test(run.text)) continue
    const a = run.attrs
    const own = a?.font ?? a?.fontAscii
    const family =
      a?.eaSlotEmpty === true || (a && asciiOnlyCjkFace(a)) ? null : own || p.inheritedEa
    if (!family || simsunGapLineFactor(family) === null) continue
    for (let i = 0; i < run.text.length; i++) {
      if (SIMSUN_GAP_CHAR_RE.test(run.text[i]))
        out.push({
          from: run.offset + i,
          to: run.offset + i + 1,
          style: `line-height:${lineHeight}`,
        })
    }
  }
  return out
}

/** symbols Chromium justifies like ideographs stay unstretched in non-CJK paragraphs
 *  (Word stretches only their spaces); CJK paragraphs keep Chromium's inter-ideograph
 *  distribution around them */
function justifySymbols(p: InlineParagraph): Segment[] {
  const out: Segment[] = []
  for (const run of p.runs) {
    if (!run.text) continue
    for (const r of justifySymbolRanges(run.text))
      out.push({ from: run.offset + r.from, to: run.offset + r.to, cls: 'doc-justify-symbol' })
  }
  return out
}

/** the run's own w:rFonts carries w:hint="eastAsia" (revision snapshots inside w:rPrChange don't count) */
const EA_HINT_RE = /<w:rFonts\b[^>]*\bw:hint="eastAsia"/
const EA_QUOTES_RE = /[‘’“”]+/g

/** quote ranges of `text` Word renders with the East Asian font (w:hint="eastAsia";
 *  surrounding text and w:lang play no part, Word for Mac probe 2026-09-01) */
export function eaHintQuoteRanges(
  rawRPr: string | null | undefined,
  text: string,
): Array<{ from: number; to: number }> {
  if (!rawRPr || !text || !/[‘’“”]/.test(text)) return []
  if (!EA_HINT_RE.test(rawRPr.replace(/<w:rPrChange[\s\S]*?<\/w:rPrChange>/, ''))) return []
  const ranges: Array<{ from: number; to: number }> = []
  for (const m of text.matchAll(EA_QUOTES_RE))
    ranges.push({ from: m.index, to: m.index + m[0].length })
  return ranges
}

function eaHintQuotes(p: InlineParagraph): Segment[] {
  const out: Segment[] = []
  for (const run of p.runs) {
    for (const r of eaHintQuoteRanges(run.attrs?.rawRPr, run.text))
      out.push({ from: run.offset + r.from, to: run.offset + r.to, cls: 'doc-ea-quotes' })
  }
  return out
}

export function inlineSegments(p: InlineParagraph): InlineSegments {
  const text = p.runs.map((r) => r.text).join('')
  const segments: Segment[] = []
  if (p.autospace) segments.push(...autospacePads(p, p.autospace))
  if (p.balancedSpaces && hasBalancedSpaceContext(text)) segments.push(...balancedSpaces(p))
  if (p.simsunGapLineHeight) segments.push(...simsunGapLifts(p, p.simsunGapLineHeight))
  if (p.justified && !textHasEastAsian(text)) segments.push(...justifySymbols(p))
  segments.push(...eaHintQuotes(p))
  const joiners = p.wordWrap && textHasHangul(text) ? krNoBreakJoiners(p) : []
  return { segments, joiners }
}

/** segments / joiners that fall inside one run, offsets relative to the run text */
export function segmentsOfRun(all: InlineSegments, run: InlineRun): InlineSegments {
  const end = run.offset + run.size
  return {
    segments: all.segments
      .filter((s) => s.from >= run.offset && s.to <= end)
      .map((s) => ({ ...s, from: s.from - run.offset, to: s.to - run.offset })),
    // a joiner on a run boundary belongs to the run that starts there
    joiners: all.joiners.filter((at) => at >= run.offset && at < end).map((at) => at - run.offset),
  }
}

export interface RunLike {
  text: string
  image?: unknown
  font?: string | null
  fontAscii?: string | null
  eastAsiaFont?: string | null
  eaSlotEmpty?: boolean | null
  rawRPr?: string | null
}

/** InlineRun list of parsed runs: a run's text precedes its picture, and the picture
 *  is a one-slot non-text inline that breaks adjacency */
export function inlineRunsOf(runs: readonly RunLike[]): { runs: InlineRun[]; byRun: InlineRun[] } {
  const out: InlineRun[] = []
  const byRun: InlineRun[] = []
  let offset = 0
  for (const r of runs) {
    const attrs =
      r.font || r.fontAscii || r.eastAsiaFont || r.rawRPr
        ? {
            font: r.font,
            fontAscii: r.fontAscii,
            eastAsiaFont: r.eastAsiaFont,
            eaSlotEmpty: r.eaSlotEmpty,
            rawRPr: r.rawRPr,
          }
        : null
    const run: InlineRun = { text: r.text, offset, size: r.text.length, attrs }
    byRun.push(run)
    if (r.text !== '') out.push(run)
    offset += r.text.length
    if (r.image) {
      out.push({ text: '', offset, size: 1, attrs: null })
      offset += 1
    }
  }
  return { runs: out, byRun }
}
