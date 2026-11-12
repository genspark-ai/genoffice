/**
 * Character-level inline layout decisions of one paragraph, computed once and
 * consumed by every emitter (live ProseMirror decorations, the static DOM of
 * protected tables / textboxes, the snapshot tool). Each rule is a pure
 * function over the paragraph's runs; the emitters only map a Segment to a
 * class / style span and a joiner to a word-joiner widget.
 *
 * Rules here (Word truths in each function's comment):
 *   autospace pads, balanced spaces, Korean no-break seams, SimSun-gap line
 *   lifts, justification symbols, hinted East Asian curly quotes, whitespace
 *   runs that give up their line box.
 */
import { textHasEastAsian } from '@genoffice/docx-engine'
import {
  asciiOnlyCjkFace,
  autospaceHangulSeam,
  autospaceLatinFirst,
  cssCsDigitLiftExpr,
  autospacePadBetween,
  autospacePadClass,
  autospacePadScale,
  autospaceSeams,
  balancedSpaceOffsets,
  codePointLengthAt,
  hasBalancedSpaceContext,
  isWhitespaceRun,
  justifySymbolRanges,
  fallbackGlyphRanges,
  symbolFallbackFontFamily,
  symbolFallbackLineFactor,
  krNoBreakOffsets,
  textHasComplexScript,
  textHasHangul,
  type AutospaceMode,
} from '../line-metrics'
import {
  cssCsFontFamily,
  cssRunFontFamily,
  latinAutospaceGapEm,
  lineHeightFactor,
  SIMSUN_GAP_CHAR_RE,
  simsunGapLineFactor,
} from './font-rules'
import { middleDotGlue, type MiddleDotGlue } from '../table-layout/tokenize'

export interface InlineRunAttrs {
  font?: string | null
  fontAscii?: string | null
  eastAsiaFont?: string | null
  eaSlotEmpty?: boolean | null
  rawRPr?: string | null
  /** w:cs face when it applies to the run (rtl run, or complex-script text) */
  csFont?: string | null
  /** the run's own w:lang w:eastAsia (direct or character style) */
  eaLang?: string | null
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
  /** exact line rule: fallback glyphs never lift their line */
  exactLine?: boolean
  /** the paragraph's own line-height declaration (fixed rules), carried by lifted glyph spans */
  fixedLineHeight?: string | null
  /** paragraph line multiple for per-line face lifts, null when lines are pinned (exact rule) */
  lineMultiple?: number | 'inherited' | null
  /** w:lang w:eastAsia a run without its own inherits (paragraph style chain, then docDefaults) */
  eaLang?: string | null
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
  const push = (
    from: number,
    to: number,
    latinFirst: boolean,
    face: string | null,
    hangul: boolean,
  ) => {
    const scale = latinFirst ? autospacePadScale(latinAutospaceGapEm(face)) : 2
    // a Korean word-wrap paragraph never breaks at a hangul seam, so its margin never meets a line
    // start; with wordWrap off the seam can open a line and takes the collapsible space like any other
    const base = hangul && p.wordWrap ? `${cls} doc-autospace-pad-m` : cls
    out.push(
      scale === 2
        ? { from, to, cls: base }
        : {
            from,
            to,
            cls: `${base} doc-autospace-pad-le`,
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
        autospaceHangulSeam(prevText, run.text),
      )
    }
    for (const seam of autospaceSeams(run.text, mode)) {
      push(
        run.offset + seam.at,
        run.offset + seam.at + codePointLengthAt(run.text, seam.at),
        seam.latinFirst,
        face,
        autospaceHangulSeam(run.text.slice(0, seam.at), run.text.slice(seam.at)),
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

const MIDDLE_DOT = '\u00b7'

/** Word's line-start / line-end character set reaches Latin text through the East Asian
 *  language: under zh-CN or no w:lang w:eastAsia the spaces beside U+00B7 and before ":"
 *  are no break opportunity, under zh-TW only the ones before (probe 2026-11-08). The
 *  span opens on the character before the spaces: Chromium breaks ahead of a no-wrap
 *  span that starts with a space, so a space run opening its run stays breakable. */
function middleDotNoBreaks(p: InlineParagraph): Segment[] {
  const chars: Array<{ ch: string; at: number; run: number; glue: MiddleDotGlue }> = []
  p.runs.forEach((run, k) => {
    if (!run.text) {
      chars.push({ ch: '', at: run.offset, run: k, glue: 'none' })
      return
    }
    const glue = middleDotGlue(run.attrs?.eaLang ?? p.eaLang ?? undefined)
    for (let i = 0; i < run.text.length; i++)
      chars.push({ ch: run.text[i], at: run.offset + i, run: k, glue })
  })
  const out: Segment[] = []
  for (let a = 1; a < chars.length; a++) {
    if (chars[a].ch !== ' ' || chars[a - 1].ch === ' ') continue
    let b = a
    while (b < chars.length && chars[b].ch === ' ') b++
    const prev = chars[a - 1]
    const next = chars[b]
    if (!prev.ch || !next?.ch) continue
    const glued =
      ((next.ch === MIDDLE_DOT || next.ch === ':') && next.glue !== 'none') ||
      (prev.ch === MIDDLE_DOT && prev.glue === 'both')
    if (!glued || prev.run !== chars[b - 1].run) continue
    const last = out[out.length - 1]
    if (last && last.to === prev.at && chars[a - 2]?.run === prev.run) last.to = chars[b - 1].at + 1
    else out.push({ from: prev.at, to: chars[b - 1].at + 1, cls: 'doc-nobreak' })
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
const RPR_CHANGE_RE = /<w:rPrChange[\s\S]*?<\/w:rPrChange>/

function hasEaHint(rawRPr: string | null | undefined): boolean {
  return !!rawRPr && EA_HINT_RE.test(rawRPr.replace(RPR_CHANGE_RE, ''))
}

/** symbols the run's Latin face lacks draw in Word's fallback face (Segoe UI Symbol,
 *  Apple Color Emoji, Arial) and the line takes that face's box where the glyph lands
 *  (probes 2026-11-07/08); a non-lifting span gets a zero inline box so the fallback
 *  face's own metrics never grow the line; an eastAsia-hinted run draws them with its
 *  East Asian face */
function symbolFallbacks(p: InlineParagraph): Segment[] {
  const out: Segment[] = []
  for (const run of p.runs) {
    if (!run.text || hasEaHint(run.attrs?.rawRPr)) continue
    const declared = run.attrs?.fontAscii ?? null
    for (const r of fallbackGlyphRanges(run.text, declared, declared ?? p.inheritedAscii)) {
      const factor = p.exactLine ? null : symbolFallbackLineFactor(r.kind)
      const style = [`font-family:${symbolFallbackFontFamily(r.kind)}`]
      if (factor !== null) {
        style.push(`--doc-line-factor:${factor}`)
        if (p.fixedLineHeight) style.push(`line-height:${p.fixedLineHeight}`)
      } else style.push('line-height:0')
      out.push({
        from: run.offset + r.from,
        to: run.offset + r.to,
        cls: factor !== null ? 'doc-run-lf doc-sym-fb' : 'doc-sym-fb',
        style: style.join(';'),
      })
    }
  }
  return out
}
const EA_QUOTES_RE = /[‘’“”]+/g

/** quote ranges of `text` Word renders with the East Asian font (w:hint="eastAsia";
 *  surrounding text and w:lang play no part, Word for Mac probe 2026-09-01) */
export function eaHintQuoteRanges(
  rawRPr: string | null | undefined,
  text: string,
): Array<{ from: number; to: number }> {
  if (!text || !/[‘’“”]/.test(text) || !hasEaHint(rawRPr)) return []
  const ranges: Array<{ from: number; to: number }> = []
  for (const m of text.matchAll(EA_QUOTES_RE))
    ranges.push({ from: m.index, to: m.index + m[0].length })
  return ranges
}

/** a space / tab run beside ink gives up its line box (styles.css .doc-ws-run): Word never
 *  sizes a line by whitespace, whatever the run's face or size, while Chromium unions the
 *  run's own box into the line (probe 2026-11-08: Calibri 11 / 14 space or tab runs leave a
 *  Times 12 line at 13.8; a no-break space run is ink and lifts it) */
function whitespaceRuns(p: InlineParagraph): Segment[] {
  const out: Segment[] = []
  if (!p.runs.some((r) => r.text && !isWhitespaceRun(r.text))) return out
  for (const run of p.runs) {
    if (run.text && isWhitespaceRun(run.text))
      out.push({ from: run.offset, to: run.offset + run.size, cls: 'doc-ws-run' })
  }
  return out
}

function eaHintQuotes(p: InlineParagraph): Segment[] {
  const out: Segment[] = []
  for (const run of p.runs) {
    for (const r of eaHintQuoteRanges(run.attrs?.rawRPr, run.text))
      out.push({ from: run.offset + r.from, to: run.offset + r.to, cls: 'doc-ea-quotes' })
  }
  return out
}

const ASCII_NUMBER_RE = /[0-9]+(?:[.,:/][0-9]+)*/g
// Latin-1 letters without the multiplication / division signs (neutrals under W7)
const LATIN_LETTER_RE = /[A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u024F\u0370-\u03FF\u0400-\u04FF]/
const STRONG_LETTER_RE =
  /[A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u024F\u0370-\u03FF\u0400-\u04FF\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/

/** ranges of `text` (one run, `before` = the paragraph text ahead of it) that Word
 *  draws with the ascii face inside a complex-script run: ASCII digit groups whose
 *  nearest preceding strong letter is not Latin (UAX#9 W7 keeps them European
 *  numbers; after a Latin letter they join the Latin stretch and stay in the cs
 *  face). Word for Mac probes 2026-11-02 / 2026-11-08: '1.2' and '2024' beside
 *  Arabic letters are Calibri under cs Arial with and without w:rtl / w:cs, and
 *  pace the line at Calibri's 1.22; 'abc 1.2 def 2024' is all Arial; '.' ':' '(' '%'
 *  without digits and Arabic-Indic digits stay in the cs face. */
export function csAsciiDigitRanges(text: string, before = ''): Array<{ from: number; to: number }> {
  const out: Array<{ from: number; to: number }> = []
  if (!/[0-9]/.test(text)) return out
  for (const m of text.matchAll(ASCII_NUMBER_RE)) {
    const ahead = before + text.slice(0, m.index)
    let strong: string | undefined
    for (let i = ahead.length - 1; i >= 0; i--) {
      if (STRONG_LETTER_RE.test(ahead[i])) {
        strong = ahead[i]
        break
      }
    }
    if (strong && LATIN_LETTER_RE.test(strong)) continue
    out.push({ from: m.index, to: m.index + m[0].length })
  }
  return out
}

/** digit groups of a cs run move to the ascii face at the run size and lift their own
 *  line to that face's factor when it is the taller one */
function csAsciiDigits(p: InlineParagraph): Segment[] {
  const out: Segment[] = []
  let before = ''
  for (const run of p.runs) {
    const cs = run.attrs?.csFont
    const ascii = (run.attrs?.fontAscii ?? run.attrs?.font) || p.inheritedAscii
    if (cs && ascii && ascii.trim().toLowerCase() !== cs.trim().toLowerCase()) {
      const family = cssRunFontFamily(ascii, null)
      const lift =
        p.lineMultiple != null && lineHeightFactor(ascii) > lineHeightFactor(cs) + 1e-3
          ? `;line-height:${cssCsDigitLiftExpr(
              family,
              cssCsFontFamily(cs, ascii),
              lineHeightFactor(ascii),
              lineHeightFactor(cs),
              p.lineMultiple,
            )}`
          : ''
      for (const r of csAsciiDigitRanges(run.text, before))
        out.push({
          from: run.offset + r.from,
          to: run.offset + r.to,
          cls: 'doc-cs-ascii',
          style: `font-family:${family}${lift}`,
        })
    }
    before += run.text
  }
  return out
}

/** `added` with every range already covered by `taken` cut out (a pad on the first
 *  digit of a CJK-digit seam leaves the rest of the number its own span) */
function subtractOverlaps(added: Segment[], taken: Segment[]): Segment[] {
  const out: Segment[] = []
  for (const a of added) {
    let pieces = [a]
    for (const t of taken) {
      pieces = pieces.flatMap((x) => {
        if (t.from >= x.to || t.to <= x.from) return [x]
        const left = t.from > x.from ? [{ ...x, to: t.from }] : []
        const right = t.to < x.to ? [{ ...x, from: t.to }] : []
        return [...left, ...right]
      })
    }
    out.push(...pieces)
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
  segments.push(...symbolFallbacks(p))
  // a range another rule already wraps (autospace pad, balanced space) keeps that
  // span: the static emitter writes each segment's text once, overlaps would double it
  segments.push(...subtractOverlaps(csAsciiDigits(p), segments))
  if (text.includes(MIDDLE_DOT) || text.includes(' :')) {
    const whole = middleDotNoBreaks(p)
    // a span another rule holds inside a glued range (balanced space) takes the wrap
    // mode itself: segments cannot nest
    for (const s of segments) {
      if (whole.some((w) => w.from <= s.from && s.to <= w.to))
        s.cls = s.cls ? `${s.cls} doc-nobreak` : 'doc-nobreak'
    }
    segments.push(...subtractOverlaps(whole, segments))
  }
  segments.push(
    ...whitespaceRuns(p).filter((w) => !segments.some((s) => s.from < w.to && w.from < s.to)),
  )
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
  csFont?: string | null
  cs?: boolean | null
  eastAsiaLang?: string | null
}

/** InlineRun list of parsed runs: a run's text precedes its picture, and the picture
 *  is a one-slot non-text inline that breaks adjacency */
export function inlineRunsOf(runs: readonly RunLike[]): { runs: InlineRun[]; byRun: InlineRun[] } {
  const out: InlineRun[] = []
  const byRun: InlineRun[] = []
  let offset = 0
  for (const r of runs) {
    // same gate as the static run span: the cs face applies to rtl runs and complex-script text
    const csFont = r.csFont && (r.cs || textHasComplexScript(r.text)) ? r.csFont : null
    const attrs =
      r.font || r.fontAscii || r.eastAsiaFont || r.rawRPr || csFont || r.eastAsiaLang
        ? {
            font: r.font,
            fontAscii: r.fontAscii,
            eastAsiaFont: r.eastAsiaFont,
            eaSlotEmpty: r.eaSlotEmpty,
            rawRPr: r.rawRPr,
            csFont,
            eaLang: r.eastAsiaLang,
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
