/**
 * Phase 3.3 element-level patch generation — regenerating OOXML fragments for
 * dirty elements.
 *
 * Fidelity philosophy (aligned with docx-engine): **no wholesale parse→serialize**
 * (it loses unmodeled attributes: effects/custom geometry/extensions etc.).
 * Instead, **surgical in-place patches** of the original <p:sp> byte slice:
 *   - Text editing (the main scenario): replace each <a:r>'s <a:t> text content in
 *     order, and patch/inject <a:rPr> formatting attributes as needed
 *     (b/i/u/sz + solidFill color).
 *   - Untouched bytes (bodyPr/spPr/geometry/placeholder attributes…) are kept verbatim.
 *
 * If the model's run count doesn't match the original XML's <a:r> count (runs
 * added/removed / paragraph structure changed), fall back to the
 * "rebuild txBody per paragraph" path (still keeping outer bytes like spPr where
 * possible). The Phase 3 editor uses the former (lossless) when only text and
 * formatting change without structural edits; structural changes use the latter.
 */
import type { SlideElement, TextElement, Paragraph, TextRun, Transform, PPrDirty } from './types'
import { escapeXmlText, escapeXmlAttr } from './xml-utils'

/**
 * In-place patch of a text/shape element's original XML.
 * originalXml = the <p:sp>'s full original byte slice (anchor.originalXml).
 */
export function patchTextElementXml(el: TextElement, originalXml: string): string {
  if (!el.text || !el.text.paragraphs.length) return originalXml

  // Collect all model runs (flattened across paragraphs, incl. "\n" soft-break sentinels),
  // one-to-one with the original XML's <a:r>/<a:br> sequence (parse rewrites <a:br/>
  // as a sentinel run, order preserved)
  const modelRuns: TextRun[] = el.text.paragraphs.flatMap((p) => p.runs)
  const runSpans = findRunSpans(originalXml)

  // Count and kind aligned position by position (sentinel ⇔ <a:br>) → lossless in-place patch of each run, br bytes untouched
  const aligned =
    runSpans.length === modelRuns.length &&
    runSpans.length > 0 &&
    runSpans.every((s, i) => (s.kind === 'br') === isSoftBreakRun(modelRuns[i]!))
  if (aligned) {
    let out = ''
    let cursor = 0
    for (let i = 0; i < runSpans.length; i++) {
      const span = runSpans[i]!
      out += originalXml.slice(cursor, span.start)
      const slice = originalXml.slice(span.start, span.end)
      out += span.kind === 'br' ? slice : patchRun(slice, modelRuns[i]!)
      cursor = span.end
    }
    out += originalXml.slice(cursor)
    return out
  }

  // Structural change → rebuild the <p:txBody> content, keeping the txBody wrapper and bodyPr
  return rebuildTxBody(el, originalXml)
}

/** pPr children we model (extracted and reordered by group during surgical patches; others like defRPr/tabLst untouched). */
const PPR_CHILD_RE: Record<'lnSpc' | 'spcBef' | 'spcAft' | 'bullet', RegExp> = {
  lnSpc: /<a:lnSpc\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:lnSpc>)/g,
  spcBef: /<a:spcBef\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:spcBef>)/g,
  spcAft: /<a:spcAft\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:spcAft>)/g,
  bullet:
    /<a:bu(?:ClrTx|Clr|SzTx|SzPct|SzPts|FontTx|Font|None|AutoNum|Char)\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:bu(?:ClrTx|Clr|SzTx|SzPct|SzPts|FontTx|Font|None|AutoNum|Char)>)/g,
}

/** Generate one group of pPr children from the model paragraph (for surgical patches). */
function buildPPrGroup(p: Paragraph, group: 'lnSpc' | 'spcBef' | 'spcAft' | 'bullet'): string {
  switch (group) {
    case 'lnSpc':
      if (p.lineExact != null)
        return `<a:lnSpc><a:spcPts val="${Math.round(p.lineExact * 100)}"/></a:lnSpc>`
      if (p.lineHeight != null)
        return `<a:lnSpc><a:spcPct val="${Math.round(p.lineHeight * 1000)}"/></a:lnSpc>`
      return ''
    case 'spcBef':
      if (p.spaceBefore != null)
        return `<a:spcBef><a:spcPts val="${Math.round(p.spaceBefore * 100)}"/></a:spcBef>`
      if (p.spaceBeforePct != null)
        return `<a:spcBef><a:spcPct val="${Math.round(p.spaceBeforePct * 1000)}"/></a:spcBef>`
      return ''
    case 'spcAft':
      if (p.spaceAfter != null)
        return `<a:spcAft><a:spcPts val="${Math.round(p.spaceAfter * 100)}"/></a:spcAft>`
      if (p.spaceAfterPct != null)
        return `<a:spcAft><a:spcPct val="${Math.round(p.spaceAfterPct * 1000)}"/></a:spcAft>`
      return ''
    case 'bullet': {
      const b = p.bullet
      if (!b) return ''
      if (b.type === 'none') return '<a:buNone/>'
      let s = ''
      if (b.color) s += `<a:buClr><a:srgbClr val="${hex6(b.color)}"/></a:buClr>`
      if (b.sizePct != null) s += `<a:buSzPct val="${Math.round(b.sizePct * 1000)}"/>`
      if (b.font) s += `<a:buFont typeface="${escapeXmlAttr(b.font)}"/>`
      s +=
        b.type === 'number'
          ? `<a:buAutoNum type="${escapeXmlAttr(b.numType ?? 'arabicPeriod')}"/>`
          : `<a:buChar char="${escapeXmlAttr(b.char ?? '•')}"/>`
      return s
    }
  }
}

/**
 * Surgically patch the <a:pPr> of one <a:p>…</a:p> slice: only reorder groups we
 * model that are flagged dirty (lnSpc/spcBef/spcAft/bu*); non-dirty groups keep
 * their original bytes and join the reorder in schema order; unmodeled children
 * like defRPr/tabLst stay in place (schema order puts them after these groups, so
 * still valid). Attrs (algn/marL/indent) are set in place.
 */
export function patchParagraphPPrXml(paraXml: string, p: Paragraph, which: PPrDirty): string {
  const ALIGN_MAP: Record<string, string> = {
    left: 'l',
    center: 'ctr',
    right: 'r',
    justify: 'just',
  }

  // Locate/create pPr (self-closing expanded to a pair for uniform handling)
  let m = /<a:pPr\b([^>]*?)(\/?)>/.exec(paraXml)
  if (!m) {
    paraXml = paraXml.replace(/^<a:p>/, '<a:p><a:pPr></a:pPr>')
    m = /<a:pPr\b([^>]*?)(\/?)>/.exec(paraXml)!
  } else if (m[2] === '/') {
    paraXml =
      paraXml.slice(0, m.index) + `<a:pPr${m[1]}></a:pPr>` + paraXml.slice(m.index + m[0].length)
    m = /<a:pPr\b([^>]*?)(\/?)>/.exec(paraXml)!
  }
  const openEnd = m.index + m[0].length
  const closeAt = paraXml.indexOf('</a:pPr>', openEnd)
  if (closeAt < 0) return paraXml
  let openTag = paraXml.slice(m.index, openEnd)
  let inner = paraXml.slice(openEnd, closeAt)

  // Attribute patch
  const setPPrAttr = (name: string, value: string | undefined) => {
    const re = new RegExp(`\\s${name}="[^"]*"`)
    if (value === undefined) {
      openTag = openTag.replace(re, '')
      return
    }
    if (re.test(openTag)) openTag = openTag.replace(re, ` ${name}="${escapeXmlAttr(value)}"`)
    else openTag = openTag.replace(/^<a:pPr/, `<a:pPr ${name}="${escapeXmlAttr(value)}"`)
  }
  if (which.align) setPPrAttr('algn', p.align ? ALIGN_MAP[p.align] : undefined)
  if (which.level) setPPrAttr('lvl', p.level ? String(p.level) : undefined)
  if (which.indents) {
    setPPrAttr('marL', p.marL != null ? String(Math.round(p.marL)) : undefined)
    setPPrAttr('indent', p.indent != null ? String(Math.round(p.indent)) : undefined)
  }

  // Child groups: extract all (keeping original bytes), regenerate dirty groups from the model, then splice back at the start in schema order
  const groups: Array<'lnSpc' | 'spcBef' | 'spcAft' | 'bullet'> = [
    'lnSpc',
    'spcBef',
    'spcAft',
    'bullet',
  ]
  const kept: Record<string, string> = {}
  for (const g of groups) {
    const re = PPR_CHILD_RE[g]
    re.lastIndex = 0
    kept[g] = (inner.match(re) ?? []).join('')
    inner = inner.replace(re, '')
  }
  const block = groups.map((g) => (which[g] ? buildPPrGroup(p, g) : kept[g])).join('')

  return paraXml.slice(0, m.index) + openTag + block + inner + paraXml.slice(closeAt)
}

/** Locate all top-level <a:p>…</a:p> spans in the txBody in document order. */
function findParagraphSpans(xml: string): Array<{ start: number; end: number }> {
  const spans: Array<{ start: number; end: number }> = []
  const re = /<a:p>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const close = xml.indexOf('</a:p>', re.lastIndex)
    if (close < 0) break
    const end = close + '</a:p>'.length
    spans.push({ start: m.index, end })
    re.lastIndex = end
  }
  return spans
}

/**
 * Element-level paragraph formatting patch: surgically patch pPr per <a:p> (run
 * bytes untouched). Returns null when the paragraph count doesn't match the model
 * (the caller falls back to rebuild).
 */
export function patchElementPPr(el: TextElement, xml: string, which: PPrDirty): string | null {
  const paras = el.text?.paragraphs ?? []
  const spans = findParagraphSpans(xml)
  if (spans.length !== paras.length || !spans.length) return null
  let out = ''
  let cursor = 0
  for (let i = 0; i < spans.length; i++) {
    const s = spans[i]!
    out += xml.slice(cursor, s.start)
    out +=
      which.paraIndices && !which.paraIndices.includes(i)
        ? xml.slice(s.start, s.end)
        : patchParagraphPPrXml(xml.slice(s.start, s.end), paras[i]!, which)
    cursor = s.end
  }
  out += xml.slice(cursor)
  return out
}

/** Soft-break sentinel run (parse rewrites <a:br/> as a run with text="\n"). */
function isSoftBreakRun(r: TextRun): boolean {
  return r.text === '\n' && !r.field
}

interface Span {
  start: number
  end: number
  kind: 'r' | 'br'
}

/** Locate all top-level <a:r>…</a:r> and <a:br/> (incl. paired form) spans in document order. */
function findRunSpans(xml: string): Span[] {
  const spans: Span[] = []
  const re = /<a:r>|<a:r\s[^>]*>|<a:br\b[^>]*\/>|<a:br\b[^>]*>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const start = m.index
    if (m[0].startsWith('<a:br')) {
      if (m[0].endsWith('/>')) {
        spans.push({ start, end: re.lastIndex, kind: 'br' })
      } else {
        const close = xml.indexOf('</a:br>', re.lastIndex)
        if (close < 0) break
        const end = close + '</a:br>'.length
        spans.push({ start, end, kind: 'br' })
        re.lastIndex = end
      }
      continue
    }
    const close = xml.indexOf('</a:r>', re.lastIndex)
    if (close < 0) break
    const end = close + '</a:r>'.length
    spans.push({ start, end, kind: 'r' })
    re.lastIndex = end
  }
  return spans
}

/** Patch a single <a:r>: replace the <a:t> text + adjust <a:rPr> formatting as needed. */
function patchRun(runXml: string, run: TextRun): string {
  let out = runXml

  // 1. Replace the <a:t> text content (keeping attributes like xml:space)
  out = out.replace(
    /(<a:t(?:\s[^>]*)?>)([\s\S]*?)(<\/a:t>)/,
    (_all, open: string, _text: string, close: string) => {
      return `${open}${escapeXmlText(run.text)}${close}`
    },
  )
  // If the original run has no <a:t> (rare), leave it alone
  if (!/<a:t/.test(runXml)) return out

  // 2. Patch <a:rPr> boolean/size attributes + solidFill color (only when the model has explicit values)
  out = patchRunProps(out, run)
  return out
}

/** Patch or inject the b/i/u/sz attributes and color on <a:rPr>. */
function patchRunProps(runXml: string, run: TextRun): string {
  const attrPatch = (rprOpen: string): string => {
    let tag = rprOpen
    tag = setBoolAttr(tag, 'b', run.bold)
    tag = setBoolAttr(tag, 'i', run.italic)
    // Underline: keep the original style (dbl/wavy… not collapsed to sng); on removal
    // an existing u becomes none, and no u is injected when there was none (keeping bytes).
    // underlineImplicit (hlink styling) is display-only — never bake it into a u attr
    const uVal = run.underline
      ? run.underlineImplicit
        ? undefined
        : (run.underlineStyle ?? 'sng')
      : /\su="[^"]*"/.test(tag)
        ? 'none'
        : undefined
    tag = setAttr(tag, 'u', uVal, /\su="[^"]*"/)
    // Strikethrough: same semantics as underline — on removal an existing strike becomes noStrike, none injected when absent
    const strikeVal = run.strike
      ? (run.strikeStyle ?? 'sngStrike')
      : /\sstrike="[^"]*"/.test(tag)
        ? 'noStrike'
        : undefined
    tag = setAttr(tag, 'strike', strikeVal, /\sstrike="[^"]*"/)
    // Superscript/subscript: 0 = none (with an existing attribute write an explicit 0 to disable; none injected when absent)
    const blVal = run.baseline
      ? String(Math.round(run.baseline * 1000))
      : /\sbaseline="[^"]*"/.test(tag)
        ? '0'
        : undefined
    tag = setAttr(tag, 'baseline', blVal, /\sbaseline="[^"]*"/)
    tag = setAttr(
      tag,
      'sz',
      run.fontSize != null && !run.fontSizeImplicit
        ? String(Math.round(run.fontSize * 100))
        : undefined,
      /\ssz="[^"]*"/,
    )
    return tag
  }

  const hasRPr = /<a:rPr\b/.test(runXml)
  if (hasRPr) {
    // Handle self-closing <a:rPr .../> and paired <a:rPr ...>…</a:rPr>
    runXml = runXml.replace(/<a:rPr\b([^>]*?)(\/?)>/, (_all, attrs: string, selfClose: string) => {
      const patched = attrPatch(`<a:rPr${attrs}`)
      return `${patched}${selfClose ? '/>' : '>'}`
    })
    // Color: patch or inject solidFill (child nodes can only be injected inside the paired form).
    // colorFollowsTheme = the display value comes from schemeClr/inheritance and the
    // user hasn't changed it → don't write, keep the original bytes (schemeClr with
    // lumMod/alpha modifiers stays linked to the theme)
    if (run.color && !run.colorFollowsTheme) {
      runXml = patchRunColor(runXml, run.color)
    }
    // Font: patch/inject <a:latin>/<a:ea> only when the user actually changed it.
    // latinFont/eaFont/fontImplicit present = font untouched → keep original bytes
    // (dual fonts and theme references stay intact)
    if (run.fontFamily && !run.latinFont && !run.eaFont && !run.fontImplicit) {
      runXml = patchRunFont(runXml, run.fontFamily)
    }
    runXml = patchRunHlink(runXml, run)
  } else {
    // No rPr: inject a minimal rPr after <a:r> (no font slots injected when the font is untouched,
    // so inheritance applies)
    const attrs = buildRPrAttrs(run)
    const color = run.color
      ? `<a:solidFill><a:srgbClr val="${hex6(run.color)}"/></a:solidFill>`
      : ''
    const font =
      run.fontFamily && !run.fontImplicit && !run.latinFont && !run.eaFont
        ? fontSlotsXml(escapeXmlAttr(run.fontFamily))
        : ''
    const inner = color + font + hlinkXml(run) // schema order: fill before the font slots, hlinkClick after
    const rpr = inner ? `<a:rPr${attrs}>${inner}</a:rPr>` : `<a:rPr${attrs}/>`
    runXml = runXml.replace(/(<a:r(?:\s[^>]*)?>)/, `$1${rpr}`)
  }
  return runXml
}

/** <a:hlinkClick> for the model's rId (empty when none). */
function hlinkXml(run: TextRun): string {
  if (!run.hyperlinkRId) return ''
  return (
    `<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${escapeXmlAttr(run.hyperlinkRId)}"` +
    (run.hyperlinkAction ? ` action="${escapeXmlAttr(run.hyperlinkAction)}"` : '') +
    (run.hyperlinkTooltip ? ` tooltip="${escapeXmlAttr(run.hyperlinkTooltip)}"` : '') +
    '/>'
  )
}

/** Sync <a:hlinkClick> in the rPr with the model: no-op when the rId already matches (keeping bytes). */
function patchRunHlink(runXml: string, run: TextRun): string {
  const existing = /<a:hlinkClick\b[^>]*?\br:id="([^"]*)"/.exec(runXml)
  if ((existing?.[1] ?? undefined) === run.hyperlinkRId) return runXml
  // Strip the old one (self-closing or paired)
  runXml = runXml.replace(
    /<a:hlinkClick\b[^>]*\/>|<a:hlinkClick\b[^>]*>[\s\S]*?<\/a:hlinkClick>/,
    '',
  )
  const hlink = hlinkXml(run)
  if (!hlink) return runXml
  // Self-closing rPr → expand to a pair
  if (/<a:rPr\b[^>]*\/>/.test(runXml)) {
    return runXml.replace(/<a:rPr\b([^>]*?)\/>/, `<a:rPr$1>${hlink}</a:rPr>`)
  }
  // hlinkClick sits after latin/ea/cs/sym but before rtl/extLst
  const m = /<a:(?:rtl|extLst)\b/.exec(runXml)
  if (m) return runXml.slice(0, m.index) + hlink + runXml.slice(m.index)
  return runXml.replace(/<\/a:rPr>/, `${hlink}</a:rPr>`)
}

/** The three script slots written together: with a:latin alone PowerPoint still renders CJK from
 *  the theme ea font and Arabic/Hebrew/Indic from the theme cs font, so a font change appears to
 *  do nothing to that text. Takes an already-escaped typeface. */
function fontSlotsXml(esc: string): string {
  return `<a:latin typeface="${esc}"/><a:ea typeface="${esc}"/><a:cs typeface="${esc}"/>`
}

/** The font slots, in CT_TextCharacterProperties order. */
const FONT_SLOTS = ['a:latin', 'a:ea', 'a:cs']
const IS_FONT_SLOT = new Set(FONT_SLOTS)
/** The rPr children that follow the font slots, so the group is inserted ahead of the first one. */
const AFTER_FONT_SLOTS = new Set(['a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'])

/**
 * Point an element's own typeface attribute at a new value. Only the start tag's own attributes
 * are considered, so a payload nested inside a paired element keeps whatever it declared, and the
 * attribute name is read as a whole token so one merely ending in "typeface" is left alone.
 */
function setTypeface(el: string, esc: string): string {
  const tag = /^<[^\s/>]+(?:"[^"]*"|'[^']*'|[^"'>])*>/.exec(el)?.[0]
  if (!tag) return el
  // sliced rather than grouped: an attribute pattern permissive enough for unquoted values also
  // swallows the '/' of a self-closing tag, which then reappears in front of whatever is added
  const name = /^<[^\s/>]+/.exec(tag)![0]
  const close = tag.endsWith('/>') ? '/>' : '>'
  let had = false
  const attrs = tag
    .slice(name.length, tag.length - close.length)
    .replace(/\s([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')/g, (whole, attr: string) => {
      if (attr !== 'typeface') return whole
      had = true
      return ` typeface="${esc}"`
    })
  // added here rather than letting a caller infer absence from an unchanged string: setting a slot
  // to the typeface it already carries returns the element untouched, which is not the same as it
  // having no typeface at all
  return name + (had ? attrs : `${attrs} typeface="${esc}"`) + close + el.slice(tag.length)
}

/** CT_TextCharacterProperties children. Anything else means the scan below cannot be trusted. */
const RPR_KNOWN_CHILDREN = new Set([
  'a:ln',
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'a:pattFill',
  'a:grpFill',
  'a:effectLst',
  'a:effectDag',
  'a:highlight',
  'a:uLnTx',
  'a:uLn',
  'a:uFillTx',
  'a:uFill',
  'a:latin',
  'a:ea',
  'a:cs',
  'a:sym',
  'a:hlinkClick',
  'a:hlinkMouseOver',
  'a:rtl',
  'a:extLst',
])

/**
 * The font patch as it stood before complex-script support, kept verbatim for any run this file
 * cannot scan structurally. Markup compatibility branches, comments, CDATA and processing
 * instructions can all make element-looking text appear where a child is not, so rather than
 * guess, such a run gets exactly the previous behaviour and a:cs is left untouched.
 */
function patchRunFontUnscannable(runXml: string, esc: string): string {
  if (/<a:latin\b/.test(runXml)) {
    runXml = runXml.replace(/(<a:latin\b[^>]*?\btypeface=")[^"]*(")/, `$1${esc}$2`)
    if (/<a:ea\b/.test(runXml)) {
      runXml = runXml.replace(/(<a:ea\b[^>]*?\btypeface=")[^"]*(")/, `$1${esc}$2`)
    } else {
      runXml = runXml.replace(
        /(<a:latin\b[^>]*\/>|<a:latin\b[^>]*>[\s\S]*?<\/a:latin>)/,
        `$1<a:ea typeface="${esc}"/>`,
      )
    }
    return runXml
  }
  const fonts = `<a:latin typeface="${esc}"/><a:ea typeface="${esc}"/>`
  if (/<a:rPr\b[^>]*\/>/.test(runXml)) {
    return runXml.replace(/<a:rPr\b([^>]*?)\/>/, `<a:rPr$1>${fonts}</a:rPr>`)
  }
  return runXml.replace(/<\/a:rPr>/, `${fonts}</a:rPr>`)
}

/**
 * Patch or inject the run font: <a:latin>/<a:ea>/<a:cs> typeface, kept in sync per fontSlotsXml.
 * Works on the direct children of the run's own <a:rPr>, so a typeface nested inside a:ln, an
 * a:extLst extension payload or any other subtree stays with whatever owns it.
 */
function patchRunFont(runXml: string, family: string): string {
  const esc = escapeXmlAttr(family)
  // Comments, CDATA and processing instructions can all carry element-looking text, so the scan
  // below would read structure that is not there
  if (/<!--|<!\[CDATA\[|<\?/.test(runXml)) return patchRunFontUnscannable(runXml, esc)
  // topLevelChildren reads a name as [a-zA-Z][\w:]*, while XML allows a leading underscore, a
  // non-ASCII character and an interior middle dot among others. A name outside what it reads
  // would be taken short and the element mistaken for a different one, so refuse the run rather
  // than widen a scanner five other callers share.
  for (const [, name] of runXml.matchAll(/<\/?([^\s/>!?][^\s/>]*)/g)) {
    if (!/^[a-zA-Z][\w:]*$/.test(name!)) return patchRunFontUnscannable(runXml, esc)
  }
  // The run's OWN properties, taken as a direct child rather than as the first a:rPr in the text:
  // an extension payload may carry one of its own, and filling that one in would leave the real
  // slots untouched.
  const runOpen = /^<[^\s/>]+(?:"[^"]*"|'[^']*'|[^"'>])*?>/.exec(runXml)?.[0]
  if (!runOpen) return patchRunFontUnscannable(runXml, esc)
  const rPr = topLevelChildren(runXml, runOpen.length, runXml.length).find(
    (c) => c.name === 'a:rPr',
  )
  // could not be located structurally, so hand it to the previous behaviour rather than
  // silently doing nothing: a payload's own closing tag can truncate the slice this is given
  if (!rPr) return patchRunFontUnscannable(runXml, esc)
  const el = runXml.slice(rPr.start, rPr.end)
  // Self-closing rPr → expand to a pair around the slots. Checked before the open-tag match:
  // that pattern's [^"'>] accepts '/', so it would swallow a self-closing tag whole and this
  // branch could never run.
  const selfAttrs = /^<a:rPr\b((?:"[^"]*"|'[^']*'|[^"'>])*?)\/>$/.exec(el)?.[1]
  if (selfAttrs !== undefined) {
    return (
      runXml.slice(0, rPr.start) +
      `<a:rPr${selfAttrs}>${fontSlotsXml(esc)}</a:rPr>` +
      runXml.slice(rPr.end)
    )
  }
  const rPrOpen = /^<a:rPr\b(?:"[^"]*"|'[^']*'|[^"'>])*?>/.exec(el)?.[0]
  if (!rPrOpen) return patchRunFontUnscannable(runXml, esc)
  const innerStart = rPr.start + rPrOpen.length
  // the closing tag's own start, since a valid one may carry whitespace before its '>'
  const innerEnd = rPr.start + el.lastIndexOf('</')
  if (innerEnd <= innerStart) return patchRunFontUnscannable(runXml, esc)
  const children = topLevelChildren(runXml, innerStart, innerEnd)
  // An unrecognised child is a markup-compatibility wrapper or something else that supplies the
  // real children indirectly, whatever prefix it was bound to. Recognising the whole list is the
  // only way to know the slots found here are the run's own.
  if (children.some((c) => !RPR_KNOWN_CHILDREN.has(c.name))) {
    return patchRunFontUnscannable(runXml, esc)
  }
  const slots = children.filter((c) => IS_FONT_SLOT.has(c.name))
  // A repeated slot is not something to reason about: the group is rebuilt from one of them and
  // the others would be dropped along with their attributes. The schema allows one of each.
  if (new Set(slots.map((c) => c.name)).size !== slots.length) {
    return patchRunFontUnscannable(runXml, esc)
  }
  // Rebuild the group whole so it comes out as latin, ea, cs however few of the three the run
  // arrived with, keeping each existing element's other attributes (pitchFamily, charset, panose)
  const block = FONT_SLOTS.map((tag) => {
    const found = slots.find((s) => s.name === tag)
    if (!found) return `<${tag} typeface="${esc}"/>`
    return setTypeface(runXml.slice(found.start, found.end), esc)
  }).join('')
  // Every removal below sits at or after this point, so the index stays valid through them
  const at = slots.length
    ? slots[0]!.start
    : (children.find((c) => AFTER_FONT_SLOTS.has(c.name))?.start ?? innerEnd)
  let out = runXml
  for (const c of [...slots].reverse()) out = out.slice(0, c.start) + out.slice(c.end)
  return out.slice(0, at) + block + out.slice(at)
}

function patchRunColor(runXml: string, color: string): string {
  const hex = hex6(color)
  // Existing solidFill/srgbClr → change val
  if (/<a:solidFill>\s*<a:srgbClr\b/.test(runXml)) {
    return runXml.replace(/(<a:solidFill>\s*<a:srgbClr\b[^>]*?\bval=")[^"]*(")/, `$1${hex}$2`)
  }
  // Existing solidFill/schemeClr → swap to srgbClr (fixed as an explicit color after editing)
  if (/<a:solidFill>\s*<a:schemeClr\b/.test(runXml)) {
    return runXml.replace(
      /<a:solidFill>\s*<a:schemeClr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:schemeClr>)\s*<\/a:solidFill>/,
      `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`,
    )
  }
  // No solidFill: expand a self-closing rPr to a pair first, then inject at the start
  const fill = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`
  if (/<a:rPr\b[^>]*\/>/.test(runXml)) {
    return runXml.replace(/<a:rPr\b([^>]*?)\/>/, `<a:rPr$1>${fill}</a:rPr>`)
  }
  return runXml.replace(/(<a:rPr\b[^>]*>)/, `$1${fill}`)
}

function buildRPrAttrs(run: TextRun): string {
  let s = ''
  if (run.fontSize != null && !run.fontSizeImplicit) s += ` sz="${Math.round(run.fontSize * 100)}"`
  if (run.bold) s += ' b="1"'
  if (run.italic) s += ' i="1"'
  if (run.underline && !run.underlineImplicit)
    s += ` u="${escapeXmlAttr(run.underlineStyle ?? 'sng')}"`
  if (run.strike) s += ` strike="${escapeXmlAttr(run.strikeStyle ?? 'sngStrike')}"`
  if (run.letterSpacing) s += ` spc="${Math.round(run.letterSpacing * 100)}"`
  if (run.baseline) s += ` baseline="${Math.round(run.baseline * 1000)}"`
  return s
}

/** Set a boolean attribute (true→="1", false→explicit ="0" overriding inheritance, undefined→untouched). */
function setBoolAttr(tag: string, name: string, val: boolean | undefined): string {
  if (val === undefined) return tag
  return setAttr(tag, name, val ? '1' : '0', new RegExp(`\\s${name}="[^"]*"`))
}

/** Set/replace/delete an attribute. value=undefined means untouched (keep the original value). */
function setAttr(tag: string, name: string, value: string | undefined, existingRe: RegExp): string {
  if (value === undefined) return tag
  if (existingRe.test(tag)) {
    return tag.replace(existingRe, ` ${name}="${escapeXmlAttr(value)}"`)
  }
  // Inject after the tag name (after <a:rPr)
  return tag.replace(/^<a:rPr/, `<a:rPr ${name}="${escapeXmlAttr(value)}"`)
}

/** Rebuild the <p:txBody>'s paragraph content on structural change, keeping the txBody wrapper and <a:bodyPr>. */
export function rebuildTxBody(el: TextElement, originalXml: string): string {
  const body = el.text!
  const paras = body.paragraphs.map((p) => generateParagraphXml(p)).join('')

  // Keep the original txBody's bodyPr / lstStyle prefix verbatim; only the <a:p>… after it is replaced
  const txOpen = /<p:txBody\b[^>]*>/.exec(originalXml)
  if (!txOpen) {
    // No txBody originally (a plain shape gained text): inject a full txBody after spPr
    const txBody = `<p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody>`
    if (/<\/p:spPr>/.test(originalXml)) {
      return originalXml.replace(/(<\/p:spPr>)/, `$1${txBody}`)
    }
    return originalXml
  }
  const txStart = txOpen.index
  const txContentStart = txStart + txOpen[0].length
  const txEnd = originalXml.lastIndexOf('</p:txBody>')
  const inner = originalXml.slice(txContentStart, txEnd)

  // Keep bodyPr + lstStyle, discard the old <a:p>
  const bodyPr =
    /<a:bodyPr\b(?:[^>]*?)(?:\/>|>[\s\S]*?<\/a:bodyPr>)/.exec(inner)?.[0] ?? '<a:bodyPr/>'
  const lstStyle = /<a:lstStyle\b(?:[^>]*?)(?:\/>|>[\s\S]*?<\/a:lstStyle>)/.exec(inner)?.[0] ?? ''

  return originalXml.slice(0, txContentStart) + bodyPr + lstStyle + paras + originalXml.slice(txEnd)
}

/**
 * Generate <a:p> from a model paragraph (for the rebuild path).
 * Paragraph properties write only explicit items (per the pPrExplicit flags; no
 * flags = a newly created element, all model values treated as explicit); display
 * values inherited from lstStyle/placeholder/master are not baked in — the rebuild
 * keeps lstStyle and placeholder attributes, so inheritance still resolves along
 * the original chain in PowerPoint.
 */
export function generateParagraphXml(p: Paragraph): string {
  const alignMap: Record<string, string> = { left: 'l', center: 'ctr', right: 'r', justify: 'just' }
  const ex = p.pPrExplicit
  const want = (k: keyof NonNullable<Paragraph['pPrExplicit']>) => !ex || !!ex[k]

  const pPrAttrs: string[] = []
  if (p.marL != null && want('marL')) pPrAttrs.push(`marL="${Math.round(p.marL)}"`)
  if (p.indent != null && want('indent')) pPrAttrs.push(`indent="${Math.round(p.indent)}"`)
  if (p.align && want('align')) pPrAttrs.push(`algn="${alignMap[p.align]}"`)
  if (p.level) pPrAttrs.push(`lvl="${p.level}"`)

  // CT_TextParagraphProperties child order: lnSpc → spcBef → spcAft → buClr → buSzPct → buFont → bu*
  let kids = ''
  if (want('lnSpc')) {
    if (p.lineExact != null)
      kids += `<a:lnSpc><a:spcPts val="${Math.round(p.lineExact * 100)}"/></a:lnSpc>`
    else if (p.lineHeight != null)
      kids += `<a:lnSpc><a:spcPct val="${Math.round(p.lineHeight * 1000)}"/></a:lnSpc>`
  }
  if (want('spcBef')) {
    if (p.spaceBefore != null)
      kids += `<a:spcBef><a:spcPts val="${Math.round(p.spaceBefore * 100)}"/></a:spcBef>`
    else if (p.spaceBeforePct != null)
      kids += `<a:spcBef><a:spcPct val="${Math.round(p.spaceBeforePct * 1000)}"/></a:spcBef>`
  }
  if (want('spcAft')) {
    if (p.spaceAfter != null)
      kids += `<a:spcAft><a:spcPts val="${Math.round(p.spaceAfter * 100)}"/></a:spcAft>`
    else if (p.spaceAfterPct != null)
      kids += `<a:spcAft><a:spcPct val="${Math.round(p.spaceAfterPct * 1000)}"/></a:spcAft>`
  }
  if (p.bullet && want('bullet')) {
    const b = p.bullet
    if (b.type === 'none') kids += '<a:buNone/>'
    else {
      if (b.color) kids += `<a:buClr><a:srgbClr val="${hex6(b.color)}"/></a:buClr>`
      if (b.sizePct != null) kids += `<a:buSzPct val="${Math.round(b.sizePct * 1000)}"/>`
      if (b.font) kids += `<a:buFont typeface="${escapeXmlAttr(b.font)}"/>`
      kids +=
        b.type === 'number'
          ? `<a:buAutoNum type="${escapeXmlAttr(b.numType ?? 'arabicPeriod')}"/>`
          : `<a:buChar char="${escapeXmlAttr(b.char ?? '•')}"/>`
    }
  }

  const attrStr = pPrAttrs.length ? ` ${pPrAttrs.join(' ')}` : ''
  const pPr = kids
    ? `<a:pPr${attrStr}>${kids}</a:pPr>`
    : pPrAttrs.length
      ? `<a:pPr${attrStr}/>`
      : ''
  const runs = p.runs.map((r) => generateRunXml(r)).join('')
  return `<a:p>${pPr}${runs}</a:p>`
}

function generateRunXml(r: TextRun): string {
  // Soft-break sentinel → <a:br/>; embedded "\n" in text (new editor Shift+Enter input) splits into alternating run+br
  if (isSoftBreakRun(r)) return '<a:br/>'
  if (r.text.includes('\n') && !r.field) {
    return r.text
      .split('\n')
      .map((part) => (part ? generateRunXml({ ...r, text: part }) : ''))
      .join('<a:br/>')
  }
  const attrs = buildRPrAttrs(r)
  // Text outline (WordArt): the CT_TextCharacterProperties child order requires ln before fill
  const ln = r.outline
    ? `<a:ln w="${Math.round(r.outline.widthEmu)}"><a:solidFill><a:srgbClr val="${hex6(r.outline.color)}"/></a:solidFill></a:ln>`
    : ''
  // Color purely inherited (rPr has no solidFill) → don't write; the inheritance chain still resolves in PowerPoint;
  // explicit schemeClr is materialized as srgbClr to keep visuals (the rebuild path can't restore the original scheme reference)
  const color =
    r.color && !r.colorInherited
      ? `<a:solidFill><a:srgbClr val="${hex6(r.color)}"/></a:solidFill>`
      : ''
  // Original dual fonts/theme references restored first; no declaration (inherited) writes nothing;
  // a user-changed font writes all three slots, so a rebuilt run keeps the typeface the user picked
  const font =
    r.latinFont || r.eaFont || r.csFont
      ? (r.latinFont ? `<a:latin typeface="${escapeXmlAttr(r.latinFont)}"/>` : '') +
        (r.eaFont ? `<a:ea typeface="${escapeXmlAttr(r.eaFont)}"/>` : '') +
        (r.csFont ? `<a:cs typeface="${escapeXmlAttr(r.csFont)}"/>` : '')
      : r.fontFamily && !r.fontImplicit
        ? fontSlotsXml(escapeXmlAttr(r.fontFamily))
        : ''
  // Run-level hyperlink: rId written back (allocated by ensureRunLinkRels for links set this session)
  const hlink = hlinkXml(r)
  const rprInner = ln + color + font + hlink
  const rPr = rprInner
    ? `<a:rPr${attrs}>${rprInner}</a:rPr>`
    : attrs
      ? `<a:rPr${attrs}/>`
      : '<a:rPr/>'
  // Dynamic fields (slide number/date): <a:fld> has the same structure as <a:r>, text is the cached value, refreshed when PowerPoint opens
  if (r.field) {
    return (
      `<a:fld id="{${fieldGuid()}}" type="${escapeXmlAttr(r.field)}">` +
      `${rPr}<a:t>${escapeXmlText(r.text)}</a:t></a:fld>`
    )
  }
  return `<a:r>${rPr}<a:t>${escapeXmlText(r.text)}</a:t></a:r>`
}

/** <a:fld> requires a unique GUID id. */
function fieldGuid(): string {
  const hex = () =>
    Math.floor(Math.random() * 0x10000)
      .toString(16)
      .toUpperCase()
      .padStart(4, '0')
  return `${hex()}${hex()}-${hex()}-${hex()}-${hex()}-${hex()}${hex()}${hex()}`
}

/** Normalize to 6-digit uppercase hex (dropping # and alpha). */
function hex6(color: string): string {
  let c = color.replace(/^#/, '').toUpperCase()
  if (c.length >= 6) c = c.slice(0, 6)
  return c
}

/** <a:alpha> child for an #RRGGBBAA color; '' when opaque or 6-digit. */
function alphaXml(color: string): string {
  const c = color.replace(/^#/, '')
  if (c.length < 8) return ''
  const a = parseInt(c.slice(6, 8), 16)
  if (!Number.isFinite(a) || a >= 255) return ''
  return `<a:alpha val="${Math.round((a / 255) * 100000)}"/>`
}

function srgbClrXml(color: string): string {
  const alpha = alphaXml(color)
  return alpha
    ? `<a:srgbClr val="${hex6(color)}">${alpha}</a:srgbClr>`
    : `<a:srgbClr val="${hex6(color)}"/>`
}

// ── Fill patch (dirtyFill) ──────────────────────────────────────────────

const FILL_TAGS = new Set([
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'a:pattFill',
  'a:grpFill',
])

/** Scan top-level children (depth 1) in the [from,to) range of an xml fragment, returning {name,start,end}. */
function topLevelChildren(
  xml: string,
  from: number,
  to: number,
): Array<{ name: string; start: number; end: number }> {
  const out: Array<{ name: string; start: number; end: number }> = []
  const re = /<(\/?)([a-zA-Z][\w:]*)((?:"[^"]*"|'[^']*'|[^"'>])*?)(\/?)>/g
  re.lastIndex = from
  let depth = 0
  let curStart = -1
  let curName = ''
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null && m.index < to) {
    const closing = m[1] === '/'
    const selfClose = m[4] === '/'
    if (!closing && !selfClose) {
      if (depth === 0) {
        curStart = m.index
        curName = m[2]!
      }
      depth++
    } else if (closing) {
      depth--
      if (depth === 0) out.push({ name: curName, start: curStart, end: re.lastIndex })
    } else if (selfClose && depth === 0) {
      out.push({ name: m[2]!, start: m.index, end: re.lastIndex })
    }
  }
  return out
}

/** Expand a self-closing <p:spPr/> (common on placeholders, geometry/fill fully inherited) into a pair for child injection. */
function expandEmptySpPr(xml: string): string {
  return xml.replace(
    /<p:spPr(\s[^>]*?)?\/>/,
    (_all, attrs: string | undefined) => `<p:spPr${attrs ?? ''}></p:spPr>`,
  )
}

/** Gradient fill write-back parameters (subset of the model's Fill gradient). */
export interface GradientFillPatch {
  stops: Array<{ pos: number; color: string }>
  /** 1/60000 degree (for linear) */
  angle?: number
  /** Radial (circle path, center outward) */
  radial?: boolean
}

type FillPatch = 'none' | string | GradientFillPatch | { rawFillXml: string }

function gsLstXml(stops: GradientFillPatch['stops']): string {
  return `<a:gsLst>${stops
    .map(
      (s) =>
        `<a:gs pos="${Math.round(Math.max(0, Math.min(1, s.pos)) * 100000)}"><a:srgbClr val="${hex6(s.color)}"/></a:gs>`,
    )
    .join('')}</a:gsLst>`
}

const RADIAL_PATH_XML =
  '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'

function buildFillXml(fill: FillPatch): string {
  if (typeof fill === 'object' && 'rawFillXml' in fill) return fill.rawFillXml
  if (typeof fill === 'object')
    return `<a:gradFill rotWithShape="1">${gsLstXml(fill.stops)}${
      fill.radial ? RADIAL_PATH_XML : `<a:lin ang="${Math.round(fill.angle ?? 0)}" scaled="1"/>`
    }</a:gradFill>`
  if (fill === 'none') return '<a:noFill/>'
  return `<a:solidFill>${srgbClrXml(fill)}</a:solidFill>`
}

/**
 * In-place patch between fills of the same kind, preserving unmodeled parameters:
 * solid->solid keeps the original color's alpha; grad->grad keeps gradFill attributes
 * (rotWithShape/flip), child nodes like tileRect, the original radial path (off-center
 * fillToRect), and a:lin's scaled. A cross-type replacement reflects user intent, so
 * the whole block is swapped.
 */
function patchFillNodeXml(existingXml: string, existingName: string, fill: FillPatch): string {
  if (typeof fill === 'string' && fill !== 'none' && existingName === 'a:solidFill') {
    // An 8-digit color carries an explicit alpha (FF = explicitly opaque); a 6-digit
    // color is a plain color change and keeps the original alpha.
    if (fill.replace(/^#/, '').length >= 8) return `<a:solidFill>${srgbClrXml(fill)}</a:solidFill>`
    const alpha = /<a:alpha\s[^>]*\/>/.exec(existingXml)?.[0]
    return alpha
      ? `<a:solidFill><a:srgbClr val="${hex6(fill)}">${alpha}</a:srgbClr></a:solidFill>`
      : `<a:solidFill><a:srgbClr val="${hex6(fill)}"/></a:solidFill>`
  }
  if (typeof fill === 'object' && !('rawFillXml' in fill) && existingName === 'a:gradFill') {
    return patchGradFillXml(existingXml, fill)
  }
  return buildFillXml(fill)
}

function patchGradFillXml(gradXml: string, patch: GradientFillPatch): string {
  const open = /^<a:gradFill((?:"[^"]*"|'[^']*'|[^"'>])*?)(\/?)>/.exec(gradXml)
  if (!open || open[2] === '/') return buildFillXml(patch)
  const innerStart = open[0].length
  const innerEnd = gradXml.lastIndexOf('</a:gradFill>')
  if (innerEnd < 0) return buildFillXml(patch)
  const children = topLevelChildren(gradXml, innerStart, innerEnd)

  const parts: string[] = []
  let gsIdx = -1
  let hasShade = false
  for (const c of children) {
    const seg = gradXml.slice(c.start, c.end)
    if (c.name === 'a:gsLst') {
      gsIdx = parts.length
      parts.push(gsLstXml(patch.stops))
    } else if (c.name === 'a:lin' || c.name === 'a:path') {
      if (patch.radial) {
        if (c.name === 'a:path') {
          parts.push(seg) // keep the original path (off-center fillToRect etc.)
          hasShade = true
        }
      } else if (c.name === 'a:lin') {
        const scaled = /\sscaled="([^"]*)"/.exec(seg)?.[1] ?? '1'
        parts.push(`<a:lin ang="${Math.round(patch.angle ?? 0)}" scaled="${scaled}"/>`)
        hasShade = true
      }
    } else {
      parts.push(seg)
    }
  }
  if (gsIdx < 0) return buildFillXml(patch)
  if (!hasShade) {
    const shade = patch.radial
      ? RADIAL_PATH_XML
      : `<a:lin ang="${Math.round(patch.angle ?? 0)}" scaled="1"/>`
    parts.splice(gsIdx + 1, 0, shade) // in the lin/path sequence it immediately follows gsLst
  }
  return `<a:gradFill${open[1] ?? ''}>${parts.join('')}</a:gradFill>`
}

/**
 * In-place patch of an element's shape fill (the fill node that is a direct child of spPr):
 * fill='none' -> <a:noFill/>; hex -> <a:solidFill>. When a direct fill child exists, patch
 * in place for the same kind and replace the whole block across kinds (never touching the
 * solidFill inside a:ln); otherwise insert after the geometry (prstGeom/custGeom) or xfrm.
 */
export function patchElementFill(originalXml: string, fill: FillPatch): string {
  originalXml = expandEmptySpPr(originalXml)
  const spPrOpen = /<p:spPr(\s[^>]*)?>/.exec(originalXml)
  if (!spPrOpen) return originalXml
  const innerStart = spPrOpen.index + spPrOpen[0].length
  const spPrClose = originalXml.indexOf('</p:spPr>', innerStart)
  if (spPrClose < 0) return originalXml

  const children = topLevelChildren(originalXml, innerStart, spPrClose)
  const existing = children.find((c) => FILL_TAGS.has(c.name))
  if (existing) {
    const fillXml = patchFillNodeXml(
      originalXml.slice(existing.start, existing.end),
      existing.name,
      fill,
    )
    return originalXml.slice(0, existing.start) + fillXml + originalXml.slice(existing.end)
  }
  const fillXml = buildFillXml(fill)
  // Insert after the geometry (OOXML order: xfrm → geom → fill → ln)
  const anchor =
    children.find((c) => c.name === 'a:prstGeom' || c.name === 'a:custGeom') ??
    children.find((c) => c.name === 'a:xfrm')
  const at = anchor ? anchor.end : innerStart
  return originalXml.slice(0, at) + fillXml + originalXml.slice(at)
}

// ── Stroke patch (dirtyStroke) ──────────────────────────────────────────

export interface StrokePatch {
  color: string
  widthEmu: number
  /** prstDash preset name; 'solid' removes the prstDash node; undefined keeps the original bytes */
  dash?: string
}

/** In-place patch of <a:ln>: change the w attribute, replace the fill child and (when requested) the prstDash child; all other bytes are kept. */
function patchLnXml(lnXml: string, stroke: StrokePatch | null): string {
  const open = /^<a:ln((?:"[^"]*"|'[^']*'|[^"'>])*?)(\/?)>/.exec(lnXml)
  if (!open) return lnXml
  let attrs = open[1] ?? ''
  if (stroke) {
    const w = `w="${Math.round(stroke.widthEmu)}"`
    attrs = /\sw="[^"]*"/.test(attrs) ? attrs.replace(/\sw="[^"]*"/, ` ${w}`) : ` ${w}${attrs}`
  }
  const fillXml = stroke ? `<a:solidFill>${srgbClrXml(stroke.color)}</a:solidFill>` : '<a:noFill/>'
  const dashXml =
    stroke?.dash && stroke.dash !== 'solid'
      ? `<a:prstDash val="${escapeXmlAttr(stroke.dash)}"/>`
      : ''
  if (open[2] === '/') return `<a:ln${attrs}>${fillXml}${dashXml}</a:ln>`
  const innerStart = open[0].length
  const innerEnd = lnXml.lastIndexOf('</a:ln>')
  if (innerEnd < 0) return lnXml
  const children = topLevelChildren(lnXml, innerStart, innerEnd)
  const fill = children.find((c) => FILL_TAGS.has(c.name))
  // fill is the first child in the CT_LineProperties sequence; insert at the front when missing
  let inner = fill
    ? lnXml.slice(innerStart, fill.start) + fillXml + lnXml.slice(fill.end, innerEnd)
    : fillXml + lnXml.slice(innerStart, innerEnd)
  if (stroke && stroke.dash !== undefined) {
    inner = inner.replace(/<a:prstDash\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:prstDash>)/, '')
    if (dashXml) {
      // CT_LineProperties order: fill → prstDash → join → headEnd → tailEnd
      const fillEnd = inner.indexOf(fillXml) + fillXml.length
      inner = inner.slice(0, fillEnd) + dashXml + inner.slice(fillEnd)
    }
  }
  return `<a:ln${attrs}>${inner}</a:ln>`
}

/**
 * In-place patch of an element's stroke (the <a:ln> direct child of spPr):
 * stroke=null -> replace the fill child with <a:noFill/> (explicit no stroke, overriding
 * theme inheritance); otherwise change the w attribute and replace the fill child with a
 * solid color. With an existing a:ln only these spots are touched; other attributes
 * (cap/cmpd/algn) and children (headEnd/tailEnd/join) keep their original bytes; the
 * prstDash child is only rewritten when stroke.dash is set; without an a:ln, insert
 * after the fill/geometry.
 */
export function patchElementStroke(originalXml: string, stroke: StrokePatch | null): string {
  originalXml = expandEmptySpPr(originalXml)
  const spPrOpen = /<p:spPr(\s[^>]*)?>/.exec(originalXml)
  if (!spPrOpen) return originalXml
  const innerStart = spPrOpen.index + spPrOpen[0].length
  const spPrClose = originalXml.indexOf('</p:spPr>', innerStart)
  if (spPrClose < 0) return originalXml

  const children = topLevelChildren(originalXml, innerStart, spPrClose)
  const existing = children.find((c) => c.name === 'a:ln')
  if (existing) {
    const patched = patchLnXml(originalXml.slice(existing.start, existing.end), stroke)
    return originalXml.slice(0, existing.start) + patched + originalXml.slice(existing.end)
  }

  const lnXml = stroke
    ? `<a:ln w="${Math.round(stroke.widthEmu)}"><a:solidFill>${srgbClrXml(stroke.color)}</a:solidFill>${
        stroke.dash && stroke.dash !== 'solid'
          ? `<a:prstDash val="${escapeXmlAttr(stroke.dash)}"/>`
          : ''
      }</a:ln>`
    : '<a:ln><a:noFill/></a:ln>'
  // Insert after fill / geometry / xfrm (OOXML order: xfrm → geom → fill → ln)
  const anchor =
    children.find((c) => FILL_TAGS.has(c.name)) ??
    children.find((c) => c.name === 'a:prstGeom' || c.name === 'a:custGeom') ??
    children.find((c) => c.name === 'a:xfrm')
  const at = anchor ? anchor.end : spPrClose
  return originalXml.slice(0, at) + lnXml + originalXml.slice(at)
}

// ── Slide background patch ──────────────────────────────────────────────

// ── Picture crop patch (dirtySrcRect) ───────────────────────────────────

/**
 * In-place patch of the <a:srcRect> inside a <p:pic>'s <p:blipFill>:
 * srcRect=null removes the srcRect (show the full image); otherwise writes the
 * l/t/r/b attributes (1/1000 %). Crop values are 0..1 fractions, converted to
 * OOXML's 0..100000 integers.
 */
export function patchPictureSrcRect(
  originalXml: string,
  srcRect: { l: number; t: number; r: number; b: number } | null,
): string {
  // Locate the <p:blipFill> range
  const blipFillOpen = /<p:blipFill(\s[^>]*)?>/.exec(originalXml)
  if (!blipFillOpen) return originalXml
  const innerStart = blipFillOpen.index + blipFillOpen[0].length
  const blipFillClose = originalXml.indexOf('</p:blipFill>', innerStart)
  if (blipFillClose < 0) return originalXml

  // Remove the old srcRect (self-closing)
  const existing = /<a:srcRect[^/]*\/>|<a:srcRect[^>]*>[\s\S]*?<\/a:srcRect>/.exec(
    originalXml.slice(blipFillOpen.index, blipFillClose + '</p:blipFill>'.length),
  )
  let xmlInner = originalXml
  if (existing) {
    const absStart = blipFillOpen.index + existing.index
    xmlInner = originalXml.slice(0, absStart) + originalXml.slice(absStart + existing[0].length)
  }

  if (!srcRect) return xmlInner

  // Build the new srcRect (write only non-zero attributes; omitting 0 values follows OOXML minimalism)
  const to5 = (v: number) => Math.round(Math.max(0, Math.min(1, v)) * 100000)
  const l = to5(srcRect.l)
  const t = to5(srcRect.t)
  const r = to5(srcRect.r)
  const b = to5(srcRect.b)
  const attrs =
    (l ? ` l="${l}"` : '') +
    (t ? ` t="${t}"` : '') +
    (r ? ` r="${r}"` : '') +
    (b ? ` b="${b}"` : '')
  const srcRectXml = `<a:srcRect${attrs}/>`

  // Insert after <a:blip> (OOXML blipFill order: blip → srcRect → fillMode)
  const blipFillOpenNew = /<p:blipFill(\s[^>]*)?>/.exec(xmlInner)
  if (!blipFillOpenNew) return xmlInner
  const innerStartNew = blipFillOpenNew.index + blipFillOpenNew[0].length
  const blipFillCloseNew = xmlInner.indexOf('</p:blipFill>', innerStartNew)
  if (blipFillCloseNew < 0) return xmlInner

  const children = topLevelChildren(xmlInner, innerStartNew, blipFillCloseNew)
  const blipChild = children.find((c) => c.name === 'a:blip')
  const insertAt = blipChild ? blipChild.end : innerStartNew
  return xmlInner.slice(0, insertAt) + srcRectXml + xmlInner.slice(insertAt)
}

/**
 * In-place patch of a slide's bodyPrefix: replace/inject the <p:bg> under <p:cSld>
 * with a solid color. Returns the new bodyPrefix (the caller writes it back to
 * slide.bodyPrefix and flags a rebuild).
 */
export function patchSlideBackgroundXml(bodyPrefix: string, color: string): string {
  const bgXml = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex6(color)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`
  const existing = /<p:bg>[\s\S]*?<\/p:bg>|<p:bg\s[^>]*\/>/.exec(bodyPrefix)
  if (existing) {
    return (
      bodyPrefix.slice(0, existing.index) +
      bgXml +
      bodyPrefix.slice(existing.index + existing[0].length)
    )
  }
  const cSld = /<p:cSld(\s[^>]*)?>/.exec(bodyPrefix)
  if (!cSld) return bodyPrefix
  const at = cSld.index + cSld[0].length
  return bodyPrefix.slice(0, at) + bgXml + bodyPrefix.slice(at)
}

// ── Slide transition patch ──────────────────────────────────────────────

export type SlideTransitionKind =
  | 'none'
  | 'morph'
  | 'fade'
  | 'push'
  | 'wipe'
  | 'split'
  | 'circle'
  | 'cover'
  | 'pull'
  | 'dissolve'
  | 'zoom'
  | 'random'

const TRANSITION_INNER: Record<Exclude<SlideTransitionKind, 'none' | 'morph'>, string> = {
  fade: '<p:fade/>',
  push: '<p:push dir="u"/>',
  wipe: '<p:wipe dir="l"/>',
  split: '<p:split orient="horz" dir="out"/>',
  circle: '<p:circle/>',
  cover: '<p:cover dir="l"/>',
  pull: '<p:pull dir="l"/>',
  dissolve: '<p:dissolve/>',
  zoom: '<p:zoom/>',
  random: '<p:random/>',
}

/**
 * Morph transition: PowerPoint 2019+'s <p159:morph> (2015/main namespace), wrapped
 * in mc:AlternateContent — older PowerPoint uses the Fallback fade without erroring.
 */
const MORPH_TRANSITION_XML =
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/main" Requires="p159">' +
  '<p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="800">' +
  '<p159:morph option="byObject"/></p:transition></mc:Choice>' +
  '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>' +
  '</mc:AlternateContent>'

const TRANSITION_RE = /<p:transition\b[^>]*\/>|<p:transition\b[^>]*>[\s\S]*?<\/p:transition>/
/** AlternateContent-wrapped transition (only counts when Choice is immediately followed by p:transition, avoiding other AC blocks). */
const AC_TRANSITION_RE =
  /<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<p:transition\b[\s\S]*?<\/mc:AlternateContent>/

/** <p:transition>'s schema insertion point: after clrMapOvr, else after </p:cSld> (before timing); -1 when not found. */
function transitionInsertPos(body: string): number {
  const clrMap = /<p:clrMapOvr\b[^>]*\/>|<p:clrMapOvr\b[^>]*>[\s\S]*?<\/p:clrMapOvr>/.exec(body)
  if (clrMap) return clrMap.index + clrMap[0].length
  const cSldEnd = body.indexOf('</p:cSld>')
  return cSldEnd >= 0 ? cSldEnd + '</p:cSld>'.length : -1
}

/**
 * In-place patch of the bodySuffix's <p:transition> (schema order: after clrMapOvr,
 * before timing). kind='none' only removes the effect; 'morph' writes an
 * AlternateContent-wrapped p159:morph. A configured auto-advance time (advTm) is
 * kept when switching effects. Returns
 * the new bodySuffix.
 */
export function patchSlideTransitionXml(bodySuffix: string, kind: SlideTransitionKind): string {
  const advTm = readSlideAdvanceTimeXml(bodySuffix)
  const stripped = bodySuffix.replace(AC_TRANSITION_RE, '').replace(TRANSITION_RE, '')
  if (kind === 'none') return advTm == null ? stripped : patchSlideAdvanceTimeXml(stripped, advTm)
  const xml =
    kind === 'morph'
      ? MORPH_TRANSITION_XML
      : `<p:transition>${TRANSITION_INNER[kind]}</p:transition>`
  const at = transitionInsertPos(stripped)
  if (at < 0) return stripped
  const out = stripped.slice(0, at) + xml + stripped.slice(at)
  return advTm == null ? out : patchSlideAdvanceTimeXml(out, advTm)
}

/** Read the current transition from the bodySuffix (for UI echo; unmodeled effects map to 'none'). */
export function readSlideTransitionXml(bodySuffix: string): SlideTransitionKind {
  // AlternateContent wrapper first (morph, or PowerPoint's newer form with p14:dur)
  const m = AC_TRANSITION_RE.exec(bodySuffix) ?? TRANSITION_RE.exec(bodySuffix)
  if (!m) return 'none'
  if (/<[\w.]+:morph[\s/>]/.test(m[0])) return 'morph'
  const kind = /<p:(fade|push|wipe|split|circle|cover|pull|dissolve|zoom|random)\b/.exec(m[0])?.[1]
  return (kind as SlideTransitionKind | undefined) ?? 'none'
}

// ── Auto-advance time patch (<p:transition advTm="ms">, for saving rehearsed timings) ────

/** Set/remove advTm on every <p:transition> open tag in the block (both Choice/Fallback in morph's AC wrapper). */
function patchAdvTmAttr(block: string, ms: number | null): string {
  return block.replace(/<p:transition\b[^>]*>/g, (tag) => {
    const cleaned = tag.replace(/\s+advTm="[^"]*"/, '')
    if (ms == null) return cleaned
    return cleaned.replace(
      /(\s*\/)?>$/,
      (_m, close: string | undefined) => ` advTm="${ms}"${close ?? ''}>`,
    )
  })
}

/**
 * In-place patch of the bodySuffix's auto-advance time (PowerPoint's "advance slide
 * after / rehearse timings"). ms=null removes it; when the slide has no transition,
 * a new empty <p:transition> carrying only advTm is created. Returns the new bodySuffix.
 */
export function patchSlideAdvanceTimeXml(bodySuffix: string, ms: number | null): string {
  const m = AC_TRANSITION_RE.exec(bodySuffix) ?? TRANSITION_RE.exec(bodySuffix)
  if (m)
    return (
      bodySuffix.slice(0, m.index) +
      patchAdvTmAttr(m[0], ms) +
      bodySuffix.slice(m.index + m[0].length)
    )
  if (ms == null) return bodySuffix
  const at = transitionInsertPos(bodySuffix)
  if (at < 0) return bodySuffix
  return bodySuffix.slice(0, at) + `<p:transition advTm="${ms}"/>` + bodySuffix.slice(at)
}

/** Read the auto-advance time from the bodySuffix (ms; null when unset). */
export function readSlideAdvanceTimeXml(bodySuffix: string): number | null {
  const m = AC_TRANSITION_RE.exec(bodySuffix) ?? TRANSITION_RE.exec(bodySuffix)
  const adv = m && /\badvTm="(\d+)"/.exec(m[0])
  return adv ? Number(adv[1]) : null
}

// ── Hidden slide patch (<p:sld show="0">, skipped during slideshow) ─────

const SLD_OPEN_RE = /<p:sld\b[^>]*>/

/** In-place patch of the show attribute on the bodyPrefix root element; hidden=false removes it. Returns the new bodyPrefix. */
export function patchSlideHiddenXml(bodyPrefix: string, hidden: boolean): string {
  const m = SLD_OPEN_RE.exec(bodyPrefix)
  if (!m) return bodyPrefix
  let tag = m[0].replace(/\s+show="[^"]*"/, '')
  if (hidden) tag = `${tag.slice(0, -1)} show="0">`
  return bodyPrefix.slice(0, m.index) + tag + bodyPrefix.slice(m.index + m[0].length)
}

/** Read whether the bodyPrefix marks the slide as hidden. */
export function readSlideHiddenXml(bodyPrefix: string): boolean {
  const m = SLD_OPEN_RE.exec(bodyPrefix)
  return !!m && /\sshow="0"/.test(m[0])
}

// ── Geometry patch (dirtyTransform) ─────────────────────────────────────

export function generateXfrmXml(t: Transform, tag = 'a:xfrm'): string {
  const attrs =
    (t.rot ? ` rot="${t.rot}"` : '') + (t.flipH ? ' flipH="1"' : '') + (t.flipV ? ' flipV="1"' : '')
  const o = t.offset
  return `<${tag}${attrs}><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/></${tag}>`
}

/**
 * In-place patch of an element's xfrm (off/ext/rot/flip all taken from the model):
 * - sp/pic/grpSp: the first <a:xfrm> inside spPr/grpSpPr; when absent (placeholder
 *   inheriting geometry), inject a full <a:xfrm> after the spPr open tag.
 * - graphicFrame (table/chart/smartart/ole passthrough): <p:xfrm> (directly
 *   containing a:off/a:ext, required by the schema). Embedded content (OLE preview
 *   pictures etc.) may carry its own a:xfrm, so p:xfrm must be matched.
 *   Non-graphicFrame passthroughs have no p:xfrm and are returned unchanged.
 * - Existing xfrm: replace the whole block, but keep children other than off/ext
 *   (a group's chOff/chExt untouched).
 */
export function patchElementXfrm(el: SlideElement, originalXml: string): string {
  const t = el.transform
  const isFrame = el.type === 'table' || el.type === 'chart' || el.type === 'passthrough'
  const xfrmTag = isFrame ? 'p:xfrm' : 'a:xfrm'
  const open = new RegExp(`<${xfrmTag}(\\s[^>]*)?>`).exec(originalXml)
  if (open) {
    const start = open.index
    const close = originalXml.indexOf(`</${xfrmTag}>`, start)
    if (close < 0) return originalXml
    const end = close + `</${xfrmTag}>`.length
    const inner = originalXml.slice(start + open[0].length, close)
    // Keep children other than off/ext (chOff/chExt…)
    const rest = inner.replace(/<a:off\s[^>]*\/>|<a:ext\s[^>]*\/>/g, '')
    const attrs =
      (t.rot ? ` rot="${t.rot}"` : '') +
      (t.flipH ? ' flipH="1"' : '') +
      (t.flipV ? ' flipV="1"' : '')
    const o = t.offset
    const replaced =
      `<${xfrmTag}${attrs}>` +
      `<a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/>` +
      rest +
      `</${xfrmTag}>`
    return originalXml.slice(0, start) + replaced + originalXml.slice(end)
  }
  if (isFrame) return originalXml
  // No xfrm: inject after the spPr / grpSpPr open tag (self-closing <p:spPr/> expanded first)
  const spPr = /<(p:spPr|p:grpSpPr|pic:spPr)(\s[^>]*?)?(\/?)>/.exec(originalXml)
  if (!spPr) return originalXml
  const tag = spPr[1]!
  if (spPr[3] === '/') {
    const openTag = `<${tag}${spPr[2] ?? ''}>`
    return (
      originalXml.slice(0, spPr.index) +
      openTag +
      generateXfrmXml(t) +
      `</${tag}>` +
      originalXml.slice(spPr.index + spPr[0].length)
    )
  }
  const insertAt = spPr.index + spPr[0].length
  return originalXml.slice(0, insertAt) + generateXfrmXml(t) + originalXml.slice(insertAt)
}

/**
 * Write back <a:normAutofit>'s fontScale/lnSpcReduction (thousandths of a percent,
 * 62500 = 62.5%). Clears the attribute when the ratio ≈1 / no reduction
 * (defaults are not written). Only effective when bodyPr already has
 * normAutofit (elements with autofit=shrink always do).
 */
export function patchBodyPrAutofit(
  xml: string,
  fontScale: number,
  lnSpcReduction?: number,
): string {
  const attrs =
    (fontScale < 0.999 ? ` fontScale="${Math.round(fontScale * 100000)}"` : '') +
    (lnSpcReduction ? ` lnSpcReduction="${Math.round(lnSpcReduction * 100000)}"` : '')
  return xml.replace(
    /<a:normAutofit\b[^>]*?(?:\/>|>\s*<\/a:normAutofit>)/,
    `<a:normAutofit${attrs}/>`,
  )
}
