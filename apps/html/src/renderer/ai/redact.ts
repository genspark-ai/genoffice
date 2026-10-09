import type { ParseMap } from '../document/parse-map'

/**
 * Withholding part of an HTML page from the model, without changing the page.
 *
 * ## What it is for
 *
 * The reader hands the whole file to a model to ask one question about layout.
 * Whatever is in it goes along. This is for the things they did not mean to
 * hand over: a key in a `<script>`, a token in a `content=` attribute, a name in
 * a comment. The words stay in the file, the page renders identically, and the
 * model is handed `{{label}}` in their place.
 *
 * It is **not** encryption. It does not stop a model that *asks* for the value,
 * and it does not stop inference from `key.startsWith("sk-")`. It stops the
 * accident.
 *
 * ## Where the mark lives
 *
 * Three marks, all of them legal and inert — a browser renders the page exactly
 * as before, and the app has no JavaScript parser, so nothing reformats them
 * out from under us (verified: no prettier/esprima/acorn/babel in the app).
 *
 * | mark | placed on | withholds |
 * |---|---|---|
 * | `data-gx-redact="label"` | any element | that element's **content** |
 * | `data-gx-redact-<attr>="label"` | any element | that **attribute's value** |
 * | `/*gx:redact:label*\/` | inside a `<script>` | the **string literal** after it |
 *
 * The first two are `data-*` attributes: legal anywhere, invisible, and they
 * round-trip because every edit in this app compiles down to a character-range
 * splice on the source text (see `document/patch.ts` — `Patch` is the only
 * primitive). The third is a JavaScript comment the engine ignores, anchored in
 * front of the literal so which value it means is never in doubt.
 *
 * ## Why the projection is not a plain string replace
 *
 * `read_source` addresses by line number and by element id, and it *prints line
 * numbers*. Replacing a 40-character key with an 11-character marker shifts every
 * line after it, so a line-addressed read would come back with the wrong lines
 * and a `str_replace` composed against them would miss. The projection therefore
 * keeps the replaced spans and can translate an offset in either direction.
 */

import {
  sanitizeLabel as coreSanitizeLabel,
  placeholderInstruction as buildInstruction,
  type PlaceholderSpec,
} from '@genoffice/agent-core/redact-core'

/** Attribute marking an element's own content. */
export const MARK_ATTR = 'data-gx-redact'
/** Prefix for marking one attribute's value: `data-gx-redact-src="label"`. */
export const MARK_ATTR_PREFIX = 'data-gx-redact-'

/** `/*gx:redact:label*\/` — legal JavaScript, ignored by the engine, inert to the page. */
const SCRIPT_MARK_RE = /\/\*gx:redact:([^]*?)\*\//g

/** One withheld region, in the coordinates of the real source. */
export interface WithheldSpan {
  /** start offset in the real source */
  rawFrom: number
  /** end offset (exclusive) in the real source */
  rawTo: number
  label: string
  kind: 'content' | 'attr' | 'literal'
  /** what the mark was attached to, for the refusal message */
  where: string
}

const ATTR_RE = /([^\s"'>/=]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+))/dg

/**
 * An attribute's value range inside a start tag, in source coordinates, or null.
 *
 * The `d` flag is doing the real work here. Deriving the value's start from
 * `m[0].length - value.length` counts the quote characters, which lands a
 * character or two off and silently truncates the label — `API key` came back as
 * `PI key`. Reading the capture group's own indices is exact, and it also walks
 * the attributes in order, so a value that happens to contain a later
 * attribute's name cannot be mistaken for it.
 */
function attrValueRange(
  startTagXml: string,
  tagStart: number,
  name: string,
): [number, number] | null {
  for (const m of startTagXml.matchAll(ATTR_RE)) {
    if (m[1] !== name) continue
    // 1 = name, then one group per quoting style: "…" , '…', bare
    const group = m[2] !== undefined ? 2 : m[3] !== undefined ? 3 : 4
    const at = m.indices?.[group]
    if (!at) continue
    return [tagStart + at[0], tagStart + at[1]]
  }
  return null
}

/** the next string literal after `from`, as [openQuote, end, innerFrom, innerTo] or null */
function stringLiteralAt(
  src: string,
  from: number,
): { end: number; innerFrom: number; innerTo: number } | null {
  for (let i = from; i < src.length; i++) {
    const c = src[i]
    if (c !== '"' && c !== "'" && c !== '`') continue
    for (let j = i + 1; j < src.length; j++) {
      if (src[j] === '\\') {
        j++
        continue
      }
      if (src[j] === c) {
        return { end: j + 1, innerFrom: i + 1, innerTo: j }
      }
      if (src[j] === '\n' && c !== '`') break
    }
  }
  return null
}

/**
 * Find every withheld region in a source file.
 *
 * The parse map is what makes this reliable: it gives each element's start-tag
 * and content ranges, so a mark is read against the element it actually belongs
 * to rather than a regex guessing at the markup. It is also what keeps a
 * `/*gx:redact:*\/` inside ordinary text from being honoured — only a real
 * `<script>` body counts.
 */
export function collectWithheld(source: string, map: ParseMap): WithheldSpan[] {
  const out: WithheldSpan[] = []
  const scriptInner: Array<[number, number]> = []

  for (const e of map.elements) {
    const startTag = source.slice(e.startTag[0], e.startTag[1])

    if (e.tag === 'script') scriptInner.push([e.inner[0], e.inner[1]])

    // data-gx-redact="label" → this element's content
    const own = attrValueRange(startTag, e.startTag[0], MARK_ATTR)
    if (own) {
      out.push({
        rawFrom: e.inner[0],
        rawTo: e.inner[1],
        label: source.slice(own[0], own[1]),
        kind: 'content',
        where: `<${e.tag}> content`,
      })
    }

    // data-gx-redact-<attr>="label" → that one attribute's value. The label is
    // the MARK's value; the target attribute is what gets withheld.
    for (const name of attrNames(startTag)) {
      if (!name.startsWith(MARK_ATTR_PREFIX)) continue
      const target = name.slice(MARK_ATTR_PREFIX.length)
      if (!target) continue
      const mark = attrValueRange(startTag, e.startTag[0], name)
      const range = attrValueRange(startTag, e.startTag[0], target)
      if (!range || range[0] === range[1]) continue
      out.push({
        rawFrom: range[0],
        rawTo: range[1],
        label: mark ? source.slice(mark[0], mark[1]) : target,
        kind: 'attr',
        where: `<${e.tag} ${target}>`,
      })
    }
  }

  // /*gx:redact:label*/ inside a real <script> → the string literal after it
  SCRIPT_MARK_RE.lastIndex = 0
  for (let m = SCRIPT_MARK_RE.exec(source); m; m = SCRIPT_MARK_RE.exec(source)) {
    const inScript = scriptInner.some(([a, b]) => m!.index >= a && m!.index < b)
    if (!inScript) continue
    const lit = stringLiteralAt(source, m.index + m[0].length)
    if (!lit) continue
    out.push({
      rawFrom: lit.innerFrom,
      rawTo: lit.innerTo,
      label: m[1] ?? '',
      kind: 'literal',
      where: 'a <script> value',
    })
  }

  return dedupeSpans(out)
}

/** every attribute name in a start tag, in order */
function attrNames(startTagXml: string): string[] {
  const head = startTagXml.replace(/^<[\w:-]+/, '')
  const re = /([^\s"'>/=]+)\s*=\s*("([^"]*)"|'([^']*)'|([^\s"'=<>`]+))/g
  const names: string[] = []
  let m: RegExpExecArray | null
  while ((m = re.exec(head))) names.push(m[1]!)
  return names
}

/**
 * Keep the outermost regions only.
 *
 * A mark on an element that also has a marked attribute describes two different
 * things — the element's content and one attribute — so overlap is normal. What
 * must not happen is the same range being withheld twice, which would make the
 * model see `{{a}}{{a}}` and the guard report a duplicate.
 */
function dedupeSpans(spans: WithheldSpan[]): WithheldSpan[] {
  const sorted = [...spans].sort((a, b) => a.rawFrom - b.rawFrom || b.rawTo - a.rawTo)
  const out: WithheldSpan[] = []
  for (const s of sorted) {
    if (s.rawTo <= s.rawFrom) continue
    const inside = out.find((k) => s.rawFrom >= k.rawFrom && s.rawTo <= k.rawTo)
    if (inside) continue
    out.push(s)
  }
  return out.sort((a, b) => a.rawFrom - b.rawFrom)
}

/**
 * The model's view of the file, plus the means to get back.
 *
 * `view` is the text the model reads. `raw` is the file. Offsets translate in
 * both directions, which is what keeps `read_source`'s line numbering honest
 * after a 40-character key becomes an 11-character marker.
 */
export class RedactionProjection {
  readonly raw: string
  readonly view: string
  readonly spans: readonly WithheldSpan[]
  /** cumulative length the spans have lost by the time we reach each one */
  private readonly drift: number[]

  constructor(raw: string, spans: readonly WithheldSpan[]) {
    this.raw = raw
    this.spans = [...spans].sort((a, b) => a.rawFrom - b.rawFrom)
    const parts: string[] = []
    const drift: number[] = []
    let cursor = 0
    let shift = 0
    for (const s of this.spans) {
      parts.push(raw.slice(cursor, s.rawFrom))
      const marker = placeholderSource(s.label)
      parts.push(marker)
      // the cumulative length change for every raw offset past this span
      shift += marker.length - (s.rawTo - s.rawFrom)
      drift.push(shift)
      cursor = s.rawTo
    }
    parts.push(raw.slice(cursor))
    this.view = parts.join('')
    this.drift = drift
  }

  get empty(): boolean {
    return this.spans.length === 0
  }

  /**
   * A raw offset → where it sits in the view.
   *
   * `drift[i]` is the cumulative length change once span `i`'s replacement is in
   * place, so the shift for any offset is the drift of the last span that ends
   * at or before it. An offset *inside* a span maps to that span's marker start:
   * the withheld characters have no view position of their own.
   */
  toViewOffset(rawOffset: number): number {
    const at = this.spans.findIndex((s) => rawOffset >= s.rawFrom && rawOffset <= s.rawTo)
    if (at >= 0) {
      return this.spans[at]!.rawFrom + (at === 0 ? 0 : this.drift[at - 1]!)
    }
    let shift = 0
    for (let i = 0; i < this.spans.length; i++) {
      if (this.spans[i]!.rawTo > rawOffset) break
      shift = this.drift[i]!
    }
    return rawOffset + shift
  }

  /** A view offset → the raw offset it came from. Inside a marker, the span's end. */
  toRawOffset(viewOffset: number): number {
    let consumed = 0 // raw chars already walked
    let at = 0 // position in the view
    for (const s of this.spans) {
      const gap = s.rawFrom - consumed
      if (viewOffset < at + gap) return consumed + (viewOffset - at)
      at += gap
      const marker = placeholderSource(s.label)
      if (viewOffset < at + marker.length) return s.rawTo
      at += marker.length
      consumed = s.rawTo
    }
    return consumed + (viewOffset - at)
  }

  /** The view the model should see, for a raw byte range. */
  projectRange(rawFrom: number, rawTo: number): string {
    return this.view.slice(this.toViewOffset(rawFrom), this.toViewOffset(rawTo))
  }

  /** Every distinct label withheld, in first-seen order. */
  labels(): string[] {
    return [...new Set(this.spans.map((s) => sanitizeLabel(s.label)).filter(Boolean))]
  }
}

/** Build the projection for a source file and its parse map. */
export function buildProjection(source: string, map: ParseMap): RedactionProjection {
  return new RedactionProjection(source, collectWithheld(source, map))
}

/**
 * A label, cleaned for a comment carrier.
 *
 * Two characters more than the shared default strips, and bound here rather
 * than re-exported: a page's label lands in a `/*gx:redact:...*\/` script
 * comment or a `data-gx-redact-<attr>` value, and a `*` or `/` in either would
 * close the comment early and take the rest of the marker with it. A
 * document's label is not in a comment, and stripping these there would rename
 * a span the reader can see.
 */
export const sanitizeLabel = (raw: string): string => coreSanitizeLabel(raw, { forComment: true })

/** The form the model sees; braces make a marker recognisable in a reply. */
const OPEN = '{{'
const CLOSE = '}}'

/**
 * The literal text the model is shown in place of what was withheld.
 *
 * Not the shared one: that builds from the default sanitiser, and a page's
 * label goes through the stricter one.
 */
export function placeholderSource(label: string): string {
  return `${OPEN}${sanitizeLabel(label) || 'private'}${CLOSE}`
}

/** What this editor's markers need that another's do not. */
const HTML_SPEC: PlaceholderSpec = {
  subject: 'page',
  opening:
    'Each stands in for something the reader has deliberately withheld from you; you cannot see what is inside, and that is the point.',
  boundary: 'a line break',
  middle: [
    'A placeholder may stand in for visible text, for the value of an attribute (a key, a token, a URL), for a value inside a <script>, or for a whole comment. In every case it is text you were not given: write around it rather than guessing.',
    'A withheld attribute or script value still governs how the page behaves. Do not invent a replacement for it and do not remove the thing that carries it.',
  ],
}

export function placeholderInstruction(labels: readonly string[]): string {
  return buildInstruction(labels, HTML_SPEC)
}

export {
  MAX_LABEL_LENGTH,
  checkPlaceholders,
  collectPlaceholders,
  isWholePlaceholder,
  readPlaceholderLabel,
  type PlaceholderIssue,
} from '@genoffice/agent-core/redact-core'
