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

/** Attribute marking an element's own content. */
export const MARK_ATTR = 'data-gx-redact'
/** Prefix for marking one attribute's value: `data-gx-redact-src="label"`. */
export const MARK_ATTR_PREFIX = 'data-gx-redact-'

/** `/*gx:redact:label*\/` — legal JavaScript, ignored by the engine, inert to the page. */
const SCRIPT_MARK_RE = /\/\*gx:redact:([^]*?)\*\//g

/** Long enough to name what a span stands for, short enough to stay readable. */
export const MAX_LABEL_LENGTH = 40

const OPEN = '{{'
const CLOSE = '}}'

/**
 * Clean a label for storage and for the model prompt.
 *
 * The character set has to be safe for **both** carriers, not just one. A label
 * goes into a `data-*` attribute value and into a comment that opens a script
 * literal, so `*` and `/` are stripped here rather than in the caller: two
 * sanitizers would drift, and a label holding both would close the comment
 * early and turn the rest of the line into code. Stripping them costs a label
 * nothing worth keeping, because a label is a name like "API key".
 */
export function sanitizeLabel(raw: string): string {
  return raw
    .replace(/[{}<>="'*/]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, MAX_LABEL_LENGTH)
    .trim()
}

/** The literal text the model is shown in place of a withheld span. */
export function placeholderSource(label: string): string {
  return `${OPEN}${sanitizeLabel(label) || 'private'}${CLOSE}`
}

export function isWholePlaceholder(text: string): boolean {
  const t = text.trim()
  return t.startsWith(OPEN) && t.endsWith(CLOSE) && t.length > OPEN.length + CLOSE.length
}

export function readPlaceholderLabel(text: string): string | null {
  if (!isWholePlaceholder(text)) return null
  return text.trim().slice(OPEN.length, -CLOSE.length)
}

export interface PlaceholderIssue {
  found: string
  expected: string
  reason: 'split' | 'missing' | 'unknown'
}

function scan(text: string): string[] {
  return [...text.matchAll(/\{\{[^{}]*\}\}/g)].map((m) => m[0])
}

/** Braces a malformed marker left behind, with well-formed ones masked out first. */
function leftoverBraces(text: string): string[] {
  const masked = text.replace(/\{\{[^{}]*\}\}/g, (m) => ' '.repeat(m.length))
  const out: string[] = []
  for (const m of masked.matchAll(/\{\{|\}\}/g)) out.push(m[0])
  for (const m of masked.matchAll(/(?<!\{)\{[^{}]*\}(?!\})/g)) out.push(m[0])
  return out
}

/** Every marker in a piece of text, in order. */
export function collectPlaceholders(text: string): string[] {
  return scan(text)
}

/**
 * Compare the markers in a model's answer against the ones it was given.
 *
 * Order is ignored — a model may rewrite the sentence — but the count and the
 * exact spelling are not. Accepting anything else writes a mangled marker into a
 * file that still opens, so the damage is silent.
 */
export function checkPlaceholders(before: string, after: string): PlaceholderIssue[] {
  const issues: PlaceholderIssue[] = []
  for (const stray of leftoverBraces(after)) {
    issues.push({ found: stray, expected: '', reason: 'split' })
  }
  const got = new Map<string, number>()
  for (const m of scan(after)) got.set(m, (got.get(m) ?? 0) + 1)
  const want = new Map<string, number>()
  for (const m of scan(before)) want.set(m, (want.get(m) ?? 0) + 1)
  for (const [marker, n] of want) {
    if ((got.get(marker) ?? 0) < n) {
      issues.push({
        found: (got.get(marker) ?? 0) ? marker : '',
        expected: marker,
        reason: 'missing',
      })
    }
  }
  for (const [marker, n] of got) {
    if (!want.has(marker)) issues.push({ found: marker, expected: '', reason: 'unknown' })
    else if (n > (want.get(marker) ?? 0))
      issues.push({ found: marker, expected: marker, reason: 'unknown' })
  }
  return issues
}

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

/** The instruction a model gets, naming the placeholders this file actually has. */
export function placeholderInstruction(labels: readonly string[]): string {
  const list = [...new Set(labels)].map((l) => `- ${placeholderSource(l)}`).join('\n')
  return [
    '## Private placeholders',
    'This page contains {{...}} placeholders. Each stands in for something the reader has deliberately withheld from you; you cannot see what is inside, and that is the point.',
    '',
    'Treat every placeholder as one indivisible object:',
    '- Copy it character for character — same letters, same order, same spacing.',
    '- Never split it across a line break or put a space inside it.',
    '- Never merge two into one, never split one into several, never reorder them.',
    '- Never rename, translate, re-case, expand or shorten it.',
    '- Never drop one, and never add a placeholder that was not already there.',
    '',
    'A placeholder may stand in for visible text, for the value of an attribute (a key, a token, a URL), for a value inside a <script>, or for a whole comment. In every case it is text you were not given: write around it rather than guessing.',
    'A withheld attribute or script value still governs how the page behaves. Do not invent a replacement for it and do not remove the thing that carries it.',
    '',
    'The placeholders in this page:',
    list,
  ].join('\n')
}
