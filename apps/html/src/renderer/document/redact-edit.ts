import { MARK_ATTR, MARK_ATTR_PREFIX, collectWithheld, sanitizeLabel } from '../ai/redact'
import { rawToDecoded, type HtmlOp } from './ops'
import { elementCovering, type ElementEntry, type ParseMap } from './parse-map'
import type { Patch } from './patch'

/**
 * Turning a source selection into the mark that withholds it.
 *
 * `ai/redact.ts` owns the question "what does the model not get to see"; this file
 * owns the other half — "given the characters the reader selected, which of the
 * three marks describes them, and what is the smallest edit that writes it". The
 * split matters: whether a selection is withheld is decided by that module's
 * projection, so the plan produced here is checked against it rather than
 * reimplementing it.
 *
 * ## Why the source view and not the preview
 *
 * A page's secrets are usually not selectable text. They sit in `<script>`
 * strings, in `content=` / `src=` / `onclick=` values, and in comments — none of
 * which a rendered page can put a caret in. The source pane can select all of
 * them, so the menu lives there and the preview is deliberately not offered.
 *
 * ## The three marks, and which selection reaches which
 *
 * | the selection is in | mark | how it is written |
 * |---|---|---|
 * | one attribute's value in a start tag | `data-gx-redact-<name>` | `set_attr` |
 * | a string literal in a `<script>` | `/*gx:redact:label*\/` | a zero-width `Patch` |
 * | one text node | `data-gx-redact` on a new `<span>` | `wrap_text` |
 * | anything else inside one element | `data-gx-redact` on it | `set_attr` |
 *
 * The first three are the shapes the engine can describe. The fourth is the
 * honest fallback for a selection that crosses markup: it withholds more than was
 * selected, so it is only reached when no finer mark fits.
 *
 * ## Why one of the four is a patch and not an op
 *
 * Every op addresses an element by sid — through its start tag, its content, one
 * of its attributes, one of its text nodes. None of them can insert at an
 * arbitrary source offset, and the `<script>` mark has to land in front of one
 * specific literal: inside the quotes the engine's own scan would find the closing
 * quote first and withhold nothing at all. So that one edit is a character-range
 * splice, the same `Patch` primitive the op layer compiles down to, applied
 * through the same door.
 */

/** what to run, once the label is known (or, for a clear, immediately) */
export interface RedactPlan {
  mode: 'mark' | 'clear'
  /** ops for the marks the op vocabulary can express; empty when `patches` carries the edit */
  ops: HtmlOp[]
  /** character-range edits for the one insertion no op addresses */
  patches: Patch[]
}

function plan(mode: 'mark' | 'clear', ops: HtmlOp[] = [], patches: Patch[] = []): RedactPlan {
  return { mode, ops, patches }
}

/** `data-gx-redact="label"` on this element: withholds its content. */
function contentMark(sid: number, label: string): RedactPlan {
  return plan('mark', [{ op: 'set_attr', sid, name: MARK_ATTR, value: label }])
}

// ── start tags ──────────────────────────────────────────────────────────────────

interface AttrHit {
  name: string
  /** the whole attribute, leading whitespace included, in source coordinates */
  from: number
  to: number
  /** the value without its quotes; equal to from/to when the attribute has none */
  valueFrom: number
  valueTo: number
  hasValue: boolean
}

/**
 * Every attribute in a start tag, with exact source ranges.
 *
 * The `d` flag is what makes this exact: a value's own capture group reports its
 * indices, whereas deriving them from the match length counts the quote
 * characters and lands a character or two off. `set_attr` writes with the same
 * regex, so reading and writing agree by construction.
 */
const ATTR_SCAN = /(\s+)([^\s"'<>/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/dgy

function startTagAttrs(xml: string, tagStart: number): AttrHit[] {
  const tagName = /^<[^\s/>]+/.exec(xml)
  if (!tagName) return []
  ATTR_SCAN.lastIndex = tagName[0].length
  const out: AttrHit[] = []
  let m: RegExpExecArray | null
  while ((m = ATTR_SCAN.exec(xml))) {
    const at = m.indices!
    const whole = at[0]!
    // the value groups in the order the pattern lists them: 3 = "…", 4 = '…', 5 = bare.
    // Each one captures the value *without* its quotes, so its own indices are
    // already the value's range — no padding, and no counting the match length.
    const group = m[3] !== undefined ? 3 : m[4] !== undefined ? 4 : m[5] !== undefined ? 5 : 0
    const hasValue = group !== 0
    const value = hasValue ? (at[group] as [number, number]) : whole
    out.push({
      name: m[2]!,
      from: tagStart + whole[0],
      to: tagStart + whole[1],
      valueFrom: tagStart + value[0],
      valueTo: tagStart + value[1],
      hasValue,
    })
  }
  ATTR_SCAN.lastIndex = 0
  return out
}

function attrsOf(source: string, e: ElementEntry): AttrHit[] {
  return startTagAttrs(source.slice(e.startTag[0], e.startTag[1]), e.startTag[0])
}

/** the innermost start tag holding the whole selection */
function startTagAt(map: ParseMap, from: number, to: number): ElementEntry | null {
  let best: ElementEntry | null = null
  for (const e of map.elements) {
    if (e.startTag[0] > from || to > e.startTag[1]) continue
    if (!best || e.startTag[1] - e.startTag[0] < best.startTag[1] - best.startTag[0]) best = e
  }
  return best
}

function scriptAt(map: ParseMap, from: number, to: number): ElementEntry | null {
  return (
    map.elements.find((e) => e.tag === 'script' && e.inner[0] <= from && to <= e.inner[1]) ?? null
  )
}

/**
 * Elements whose text is not markup: a `<span>` written inside one of them is
 * shown to the reader as text, or breaks the script outright. They never get the
 * wrapping mark; the element-level mark is the honest one.
 */
const RAW_TEXT = new Set(['script', 'style', 'textarea', 'title'])

function textNodeAt(
  map: ParseMap,
  from: number,
  to: number,
): { el: ElementEntry; index: number; node: [number, number] } | null {
  for (const e of map.elements) {
    for (let index = 0; index < e.textNodes.length; index++) {
      const node = e.textNodes[index]!
      if (node[0] <= from && to <= node[1]) return { el: e, index, node }
    }
  }
  return null
}

// ── <script> bodies ────────────────────────────────────────────────────────────

interface ScriptLiteral {
  /** the opening quote: where the mark goes, so the engine's scan finds this literal */
  quote: number
  /** the closing quote, inclusive */
  close: number
  valueFrom: number
  valueTo: number
}

/**
 * Every string literal in a `<script>` body, in order.
 *
 * Same rule the engine reads them back with: a quote opens and closes, a
 * backslash escapes the next character, and a bare newline ends a `'…'` or
 * `"…"` literal but not a template one.
 */
function scriptLiterals(src: string, from: number, to: number): ScriptLiteral[] {
  const out: ScriptLiteral[] = []
  let i = from
  while (i < to) {
    const c = src[i]!
    if (c !== '"' && c !== "'" && c !== '`') {
      i++
      continue
    }
    let j = i + 1
    for (; j < to; j++) {
      if (src[j] === '\\') {
        j++
        continue
      }
      if (src[j] === c) break
      if (src[j] === '\n' && c !== '`') break
    }
    // unterminated: nothing after this point can be read back as a literal
    if (j >= to) break
    out.push({ quote: i, close: j, valueFrom: i + 1, valueTo: j })
    i = j + 1
  }
  return out
}

const SCRIPT_MARK = /\/\*gx:redact:[^]*?\*\//g

function scriptMarks(src: string, from: number, to: number): Array<{ start: number; end: number }> {
  const out: Array<{ start: number; end: number }> = []
  SCRIPT_MARK.lastIndex = from
  for (let m = SCRIPT_MARK.exec(src); m && m.index < to; m = SCRIPT_MARK.exec(src)) {
    out.push({ start: m.index, end: m.index + m[0].length })
  }
  SCRIPT_MARK.lastIndex = 0
  return out
}

/**
 * The label as it can live inside a JavaScript comment.
 *
 * `sanitizeLabel` makes a label safe to store in an attribute, which is not the
 * same job: one `*` or `/` next to the closing `*\/` would end the comment early
 * and turn the rest of the label into code. A label is the reader's own typing, so
 * it is reduced rather than refused.
 */
// sanitizeLabel already strips the characters that could close a comment, so a
// label is safe in both carriers. Re-reducing it here would be a second rule to
// keep in step with the sanitizer.
const commentLabel = (label: string): string => label

/** the literal a selection lands on: the one it touches, else the next one after it */
function literalFor(
  lits: readonly ScriptLiteral[],
  from: number,
  to: number,
): ScriptLiteral | null {
  return (
    lits.find((l) => l.quote < to && from <= l.close) ?? lits.find((l) => l.quote >= from) ?? null
  )
}

// ── the plan ───────────────────────────────────────────────────────────────────

/**
 * The edit that withholds `[from, to)`, or null when no mark describes it.
 *
 * The two fallbacks both withhold *more* than was selected — the element's whole
 * content, or the whole `<script>` when the selection was code with no literal to
 * anchor to — and both are deliberate. The reader asked for these characters not
 * to reach a model, and quietly handing them over because the ideal mark was not
 * available is the one failure this feature exists to prevent. A null is
 * reserved for the case where nothing at all can be written: a void element, which
 * has no content, and a label that sanitizes away.
 */
export function redactMarkPlan(
  source: string,
  map: ParseMap,
  from: number,
  to: number,
  rawLabel: string,
): RedactPlan | null {
  const label = sanitizeLabel(rawLabel)
  if (!label || to <= from) return null

  // 1. a value inside a start tag — `src=`, `content=`, `onclick=`, `data-*=`.
  // The whole attribute counts, not only its value: in the source the natural
  // gesture is to drag over `src="…"`, name and quotes included.
  const owner = startTagAt(map, from, to)
  if (owner) {
    const hit = attrsOf(source, owner).find((a) => a.hasValue && a.from <= from && to <= a.to)
    if (hit) {
      return plan('mark', [
        { op: 'set_attr', sid: owner.sid, name: MARK_ATTR_PREFIX + hit.name, value: label },
      ])
    }
    // the tag itself rather than one of its values: withhold the element's content.
    if (owner.inner[0] === owner.inner[1]) {
      // A void element has no content, and a mark on an empty range is inert —
      // the projection drops it. What a void element *does* carry is its
      // attributes, and for an `<img>` that is the whole payload. Returning
      // null here would make the command look like it did nothing, so a
      // selection covering the whole tag withholds every value on it instead.
      const valued = attrsOf(source, owner).filter((a) => a.valueFrom !== a.valueTo)
      if (valued.length === 0) return null
      return plan(
        'mark',
        valued.map((a) => ({
          op: 'set_attr' as const,
          sid: owner.sid,
          name: MARK_ATTR_PREFIX + a.name,
          value: label,
        })),
      )
    }
    return contentMark(owner.sid, label)
  }

  // 2. a value inside a <script>: the comment goes in front of the literal
  const script = scriptAt(map, from, to)
  if (script) {
    const lit = literalFor(scriptLiterals(source, script.inner[0], script.inner[1]), from, to)
    if (!lit) return contentMark(script.sid, label)
    return plan(
      'mark',
      [],
      [{ from: lit.quote, to: lit.quote, text: `/*gx:redact:${commentLabel(label)}*/` }],
    )
  }

  // 3. one text node: wrap the range, addressed in decoded offsets
  const tn = textNodeAt(map, from, to)
  if (tn && !RAW_TEXT.has(tn.el.tag)) {
    const raw = source.slice(tn.node[0], tn.node[1])
    return plan('mark', [
      {
        op: 'wrap_text',
        sid: tn.el.sid,
        index: tn.index,
        start: rawToDecoded(raw, from - tn.node[0]),
        end: rawToDecoded(raw, to - tn.node[0]),
        tag: 'span',
        attrs: { [MARK_ATTR]: label },
      },
    ])
  }

  // 4. anything else inside a single element: withhold its content
  const covering = elementCovering(map, from, to)
  return covering ? contentMark(covering.sid, label) : null
}

/**
 * The edit that *un*-withholds `[from, to)`, or null when nothing is withheld
 * there. Choosing this over a fresh mark is what makes a second visit to an
 * already-marked region undo it.
 *
 * "Something is withheld here" is asked of the projection, never re-derived: the
 * spans come from `collectWithheld`, and this only locates the mark that produced
 * one so it can be taken off again.
 */
export function redactClearPlan(
  source: string,
  map: ParseMap,
  from: number,
  to: number,
): RedactPlan | null {
  const spans = collectWithheld(source, map)
  if (to <= from) return null

  // Content: an element carrying data-gx-redact that the selection touches. The
  // test is overlap, not containment, so selecting the wrapper the mark came with
  // clears it too — but a selection that stops at the start tag is left alone.
  const content = spans.find((s) => s.kind === 'content' && from < s.rawTo && s.rawFrom < to)
  if (content) {
    const el = map.elements.find(
      (e) => e.inner[0] === content.rawFrom && attrsOf(source, e).some((a) => a.name === MARK_ATTR),
    )
    if (!el) return null
    // A wrapper this feature created goes away whole; anything the author wrote
    // only loses the attribute. Selecting the element's own tags is what asks for
    // the first, and it cannot silently strip styling off a <p> or an <img>.
    if (el.tag === 'span' && from <= el.range[0] && to >= el.range[1]) {
      return plan('clear', [{ op: 'unwrap', sid: el.sid }])
    }
    return plan('clear', [{ op: 'set_attr', sid: el.sid, name: MARK_ATTR, value: null }])
  }

  // An attribute value: the mark goes, the attribute stays. Overlap again, so that
  // dragging over `src="…"` clears what marking `src="…"` wrote.
  const attr = spans.find((s) => s.kind === 'attr' && from < s.rawTo && s.rawFrom < to)
  if (attr) {
    for (const e of map.elements) {
      const attrs = attrsOf(source, e)
      for (const mark of attrs) {
        if (!mark.name.startsWith(MARK_ATTR_PREFIX)) continue
        const target = attrs.find((a) => a.name === mark.name.slice(MARK_ATTR_PREFIX.length))
        if (target && target.valueFrom === attr.rawFrom && target.valueTo === attr.rawTo) {
          return plan('clear', [{ op: 'set_attr', sid: e.sid, name: mark.name, value: null }])
        }
      }
    }
    return null
  }

  // a <script> literal: the comment is the mark, so the comment is what goes.
  // The test is deliberately loose about the quotes — selecting the literal with
  // them is the same gesture as selecting the value inside them.
  if (spans.some((s) => s.kind === 'literal' && s.rawFrom <= to && from <= s.rawTo)) {
    for (const e of map.elements) {
      if (e.tag !== 'script') continue
      const lits = scriptLiterals(source, e.inner[0], e.inner[1])
      for (const mark of scriptMarks(source, e.inner[0], e.inner[1])) {
        const lit = lits.find((l) => l.quote >= mark.end)
        if (lit && lit.quote < to && from <= lit.close) {
          return plan('clear', [], [{ from: mark.start, to: mark.end, text: '' }])
        }
      }
    }
  }

  return null
}
