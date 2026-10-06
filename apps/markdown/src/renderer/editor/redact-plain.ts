/**
 * Withholding values in the source-text surface: a `.txt` or `.json` opened by
 * the markdown app.
 *
 * The block-editor mark (a `<span data-redaction>`) does not work here. It is
 * HTML, and these files are not HTML — and worse, a comment mark is outright
 * illegal in JSON, so the obvious `// gx:redact:label` would leave the reader
 * with a file that no parser can read. Losing the file to hide it from a model
 * is the worse failure by a wide margin, so each format gets the mark it can
 * carry:
 *
 * | surface | mark | why |
 * |---|---|---|
 * | plain text | `// gx:redact:label` on the line above the value | travels with its value; a reader marking an `.env` is choosing to add a visible line to a file that has no comment syntax of its own |
 * | JSON | a root `__gxRedact` array of `{ path, label }` | the only in-band form that keeps the file valid and every real value in place |
 *
 * As everywhere else, the value is never replaced. The mark records what to
 * hide; the projection is what the model reads.
 */
import { placeholderSource, sanitizeLabel } from './redact'

/** A mark line, on its own, immediately above the value it covers. */
const PLAIN_MARK_RE = /^[ \t]*\/\/[ \t]*gx:redact:(.+?)[ \t]*$/gm

/** The root key JSON marks live under. */
export const JSON_REDACT_KEY = '__gxRedact'

/** `// gx:redact:<label>` as a whole line, newline included. */
export function plainMarkLine(label: string): string {
  return `// gx:redact:${sanitizeLabel(label) || 'private'}\n`
}

/** Every label a plain-text file carries, in document order. */
export function plainLabels(text: string): string[] {
  const out: string[] = []
  for (const match of text.matchAll(PLAIN_MARK_RE)) {
    const label = sanitizeLabel(match[1] ?? '')
    out.push(label || 'private')
  }
  return out
}

/** Is there a mark anywhere in this text? */
export function hasPlainMark(text: string): boolean {
  PLAIN_MARK_RE.lastIndex = 0
  const found = PLAIN_MARK_RE.test(text)
  PLAIN_MARK_RE.lastIndex = 0
  return found
}

/**
 * Insert the mark above the line holding `from`.
 *
 * Line-relative, not offset-relative, so the mark follows its value when the
 * reader edits above it — the same property the editor-side marks have. The
 * selection is collapsed to its start: a mark covers one line either way, and
 * taking the selection's first line is what a reader means by "this one".
 */
export function markPlainRange(text: string, from: number, _to: number, label: string): string {
  const at = Math.max(0, Math.min(from, text.length))
  const lineStart = text.lastIndexOf('\n', at - 1) + 1
  return text.slice(0, lineStart) + plainMarkLine(label) + text.slice(lineStart)
}

/**
 * What a model reads instead of the file.
 *
 * The line under each mark becomes the placeholder; the mark itself goes too,
 * since `gx:redact:<label>` is a note to the reader, not content. Everything
 * else is byte-identical.
 */
export function projectPlain(text: string): string {
  const lines = text.split('\n')
  const out: string[] = []
  let pending: string | null = null
  for (const line of lines) {
    const mark = /^[ \t]*\/\/[ \t]*gx:redact:(.+?)[ \t]*$/.exec(line)
    if (mark) {
      // the mark itself is a note to the reader, not content: it goes too
      pending = sanitizeLabel(mark[1] ?? '') || 'private'
      continue
    }
    if (pending !== null) {
      out.push(placeholderSource(pending))
      pending = null
      continue
    }
    out.push(line)
  }
  // a mark with nothing under it: keep the file's shape rather than inventing
  // a value for it
  if (pending !== null) out.push(placeholderSource(pending))
  return out.join('\n')
}

// ── JSON ────────────────────────────────────────────────────────────────────

interface JsonMark {
  readonly path: string
  readonly label: string
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/** RFC 6901 escaping, so a key containing `/` or `~` still round-trips. */
export function encodePointerSegment(segment: string): string {
  return segment.replace(/~/g, '~0').replace(/\//g, '~1')
}

export function decodePointerSegment(segment: string): string {
  return segment.replace(/~1/g, '/').replace(/~0/g, '~')
}

export function jsonPointer(segments: readonly string[]): string {
  return segments.map((segment) => `/${encodePointerSegment(segment)}`).join('')
}

function readMarks(root: unknown): JsonMark[] {
  if (!isRecord(root)) return []
  const raw = root[JSON_REDACT_KEY]
  if (!Array.isArray(raw)) return []
  const out: JsonMark[] = []
  for (const entry of raw) {
    if (!isRecord(entry)) continue
    if (typeof entry.path !== 'string' || entry.path === '') continue
    const label = sanitizeLabel(typeof entry.label === 'string' ? entry.label : '')
    out.push({ path: entry.path, label: label || 'private' })
  }
  return out
}

/** The labels a parsed JSON document carries. */
export function jsonLabels(json: unknown): string[] {
  return readMarks(json).map((mark) => mark.label)
}

/** Can this document carry marks at all? */
export function jsonCanHoldMarks(json: unknown): boolean {
  return isRecord(json)
}

/** Does `path` address a value that is actually in `document`? */
export function jsonPathResolves(document: unknown, path: readonly string[]): boolean {
  return resolvesIn(document, path)
}

/** Add a mark for `segments`, keeping every real value in place. */
export function markJsonPath(
  json: unknown,
  segments: readonly string[],
  label: string,
): { ok: true; json: unknown } | { ok: false; reason: string } {
  if (!isRecord(json)) {
    // No root object means nowhere to record a mark that stays valid JSON.
    return { ok: false, reason: 'the document has no root object to hold the mark' }
  }
  if (segments.length === 0 || !resolvesIn(json, segments)) {
    // A pointer that addresses nothing is worse than no mark: the writer would
    // create the key, so the model would see a value the file never had while
    // the real one stayed visible.
    return { ok: false, reason: 'that value is not in the document' }
  }
  const path = jsonPointer(segments)
  if (readMarks(json).some((mark) => mark.path === path)) {
    return { ok: false, reason: 'that value is already withheld' }
  }
  const next = { ...json }
  next[JSON_REDACT_KEY] = [...readMarks(json), { path, label: sanitizeLabel(label) || 'private' }]
  return { ok: true, json: next }
}

/** Drop the mark for one path. */
export function clearJsonPath(
  json: unknown,
  segments: readonly string[],
): { ok: true; json: unknown } | { ok: false; reason: string } {
  if (!isRecord(json)) return { ok: false, reason: 'the document has no root object' }
  const path = jsonPointer(segments)
  const marks = readMarks(json)
  const kept = marks.filter((mark) => mark.path !== path)
  if (kept.length === marks.length) return { ok: false, reason: 'that value is not withheld' }
  const next = { ...json }
  if (kept.length === 0) delete next[JSON_REDACT_KEY]
  else next[JSON_REDACT_KEY] = kept
  return { ok: true, json: next }
}

function setAtPointer(root: unknown, segments: readonly string[], value: unknown): boolean {
  let node = root
  for (let i = 0; i < segments.length - 1; i += 1) {
    const next = (node as Record<string, unknown>)[segments[i]!]
    if (!isRecord(next) && !Array.isArray(next)) return false
    node = next as unknown as Record<string, unknown>
  }
  if (Array.isArray(node)) {
    const index = Number(segments[segments.length - 1])
    if (!Number.isInteger(index) || index < 0 || index >= node.length) return false
    node[index] = value
    return true
  }
  if (!isRecord(node)) return false
  const last = segments[segments.length - 1]
  if (!(last! in node)) return false
  node[last!] = value
  return true
}

/**
 * What a model reads instead of the parsed document.
 *
 * Marked values become placeholders and the mark array itself is removed, so
 * the model sees the document's shape with those values withheld — not a
 * document carrying instructions about itself.
 *
 * A mark whose pointer addresses nothing withholds the **whole** document. A
 * partial view would mean quietly showing the model a value a mark was
 * supposed to cover, and the one thing this projection must never do is
 * return less than it promises.
 */
export function projectJson(json: unknown): unknown {
  const marks = readMarks(json)
  if (marks.length === 0) return json
  const next = structuredClone(json) as unknown
  for (const mark of marks) {
    const segments = mark.path
      .split('/')
      .slice(1)
      .map(decodePointerSegment)
      .filter((segment) => segment !== '')
    if (segments.length === 0) return withheldDocument(next)
    if (!setAtPointer(next, segments, placeholderSource(mark.label))) {
      return withheldDocument(next)
    }
  }
  if (isRecord(next)) delete next[JSON_REDACT_KEY]
  return next
}

/** The fail-closed fallback: every string in the document withheld. */
function withheldDocument(json: unknown): unknown {
  const blank = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(blank)
    if (isRecord(value)) {
      const out: Record<string, unknown> = {}
      for (const [key, entry] of Object.entries(value)) {
        // the mark list itself is dropped, as in the ordinary projection
        if (key === JSON_REDACT_KEY) continue
        out[key] = blank(entry)
      }
      return out
    }
    return typeof value === 'string' ? placeholderSource('private') : value
  }
  return blank(json)
}

/**
 * The JSON path of the value that starts on the line holding `offset`.
 *
 * The dialog asks for a name, not for a path, and the reader picked a
 * selection rather than a key — so the mark is placed on the value that begins
 * where their selection begins.
 *
 * The **full** path from the root, not just the key on that line. Resolving
 * only the key records a pointer that matches nothing: in
 * `{"db":{"password":…}}` the `"password"` line would record `/password`, and
 * the writer would then create that key at the root. The mark would protect
 * nothing *and* hand the model a value the file never had — the secret still
 * visible, plus a fabricated one. That is why an unresolvable line returns
 * null rather than a guess.
 *
 * The path is rebuilt from the file's own indentation, so it holds for any
 * consistently indented JSON rather than only for what this module writes.
 */
export function jsonPathAtLine(json: unknown, text: string, offset: number): string[] | null {
  if (!isRecord(json)) return null
  const lineStart = text.lastIndexOf('\n', Math.max(0, offset - 1)) + 1
  const lineEnd = text.indexOf('\n', lineStart)
  const line = text.slice(lineStart, lineEnd < 0 ? undefined : lineEnd)
  const KEY_LINE = /^( *?)"((?:[^"\\]|\\.)*)"\s*:/
  const opening = KEY_LINE.exec(line)
  if (!opening) return null
  const rest = line.slice(opening[0].length).trim()
  // A line that only opens a container is not a value of its own.
  if (!rest || rest.startsWith('{') || rest.startsWith('[')) return null

  // Rebuild the nesting stack from the keys, in the order the file shows them.
  const stack: Array<{ indent: number; key: string }> = []
  for (const raw of text.split('\n')) {
    const match = KEY_LINE.exec(raw)
    if (!match) continue
    const indent = match[1]!.length
    while (stack.length > 0 && stack[stack.length - 1]!.indent >= indent) stack.pop()
    stack.push({ indent, key: match[2]! })
    if (raw === line) {
      const path = stack.map((entry) => entry.key)
      // A key the parsed document does not carry is not a value to withhold.
      return resolvesIn(json, path) ? path : null
    }
  }
  return null
}

/** Does `path` actually address a value in `document`? */
function resolvesIn(document: unknown, path: readonly string[]): boolean {
  let node: unknown = document
  for (const segment of path) {
    if (Array.isArray(node)) {
      const index = Number(segment)
      if (!Number.isInteger(index) || index < 0 || index >= node.length) return false
      node = node[index]
      continue
    }
    if (!isRecord(node) || !(segment in node)) return false
    node = node[segment]
  }
  return true
}
