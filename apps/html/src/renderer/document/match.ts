/**
 * Tolerant search for a model-supplied `old` snippet in HTML source. Runs a
 * ladder of increasingly lenient comparisons and stops at the first rung that
 * yields exactly one candidate; several candidates at any rung is ambiguity,
 * never "take the first".
 */
export type MatchFailure =
  | { kind: 'not_found'; nearest: { from: number; to: number; similarity: number } | null }
  | { kind: 'ambiguous'; candidates: number[] }

export type MatchResult =
  { ok: true; from: number; to: number; rung: number } | ({ ok: false } & MatchFailure)

export interface MatchOptions {
  /** restrict the search to this range of the text */
  within?: [number, number]
  /** several hits are acceptable (replace_all) */
  allowMultiple?: boolean
}

function findAll(hay: string, needle: string, offset: number): number[] {
  const out: number[] = []
  if (!needle) return out
  let i = hay.indexOf(needle)
  while (i >= 0) {
    out.push(i + offset)
    i = hay.indexOf(needle, i + 1)
  }
  return out
}

/** collapse each line's trailing whitespace; returns the text plus an index map back to the original */
function stripTrailingWs(text: string): { text: string; map: number[] } {
  const map: number[] = []
  let out = ''
  const lines = text.split('\n')
  let pos = 0
  lines.forEach((line, i) => {
    const trimmed = line.replace(/[ \t]+$/, '')
    for (let k = 0; k < trimmed.length; k++) {
      map.push(pos + k)
      out += trimmed[k]
    }
    pos += line.length
    if (i < lines.length - 1) {
      map.push(pos)
      out += '\n'
      pos += 1
    }
  })
  map.push(pos)
  return { text: out, map }
}

/** strip leading indentation of every line (the caller re-indents `new` itself) */
function stripLeadingWs(text: string): { text: string; map: number[] } {
  const map: number[] = []
  let out = ''
  const lines = text.split('\n')
  let pos = 0
  lines.forEach((line, i) => {
    const lead = line.length - line.replace(/^[ \t]+/, '').length
    for (let k = lead; k < line.length; k++) {
      map.push(pos + k)
      out += line[k]
    }
    pos += line.length
    if (i < lines.length - 1) {
      map.push(pos)
      out += '\n'
      pos += 1
    }
  })
  map.push(pos)
  return { text: out, map }
}

const ENTITY_MAP: Record<string, string> = {
  '&amp;': '&',
  '&quot;': '"',
  '&apos;': "'",
  '&#39;': "'",
  '&nbsp;': ' ',
  '\u00a0': ' ',
  '‘': "'",
  '’': "'",
  '“': '"',
  '”': '"',
  '–': '-',
  '—': '-',
}
const ENTITY_RE = /&amp;|&quot;|&apos;|&#39;|&nbsp;|[\u00a0‘’“”–—]/g

/** canonical text: entities and typographic quotes/dashes folded, whitespace runs collapsed */
function canonical(text: string): { text: string; map: number[] } {
  const map: number[] = []
  let out = ''
  let i = 0
  while (i < text.length) {
    ENTITY_RE.lastIndex = i
    const m = ENTITY_RE.exec(text)
    if (m && m.index === i) {
      const rep = ENTITY_MAP[m[0]]!
      map.push(i)
      out += rep
      i += m[0].length
      continue
    }
    const ch = text[i]!
    if (/\s/.test(ch)) {
      if (!out.endsWith(' ')) {
        map.push(i)
        out += ' '
      }
      i++
      continue
    }
    map.push(i)
    out += ch
    i++
  }
  map.push(text.length)
  return { text: out, map }
}

function levenshteinRatio(a: string, b: string): number {
  if (a === b) return 1
  if (!a.length || !b.length) return 0
  const prev = new Array<number>(b.length + 1)
  const cur = new Array<number>(b.length + 1)
  for (let j = 0; j <= b.length; j++) prev[j] = j
  for (let i = 1; i <= a.length; i++) {
    cur[0] = i
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(
        prev[j]! + 1,
        cur[j - 1]! + 1,
        prev[j - 1]! + (a[i - 1] === b[j - 1] ? 0 : 1),
      )
    }
    for (let j = 0; j <= b.length; j++) prev[j] = cur[j]!
  }
  return 1 - prev[b.length]! / Math.max(a.length, b.length)
}

interface Rung {
  name: string
  normalize: (text: string) => { text: string; map: number[] }
}

const RUNGS: Rung[] = [
  {
    name: 'exact',
    normalize: (t) => ({ text: t, map: Array.from({ length: t.length + 1 }, (_, i) => i) }),
  },
  { name: 'trailing-ws', normalize: stripTrailingWs },
  { name: 'indent', normalize: (t) => stripLeadingWs(stripTrailingWs(t).text) },
  { name: 'canonical', normalize: canonical },
]

/** `indent` rung composes two normalizations; rebuild its map against the original text */
function normalizeWithMap(rung: Rung, text: string): { text: string; map: number[] } {
  if (rung.name !== 'indent') return rung.normalize(text)
  const a = stripTrailingWs(text)
  const b = stripLeadingWs(a.text)
  return { text: b.text, map: b.map.map((i) => a.map[i]!) }
}

export function findSnippet(text: string, old: string, options: MatchOptions = {}): MatchResult {
  const [lo, hi] = options.within ?? [0, text.length]
  const hay = text.slice(lo, hi)
  if (!old) return { ok: false, kind: 'not_found', nearest: null }

  for (let r = 0; r < RUNGS.length; r++) {
    const rung = RUNGS[r]!
    const h = normalizeWithMap(rung, hay)
    const n = normalizeWithMap(rung, old).text.trim()
    if (!n) continue
    const hits = findAll(h.text, n, 0)
    if (hits.length === 0) continue
    if (hits.length > 1 && !options.allowMultiple) {
      return { ok: false, kind: 'ambiguous', candidates: hits.map((i) => lo + h.map[i]!) }
    }
    const i = hits[0]!
    return { ok: true, from: lo + h.map[i]!, to: lo + h.map[i + n.length]!, rung: r }
  }

  // block anchor: first and last lines exact (trimmed), interior similar enough
  const oldLines = old
    .split('\n')
    .map((l) => l.trim())
    .filter((l, i, arr) => l || (i > 0 && i < arr.length - 1))
  if (oldLines.length >= 3) {
    const first = oldLines[0]!
    const last = oldLines[oldLines.length - 1]!
    const hayLines = hay.split('\n')
    const starts: number[] = []
    let pos = 0
    for (const line of hayLines) {
      starts.push(pos)
      pos += line.length + 1
    }
    const candidates: Array<{ from: number; to: number; sim: number }> = []
    for (let i = 0; i < hayLines.length; i++) {
      if (hayLines[i]!.trim() !== first) continue
      const j = i + oldLines.length - 1
      if (j >= hayLines.length || hayLines[j]!.trim() !== last) continue
      const interior = hayLines
        .slice(i + 1, j)
        .map((l) => l.trim())
        .join('\n')
      const sim = levenshteinRatio(interior, oldLines.slice(1, -1).join('\n'))
      if (sim >= 0.8) {
        const leadI = hayLines[i]!.length - hayLines[i]!.trimStart().length
        const trailJ = hayLines[j]!.length - hayLines[j]!.trimEnd().length
        candidates.push({
          from: starts[i]! + leadI,
          to: starts[j]! + hayLines[j]!.length - trailJ,
          sim,
        })
      }
    }
    if (candidates.length === 1) {
      const c = candidates[0]!
      return { ok: true, from: lo + c.from, to: lo + c.to, rung: RUNGS.length }
    }
    if (candidates.length > 1) {
      return { ok: false, kind: 'ambiguous', candidates: candidates.map((c) => lo + c.from) }
    }
  }

  return { ok: false, kind: 'not_found', nearest: nearestSnippet(hay, old, lo) }
}

/** best-similarity window of the same line count as `old`, for "did you mean" feedback */
function nearestSnippet(hay: string, old: string, offset: number) {
  const oldLines = old.split('\n')
  const hayLines = hay.split('\n')
  if (hayLines.length === 0) return null
  const starts: number[] = []
  let pos = 0
  for (const line of hayLines) {
    starts.push(pos)
    pos += line.length + 1
  }
  let best: { from: number; to: number; similarity: number } | null = null
  const target = oldLines.map((l) => l.trim()).join('\n')
  const window = Math.min(oldLines.length, hayLines.length)
  // cap the scan for very large documents
  const maxStarts = Math.min(hayLines.length - window + 1, 4000)
  for (let i = 0; i < maxStarts; i++) {
    const slice = hayLines
      .slice(i, i + window)
      .map((l) => l.trim())
      .join('\n')
    if (Math.abs(slice.length - target.length) > Math.max(40, target.length)) continue
    const sim = levenshteinRatio(slice, target)
    if (!best || sim > best.similarity) {
      best = {
        from: offset + starts[i]!,
        to: offset + starts[i + window - 1]! + hayLines[i + window - 1]!.length,
        similarity: sim,
      }
    }
  }
  return best && best.similarity >= 0.6 ? best : null
}

export function lineOf(text: string, offset: number): number {
  let line = 1
  for (let i = 0; i < offset && i < text.length; i++) if (text.charCodeAt(i) === 10) line++
  return line
}
