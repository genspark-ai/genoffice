import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  type MarkedToken,
  resolveMarkdownStyle,
  type StyleHints,
  styleHintsFromTokens,
  withMarkdownStyle,
} from './markdownStyle'

/**
 * Block-level lossless save.
 *
 * `getMarkdown()` re-serializes the whole document, so a one-word edit rewrites
 * every block: list markers, emphasis style, escapes and blank lines are
 * normalized, and anything the schema cannot hold (link reference definitions,
 * footnote definitions, raw HTML blocks) is dropped or rewritten. That turns a
 * small edit into a screen of unrelated diff for anyone keeping notes in git.
 *
 * Instead, the source is split into the top-level blocks marked produced when
 * the file was opened, each remembered together with the editor nodes it
 * became. On save, blocks whose nodes are unchanged are copied from the source
 * verbatim; only the runs of nodes that changed go through the serializer.
 */

interface Unit {
  /** the block's own text, without the blank lines that follow it */
  core: string
  /** blank lines and node-less blocks (reference definitions) up to the next block */
  glue: string
  /** whether `core` is blank lines that became empty paragraphs (never needs a separator) */
  space: boolean
  /** index of the unit's first node in the flattened node list */
  first: number
  /** how many top-level nodes the block became */
  count: number
  /** the block's own conventions, for re-serializing it after an edit */
  style: StyleHints
}

export interface SourceMap {
  /** blocks in source order; the first unit is a node-less carrier for leading glue */
  units: Unit[]
  /** JSON keys of the document's top-level nodes when the source was loaded */
  keys: string[]
  /** whether the source ended with a newline */
  finalNewline: boolean
  /** the document's majority conventions, for blocks that are new */
  style: StyleHints
}

/** same split @tiptap/markdown applies before parsing (its extractAbsorbedBlankLines) */
const TRAILING_BLANK_LINES = /\n[^\S\n]*(?:\n[^\S\n]*)+$/
const MAX_LCS_CELLS = 4_000_000

function isEmptyParagraph(node: PmNode): boolean {
  return node.type.name === 'paragraph' && node.childCount === 0
}

function nodeKey(node: PmNode): string {
  return JSON.stringify(node.toJSON())
}

function lex(editor: Editor, source: string): MarkedToken[] | null {
  const manager = editor.markdown
  if (!manager) return null
  const marked = manager.instance as unknown as {
    Lexer: new (options: unknown) => { lex(src: string): MarkedToken[] }
    defaults: unknown
  }
  try {
    return new marked.Lexer(marked.defaults).lex(source)
  } catch {
    return null
  }
}

/** mirror of the manager's normalization: trailing blank lines become their own space token */
function splitAbsorbedBlankLines(tokens: MarkedToken[]): MarkedToken[] {
  return tokens.flatMap((token, index) => {
    if (token.type === 'space' || tokens[index + 1]?.type === 'space') return [token]
    const trailing = TRAILING_BLANK_LINES.exec(token.raw)
    if (!trailing) return [token]
    return [
      { ...token, raw: token.raw.slice(0, -trailing[0].length) },
      { type: 'space', raw: trailing[0] },
    ]
  })
}

/** how many empty paragraphs the manager makes of a space token between blocks */
function implicitParagraphs(raw: string, hasPrevious: boolean, hasNext: boolean): number {
  const separators = (raw.replace(/\r\n/g, '\n').match(/\n\n/g) ?? []).length
  return Math.max(separators - (hasPrevious && hasNext ? 1 : 0), 0)
}

/**
 * Split `source` into blocks and pair them with `doc`'s top-level nodes.
 * Returns null when the pairing cannot be established (tokens that do not tile
 * the source, a block that parses differently on its own); callers then save
 * through the plain serializer.
 */
export function buildSourceMap(editor: Editor, doc: PmNode, source: string): SourceMap | null {
  const manager = editor.markdown
  if (!manager) return null
  const rawTokens = lex(editor, source)
  if (!rawTokens) return null
  let offset = 0
  for (const token of rawTokens) {
    if (!source.startsWith(token.raw, offset)) return null
    offset += token.raw.length
  }
  if (offset !== source.length) return null

  const tokens = splitAbsorbedBlankLines(rawTokens)
  const nonSpace = tokens.map((t) => t.type !== 'space')
  // a block parsed alone still needs the `[id]: url` definitions, or `![alt][id]` stays text
  const definitions = tokens
    .filter((t) => t.type === 'def')
    .map((t) => t.raw.trimEnd())
    .join('\n')
  const units: Unit[] = [{ core: '', glue: '', space: false, first: 0, count: 0, style: {} }]
  const types: string[] = []
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]!
    let produced: string[]
    let space = false
    if (token.type === 'space') {
      const hasPrevious = nonSpace.slice(0, i).some(Boolean)
      const hasNext = nonSpace.slice(i + 1).some(Boolean)
      produced = new Array<string>(implicitParagraphs(token.raw, hasPrevious, hasNext)).fill(
        'paragraph',
      )
      space = true
    } else {
      let parsed
      const core = token.raw.replace(TRAILING_BLANK_LINES, '')
      const alone =
        definitions && token.type !== 'def' && core.includes(']')
          ? `${core}\n\n${definitions}`
          : core
      try {
        parsed = manager.parse(alone).content ?? []
      } catch {
        return null
      }
      produced = parsed.map((json) => String(json.type))
    }
    if (produced.length === 0) {
      units[units.length - 1]!.glue += token.raw
      continue
    }
    units.push({
      core: token.raw,
      glue: '',
      space,
      first: types.length,
      count: produced.length,
      style: space ? {} : styleHintsFromTokens([token]),
    })
    types.push(...produced)
  }

  // the editor keeps an empty paragraph after a trailing list/code/table so
  // the caret can leave it; it has no source and serializes to nothing
  const last = doc.childCount > 0 ? doc.child(doc.childCount - 1) : null
  if (last && doc.childCount === types.length + 1 && isEmptyParagraph(last)) {
    units.push({ core: '', glue: '', space: true, first: types.length, count: 1, style: {} })
    types.push('paragraph')
  }
  if (types.length !== doc.childCount) return null
  const keys: string[] = []
  for (let i = 0; i < doc.childCount; i++) {
    const child = doc.child(i)
    if (child.type.name !== types[i]) return null
    keys.push(nodeKey(child))
  }
  return {
    units,
    keys,
    finalNewline: source.endsWith('\n'),
    style: styleHintsFromTokens(rawTokens),
  }
}

/** longest common subsequence as old→new index pairs, monotonic; -1 where unmatched */
function matchNodes(oldKeys: string[], newKeys: string[]): Int32Array {
  const n = oldKeys.length
  const m = newKeys.length
  const match = new Int32Array(n).fill(-1)
  let prefix = 0
  while (prefix < n && prefix < m && oldKeys[prefix] === newKeys[prefix]) {
    match[prefix] = prefix
    prefix++
  }
  let suffix = 0
  while (
    suffix < n - prefix &&
    suffix < m - prefix &&
    oldKeys[n - 1 - suffix] === newKeys[m - 1 - suffix]
  ) {
    match[n - 1 - suffix] = m - 1 - suffix
    suffix++
  }
  const rows = n - prefix - suffix
  const cols = m - prefix - suffix
  if (rows === 0 || cols === 0 || rows * cols > MAX_LCS_CELLS) return match
  // classic DP over the middle; the table is small after trimming
  const width = cols + 1
  const table = new Int32Array((rows + 1) * width)
  for (let i = rows - 1; i >= 0; i--) {
    const oldKey = oldKeys[prefix + i]
    for (let j = cols - 1; j >= 0; j--) {
      table[i * width + j] =
        oldKey === newKeys[prefix + j]
          ? table[(i + 1) * width + j + 1]! + 1
          : Math.max(table[(i + 1) * width + j]!, table[i * width + j + 1]!)
    }
  }
  let i = 0
  let j = 0
  while (i < rows && j < cols) {
    if (oldKeys[prefix + i] === newKeys[prefix + j]) {
      match[prefix + i] = prefix + j
      i++
      j++
    } else if (table[(i + 1) * width + j]! >= table[i * width + j + 1]!) {
      i++
    } else {
      j++
    }
  }
  return match
}

/** a unit survives only when all its nodes matched one contiguous run of the new document */
function keptAt(unit: Unit, match: Int32Array): number {
  const target = match[unit.first]!
  if (target < 0) return -1
  for (let k = 1; k < unit.count; k++) {
    if (match[unit.first + k] !== target + k) return -1
  }
  return target
}

/**
 * Serialize a run of changed nodes. When the run lines up with the blocks it
 * replaces — same number of nodes, same node type at every position — each
 * block's nodes are rendered in that block's own conventions; otherwise
 * (blocks inserted, merged, split or removed) the run shares the replaced
 * blocks' conventions, first answer wins, over the document's.
 */
function serializeRun(
  manager: NonNullable<Editor['markdown']>,
  nodes: PmNode[],
  replaced: Unit[],
  oldKeys: string[],
  documentStyle: StyleHints,
): string {
  const serialize = (group: PmNode[], style: StyleHints[]): string =>
    withMarkdownStyle(resolveMarkdownStyle(...style, documentStyle), () =>
      manager.serialize({ type: 'doc', content: group.map((n) => n.toJSON()) }),
    )
  const replacedCount = replaced.reduce((sum, unit) => sum + unit.count, 0)
  const aligned =
    replaced.length >= 2 &&
    replacedCount === nodes.length &&
    replaced.every((unit, u) => {
      const before = replaced.slice(0, u).reduce((sum, r) => sum + r.count, 0)
      for (let k = 0; k < unit.count; k++) {
        const oldType = (JSON.parse(oldKeys[unit.first + k]!) as { type: string }).type
        if (nodes[before + k]!.type.name !== oldType) return false
      }
      return true
    })
  if (!aligned) {
    return serialize(
      nodes,
      replaced.map((unit) => unit.style),
    )
  }
  const parts: string[] = []
  let offset = 0
  for (const unit of replaced) {
    parts.push(serialize(nodes.slice(offset, offset + unit.count), [unit.style]))
    offset += unit.count
  }
  return parts.join('\n\n')
}

/**
 * Serialize `doc` against `map`: unchanged blocks come from the source, changed
 * runs from the serializer, blank lines and definitions from wherever they were.
 */
export function spliceMarkdown(editor: Editor, doc: PmNode, map: SourceMap): string {
  const manager = editor.markdown
  if (!manager) return editor.getMarkdown()
  const newNodes: PmNode[] = []
  for (let i = 0; i < doc.childCount; i++) newNodes.push(doc.child(i))
  const match = matchNodes(map.keys, newNodes.map(nodeKey))

  let out = ''
  let needSeparator = false
  let pending: PmNode[] = []
  let pendingGlue: string[] = []
  // the blocks a serialized run replaces, in source order
  let pendingUnits: Unit[] = []
  let next = 0

  const separate = (): void => {
    if (out === '' || out.endsWith('\n\n')) return
    out += out.endsWith('\n') ? '\n' : '\n\n'
  }
  const flush = (): void => {
    if (pending.length > 0) {
      const text = serializeRun(manager, pending, pendingUnits, map.keys, map.style)
      if (text) {
        separate()
        out += text
        needSeparator = true
      }
      pending = []
    }
    pendingUnits = []
    for (const glue of pendingGlue) {
      separate()
      out += glue
      needSeparator = true
    }
    pendingGlue = []
  }

  for (const unit of map.units) {
    if (unit.count === 0) {
      out += unit.glue
      continue
    }
    const target = keptAt(unit, match)
    if (target < 0) {
      // replaced or deleted: its blank lines go, definitions it carried stay
      if (unit.glue.trim() !== '') pendingGlue.push(unit.glue.replace(/^\n+/, ''))
      pendingUnits.push(unit)
      needSeparator = true
      continue
    }
    while (next < target) pending.push(newNodes[next++]!)
    flush()
    if (needSeparator && !unit.space) separate()
    out += unit.core + unit.glue
    needSeparator = false
    next = target + unit.count
  }
  while (next < newNodes.length) pending.push(newNodes[next++]!)
  // the editor's caret paragraph after a trailing list/code/table has no source
  while (pending.length > 0 && isEmptyParagraph(pending[pending.length - 1]!)) pending.pop()
  flush()
  // a kept tail carries the source's own ending; a serialized one gets it back
  if (map.finalNewline && out !== '' && !out.endsWith('\n')) out += '\n'
  return out
}
