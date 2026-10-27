/**
 * The vault's cross-file link index — what the backlinks panel and the graph
 * view both read. Nothing here writes a byte: it parses the `[[...]]` /
 * `![[...]]` literals every Obsidian-compatible vault already carries, so a
 * graph drawn here and a backlink listed here correspond to what Obsidian
 * itself shows for the same files.
 *
 * Per-file answers are memoized by mtime: a steady-state request stats the
 * tree once, re-reads only what changed, and the expensive half (resolving
 * every target to a file) runs just for those.
 */
import { readFile, stat } from 'node:fs/promises'
import { basename, extname, join, relative } from 'node:path'

import { MAX_DEPTH, SKIP_DIRS, nodeFs, resolveWikiLinks } from './wiki-resolve'
import type { WikiFs } from './wiki-resolve'

/** note files the index walks; media joins the graph only as a link target */
const NOTE_EXTS = new Set(['.md', '.markdown', '.txt'])

/** notes one build may walk, mirroring the suggester's cap */
export const MAX_INDEX_FILES = 2000

/**
 * A filesystem the index walks; injectable so tests drive a real temp tree.
 *
 * WikiFs is already awaited (isFile / list return promises, the listing cached
 * per directory), so the two reads a build adds are awaited too: the whole
 * index runs off the blocking filesystem calls.
 */
export interface VaultFs extends WikiFs {
  read(path: string): Promise<string>
  mtimeMs(path: string): Promise<number>
}

export function nodeVaultFs(): VaultFs {
  const fs = nodeFs()
  return {
    ...fs,
    read: (path) => readFile(path, 'utf8'),
    mtimeMs: (path) =>
      stat(path)
        .then((s) => s.mtimeMs)
        .catch(() => 0),
  }
}

/**
 * The wiki targets one document carries, in order, deduplicated. Fenced code
 * blocks and inline code spans are stripped first — they are where a vault
 * *talks about* the syntax, which is not a link — and an escaped `\\[\\[` is
 * not a link, matching what the editor's own tokenizer recognizes. What comes
 * out is the target as the resolver keys it: `Note`, `Note#Heading`, with any
 * alias removed.
 */
export function extractWikiTargets(text: string): string[] {
  const out: string[] = []
  const seen = new Set<string>()
  let inFence = false
  let fence = ''
  for (const line of text.split('\n')) {
    const opens = /^\s{0,3}(```|~~~)/.exec(line)
    if (opens) {
      const marker = opens[1]!
      if (!inFence) {
        inFence = true
        fence = marker
      } else if (marker === fence) {
        inFence = false
        fence = ''
      }
      continue
    }
    if (inFence) continue
    const clean = line.replace(/`[^`\n]*`/g, ' ')
    for (const match of clean.matchAll(WIKI_TARGET_RE)) {
      const target = match[1]!
      if (!seen.has(target)) {
        seen.add(target)
        out.push(target)
      }
    }
  }
  return out
}

const WIKI_TARGET_RE = /!?\[\[([^[|\n]+?)(?:\|[^[|\n]*)?\]\]/g

/**
 * `#tag` as Obsidian reads it: `#` followed by a letter (any script), then
 * letters/digits/`-`/`_`/`/`. Not a heading (`#` needs a space after to be
 * one), not inside a fence or inline code — the same stripping walk
 * extractWikiTargets uses.
 */
const TAG_RE = /(^|[^\w#])#(\p{L}[\p{L}\p{N}/_-]*)/gu

export function extractTags(text: string): string[] {
  const out: string[] = []
  const seen = new Set<string>()
  let inFence = false
  let fence = ''
  for (const line of text.split('\n')) {
    const opens = /^\s{0,3}(```|~~~)/.exec(line)
    if (opens) {
      const marker = opens[1]!
      if (!inFence) {
        inFence = true
        fence = marker
      } else if (marker === fence) {
        inFence = false
        fence = ''
      }
      continue
    }
    if (inFence) continue
    const clean = line.replace(/`[^`\n]*`/g, ' ')
    for (const match of clean.matchAll(TAG_RE)) {
      const tag = match[2]!
      if (!seen.has(tag)) {
        seen.add(tag)
        out.push(tag)
      }
    }
  }
  return out
}

export interface VaultFileEntry {
  path: string
  mtimeMs: number
  /** the targets this file's links carry, deduplicated in order */
  targets: string[]
  /** the #tags this file carries, deduplicated in order */
  tags: string[]
}

interface IndexState {
  files: Map<string, VaultFileEntry>
  /** resolved edges per file, valid for the mtime the file was parsed at */
  edges: Map<string, { mtimeMs: number; to: { target: string; path: string }[] }>
}

const indexes = new Map<string, IndexState>()

/** in-flight builds per root, chained — see refreshEdges */
const building = new Map<string, Promise<Map<string, { target: string; path: string }[]>>>()

function indexFor(root: string): IndexState {
  let state = indexes.get(root)
  if (!state) {
    state = { files: new Map(), edges: new Map() }
    indexes.set(root, state)
  }
  return state
}

/**
 * Walk the vault's notes, re-reading only what changed since the last build.
 * Same root walk as the resolver and the suggester, so the three surfaces
 * always see the same vault.
 */
function refreshFiles(root: string, fs: VaultFs): Promise<IndexState> {
  const state = indexFor(root)
  const seen = new Set<string>()
  // taken BEFORE the walk: a note the walk is about to add is not in the set
  // yet, which is the whole point — after the walk both sets look identical
  const before = new Set(state.files.keys())
  let setMoved = false
  let queue: { dir: string; depth: number }[] = [{ dir: root, depth: 0 }]
  let budget = MAX_INDEX_FILES * 4
  return (async () => {
    while (queue.length > 0 && seen.size < MAX_INDEX_FILES && budget > 0) {
      const next: typeof queue = []
      for (const { dir, depth } of queue) {
        let listing: { files: string[]; dirs: string[] }
        try {
          listing = await fs.list(dir)
        } catch {
          continue
        }
        budget -= listing.files.length + listing.dirs.length
        for (const file of listing.files) {
          if (!NOTE_EXTS.has(extname(file).toLowerCase())) continue
          const path = join(dir, file)
          if (seen.has(path)) continue
          seen.add(path)
          const mtimeMs = await fs.mtimeMs(path)
          const known = state.files.get(path)
          if (known && known.mtimeMs === mtimeMs) continue
          let targets: string[]
          let tags: string[]
          try {
            const text = await fs.read(path)
            targets = extractWikiTargets(text)
            tags = extractTags(text)
          } catch {
            continue // unreadable this round; the next one retries
          }
          state.files.set(path, { path, mtimeMs, targets, tags })
          state.edges.delete(path)
        }
        if (seen.size >= MAX_INDEX_FILES || depth + 1 > MAX_DEPTH) continue
        for (const sub of listing.dirs) {
          if (SKIP_DIRS.has(sub)) continue
          next.push({ dir: join(dir, sub), depth: depth + 1 })
        }
      }
      queue = next
    }
    // A memoized edge answers "does this target exist?", so it is only valid
    // while the file SET is what it was resolved against. Creating the target a
    // link pointed at, or renaming it away, changes the answer without changing
    // the source file at all — keyed on the source mtime alone, A keeps its dead
    // path until A is re-saved. Drop every memo when the set moved.
    for (const path of [...state.files.keys()]) {
      if (!seen.has(path)) {
        state.files.delete(path)
        setMoved = true
      }
    }
    // compare the paths, not the sizes: a walk that stops on the file cap can add
    // a note while another falls out, leaving the count unchanged
    for (const path of seen) {
      if (!before.has(path)) {
        setMoved = true
        break
      }
    }
    if (setMoved) state.edges.clear()
    return state
  })()
}

/**
 * Resolved outgoing edges for every indexed file, computing what is missing.
 *
 * Chained per root. The panel and the graph usually open together, and each
 * builds walks the tree and writes one shared IndexState — left unchained they
 * interleave across their awaits and read a half-written file set. The second
 * caller waits and then finds the memo warm.
 */
function refreshEdges(
  root: string,
  fs: VaultFs,
): Promise<Map<string, { target: string; path: string }[]>> {
  const previous = building.get(root) ?? Promise.resolve(new Map())
  const next = previous
    .catch(() => new Map())
    .then(async () => {
      const state = await refreshFiles(root, fs)
      const out = new Map<string, { target: string; path: string }[]>()
      for (const [path, entry] of state.files) {
        const done = state.edges.get(path)
        if (done && done.mtimeMs === entry.mtimeMs) {
          out.set(path, done.to)
          continue
        }
        const answers =
          entry.targets.length > 0
            ? await resolveWikiLinks(entry.targets, { fromPath: path, root }, fs)
            : new Map()
        const to: { target: string; path: string }[] = []
        for (const [target, answer] of answers) {
          if (answer.found && answer.path && answer.path !== path) {
            to.push({ target, path: answer.path })
          }
        }
        state.edges.set(path, { mtimeMs: entry.mtimeMs, to })
        out.set(path, to)
      }
      return out
    })
  building.set(root, next)
  return next.finally(() => {
    if (building.get(root) === next) building.delete(root)
  })
}

export interface WikiBacklink {
  /** absolute path of the note whose link lands here */
  source: string
  /** the target as written there (`Note#Heading`) */
  target: string
}

/**
 * Every note whose link resolves to `path`. Self-links (`[[#Heading]]` in the
 * note itself) are answers too, but they point a note at itself — excluded,
 * same as the graph, because a panel of a note linking to itself is noise.
 */
export async function wikiBacklinks(
  root: string,
  path: string,
  fs: VaultFs,
): Promise<WikiBacklink[]> {
  if (!path) return []
  const edges = await refreshEdges(root, fs)
  const out: WikiBacklink[] = []
  const seen = new Set<string>()
  for (const [source, list] of edges) {
    for (const edge of list) {
      if (edge.path !== path) continue
      const key = `${source}\u0000${edge.target}`
      if (seen.has(key)) continue
      seen.add(key)
      out.push({ source, target: edge.target })
    }
  }
  return out
}

export interface WikiTagEntry {
  tag: string
  /** absolute paths of notes carrying the tag, sorted */
  files: string[]
}

/**
 * Every #tag in the vault with the notes that carry it, most-used first.
 * Falls out of the same index the backlinks and graph read — no second walk.
 */
export async function wikiTags(root: string, fs: VaultFs): Promise<WikiTagEntry[]> {
  const state = await refreshFiles(root, fs)
  const byTag = new Map<string, string[]>()
  for (const [path, entry] of state.files) {
    for (const tag of entry.tags) {
      const list = byTag.get(tag)
      if (list) list.push(path)
      else byTag.set(tag, [path])
    }
  }
  return [...byTag.entries()]
    .map(([tag, files]) => ({ tag, files: files.sort() }))
    .sort((a, b) => b.files.length - a.files.length || (a.tag < b.tag ? -1 : 1))
}

export interface WikiGraphNode {
  path: string
  /** stem, the label a node carries */
  name: string
  /** path relative to the root, the label's smaller half */
  relPath: string
}

export interface WikiGraphEdge {
  from: string
  to: string
}

export interface WikiGraph {
  nodes: WikiGraphNode[]
  edges: WikiGraphEdge[]
  /** the node cap cut the vault short: this is a slice, not the whole graph */
  truncated: boolean
  /** notes the walk found, which is at least nodes.length and larger when cut */
  totalNotes: number
}

/** the graph's node cap; a bigger vault takes the first walk order */
export const MAX_GRAPH_NODES = 400

/**
 * The vault as nodes and edges: every indexed note is a node (linked or not —
 * orphans are half the point of a graph), every resolved inter-note link an
 * edge. Self-links are dropped, matching the backlinks panel.
 */
export async function wikiGraph(root: string, fs: VaultFs): Promise<WikiGraph> {
  const edges = await refreshEdges(root, fs)
  const nodes: WikiGraphNode[] = []
  const state = indexFor(root)
  const totalNotes = state.files.size
  for (const path of state.files.keys()) {
    nodes.push({
      path,
      name: basename(path, extname(path)),
      relPath: relative(root, path),
    })
    if (nodes.length >= MAX_GRAPH_NODES) break
  }
  const nodeSet = new Set(nodes.map((n) => n.path))
  const out: WikiGraphEdge[] = []
  const seen = new Set<string>()
  for (const [from, list] of edges) {
    if (!nodeSet.has(from)) continue
    for (const edge of list) {
      if (!nodeSet.has(edge.path)) continue
      const key = `${from}\u0000${edge.path}`
      if (seen.has(key)) continue
      seen.add(key)
      out.push({ from, to: edge.path })
    }
  }
  // a cut graph is a slice: say so instead of letting a 400-of-3000 draw
  // read as the whole vault
  return { nodes, edges: out, truncated: totalNotes > nodes.length, totalNotes }
}

/**
 * Start a build without waiting for it.
 *
 * Called once the vault root is known, so the panel and the graph usually meet
 * a warm index instead of waiting out a cold build. The mtime memo means a
 * later request pays only for what changed.
 */
export function warmWikiIndex(root: string, fs: VaultFs = nodeVaultFs()): void {
  if (!root) return
  void refreshEdges(root, fs).catch(() => undefined)
}
