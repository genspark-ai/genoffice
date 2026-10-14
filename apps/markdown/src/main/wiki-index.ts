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
import { readFileSync, statSync } from 'node:fs'
import { basename, extname, join, relative } from 'node:path'

import { MAX_DEPTH, SKIP_DIRS, nodeFs, resolveWikiLinks } from './wiki-resolve'
import type { WikiFs } from './wiki-resolve'

/** note files the index walks; media joins the graph only as a link target */
const NOTE_EXTS = new Set(['.md', '.markdown', '.txt'])

/** notes one build may walk, mirroring the suggester's cap */
export const MAX_INDEX_FILES = 2000

/** a filesystem the index walks; injectable so tests drive a real temp tree */
export interface VaultFs extends WikiFs {
  read(path: string): string
  mtimeMs(path: string): number
}

export function nodeVaultFs(): VaultFs {
  const fs = nodeFs()
  return {
    ...fs,
    read: (path) => readFileSync(path, 'utf8'),
    mtimeMs: (path) => {
      try {
        return statSync(path).mtimeMs
      } catch {
        return 0
      }
    },
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

export interface VaultFileEntry {
  path: string
  mtimeMs: number
  /** the targets this file's links carry, deduplicated in order */
  targets: string[]
}

interface IndexState {
  files: Map<string, VaultFileEntry>
  /** resolved edges per file, valid for the mtime the file was parsed at */
  edges: Map<string, { mtimeMs: number; to: { target: string; path: string }[] }>
}

const indexes = new Map<string, IndexState>()

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
function refreshFiles(root: string, fs: VaultFs): IndexState {
  const state = indexFor(root)
  const seen = new Set<string>()
  let queue: { dir: string; depth: number }[] = [{ dir: root, depth: 0 }]
  let budget = MAX_INDEX_FILES * 4
  while (queue.length > 0 && seen.size < MAX_INDEX_FILES && budget > 0) {
    const next: typeof queue = []
    for (const { dir, depth } of queue) {
      let listing: { files: string[]; dirs: string[] }
      try {
        listing = fs.list(dir)
      } catch {
        continue
      }
      budget -= listing.files.length + listing.dirs.length
      for (const file of listing.files) {
        if (!NOTE_EXTS.has(extname(file).toLowerCase())) continue
        const path = join(dir, file)
        if (seen.has(path)) continue
        seen.add(path)
        const mtimeMs = fs.mtimeMs(path)
        const known = state.files.get(path)
        if (known && known.mtimeMs === mtimeMs) continue
        let targets: string[]
        try {
          targets = extractWikiTargets(fs.read(path))
        } catch {
          continue // unreadable this round; the next one retries
        }
        state.files.set(path, { path, mtimeMs, targets })
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
  for (const path of [...state.files.keys()]) {
    if (!seen.has(path)) {
      state.files.delete(path)
      state.edges.delete(path)
    }
  }
  return state
}

/** resolved outgoing edges for every indexed file, computing what is missing */
function refreshEdges(root: string, fs: VaultFs): Map<string, { target: string; path: string }[]> {
  const state = refreshFiles(root, fs)
  const out = new Map<string, { target: string; path: string }[]>()
  for (const [path, entry] of state.files) {
    const done = state.edges.get(path)
    if (done && done.mtimeMs === entry.mtimeMs) {
      out.set(path, done.to)
      continue
    }
    const answers =
      entry.targets.length > 0
        ? resolveWikiLinks(entry.targets, { fromPath: path, root }, fs)
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
export function wikiBacklinks(root: string, path: string, fs: VaultFs): WikiBacklink[] {
  if (!path) return []
  const edges = refreshEdges(root, fs)
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
}

/** the graph's node cap; a bigger vault takes the first walk order */
export const MAX_GRAPH_NODES = 400

/**
 * The vault as nodes and edges: every indexed note is a node (linked or not —
 * orphans are half the point of a graph), every resolved inter-note link an
 * edge. Self-links are dropped, matching the backlinks panel.
 */
export function wikiGraph(root: string, fs: VaultFs): WikiGraph {
  const edges = refreshEdges(root, fs)
  const nodes: WikiGraphNode[] = []
  const state = indexFor(root)
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
  return { nodes, edges: out }
}
