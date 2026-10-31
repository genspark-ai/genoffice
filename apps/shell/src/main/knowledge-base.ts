/// Knowledge-base channel: the user's local files as AI can search and read,
/// so an answer can cite the folder and file it came from. Lives in the shell
/// main process because that is the one place that already holds both halves —
/// the starred list (readStarredFiles, docs-main) and the full-text index
/// (file-index). The handlers are app-agnostic IPC; docs wires them into its
/// AI panel first, the other apps can follow.
///
/// The corpus is a scope: 'indexed' (the default) is every file the local
/// index covers — the save folder, the added folder roots, recents and the
/// starred set — while 'starred' is the original hand-picked set. kb:list
/// stays a bounded prompt window either way (starred first, then the most
/// recently touched indexed files); search and read are what reach the whole
/// corpus.
///
/// On top of that sits the round scope: an explicit file selection (the
/// desktop's multi-select → Ask AI) handed over as one conversation's corpus.
/// While a request carries `onlyPaths`, search intersects the index with that
/// list and reports the paths the index does not cover (`notIndexed`), and
/// read narrows the same way — a selected path is the only thing readable
/// that call, corpus or not. A selected file the index missed still reads
/// when its location is one a desktop selection can genuinely come from
/// (inside a folder root, on the stored alias list, or a recent/starred
/// file — the caller's verdict) and it exists as a readable office type: an
/// office-named file anywhere else on disk is refused, so the scope cannot
/// turn the read channel into an arbitrary-file reader. No ad-hoc indexing
/// happens for the selection (v1 boundary: an unindexed selected file that
/// fails the location gate is simply not answerable, and the panel's chip
/// says so).
///
/// Why the starred search still filters in JS rather than in SQLite: the FTS
/// index's path column stores tokenized FOLDER names (see store.ts upsert),
/// not real paths, so a `path IN (...)` cannot express "within the starred
/// set". Starred corpora are small, so fetching a bounded pool and filtering
/// by real path is the cheap, correct filter. The indexed scope needs no
/// post-filter at all, so it takes the store's own ranked page directly.

import { statSync } from 'node:fs'
import { basename, dirname, isAbsolute } from 'node:path'
import { parseFileToText } from '@genoffice/file-parse'
import { isSupportedTreeFile } from './folder-tree'
import type { IndexedFile } from './file-index/store'

/** characters of parsed text one read_knowledge_file page returns (aligned
 *  with the chat-attachment reader, files-skill.ts, so models meet one paging
 *  convention everywhere) */
export const KB_READ_CHUNK_CHARS = 24_000

/** byte cap per corpus file, aligned with the chat-attachment reader's
 *  ATTACHMENT_MAX_BYTES: a larger file is refused before parse rather than
 *  parsed and held in the text cache */
export const KB_READ_MAX_BYTES = 50 * 1024 * 1024

/** query cap, same as the home search box */
const KB_QUERY_MAX_CHARS = 200

/** how many indexed hits to fetch before filtering to the starred set */
const KB_SEARCH_POOL = 200

/**
 * How many files kb:list returns for the indexed scope: the listing is a
 * prompt window (the model searches first; the system prompt says so), and a
 * hundred rows is about what the knowledge-base skill's context budget shows
 * before it degrades to names only.
 */
export const KB_LIST_CAP = 100

/** the knowledge-base corpus scope (app-settings.json `kbScope`); default: the whole index */
export type KbScope = 'indexed' | 'starred'

export function normalizeKbScope(raw: unknown): KbScope {
  return raw === 'starred' ? 'starred' : 'indexed'
}

// ── the round scope: an explicit file selection as one conversation's corpus ──

/** how many paths one scoped kb request may carry; a desktop multi-select is a handful, the cap keeps a hostile payload bounded */
export const KB_ONLY_PATHS_MAX = 50

/** per-path length cap, same as the open-path channel */
const KB_PATH_MAX_CHARS = 4096

/**
 * Validate an `onlyPaths` payload (kb:search / kb:read): null when the field
 * is absent or empty — the default kbScope corpus applies — and a deduped
 * path list otherwise. Throws a plain message on a non-array, a non-string
 * entry, a relative path, an over-long path, or too many paths: a renderer
 * cannot widen the round scope into an unbounded read surface.
 */
export function parseKbOnlyPaths(raw: unknown): readonly string[] | null {
  if (raw === undefined || raw === null) return null
  if (!Array.isArray(raw)) throw new Error('onlyPaths must be an array of absolute paths')
  if (raw.length > KB_ONLY_PATHS_MAX) {
    throw new Error(`onlyPaths accepts at most ${KB_ONLY_PATHS_MAX} paths`)
  }
  const seen = new Set<string>()
  const out: string[] = []
  for (const entry of raw) {
    if (typeof entry !== 'string') throw new Error('onlyPaths entries must be strings')
    if (!isAbsolute(entry)) throw new Error(`onlyPaths entries must be absolute: ${entry}`)
    if (entry.length > KB_PATH_MAX_CHARS) {
      throw new Error(`onlyPaths entry is too long (${entry.length} > ${KB_PATH_MAX_CHARS})`)
    }
    if (seen.has(entry)) continue
    seen.add(entry)
    out.push(entry)
  }
  return out.length > 0 ? out : null
}

/**
 * A round-scope path the index missed stays readable when its location is
 * one the caller whitelists — inside a folder root, on the desktop's stored
 * alias list, or a tracked recent/starred file (the locations a desktop
 * selection can genuinely hold) — and it exists, is a regular file, and
 * carries one of the tree's supported office extensions. Location plus type
 * keeps the scope from turning the read channel into an arbitrary-file
 * reader: an office-named file anywhere on disk is not enough on its own.
 */
export function isReadableScopePath(
  path: string,
  opts: {
    /** the location verdict (root / alias / recent / starred), computed by the caller */
    locationAllowed: boolean
    statFile: (path: string) => { isFile: () => boolean } | null
  },
): boolean {
  if (!opts.locationAllowed) return false
  if (!isSupportedTreeFile(basename(path))) return false
  return opts.statFile(path)?.isFile() ?? false
}

/**
 * The kb:read gate, pure so it is unit-testable. Without a round scope the
 * kbScope corpus verdict is the whole gate (starred, or in the index) — the
 * read surface the channel always had. With a round scope attached the scope
 * narrows instead of widening: only a selected path is readable that call
 * (search already intersects with the selection; read must too, or a corpus
 * file outside the chip would still leak through), and a selected path the
 * corpus does not cover reads when its location is whitelisted and it is a
 * readable office type.
 */
export function kbReadAllows(
  path: string,
  opts: {
    /** the kbScope-side verdict (starred / index membership), computed by the caller */
    inCorpus: boolean
    /** this request's validated onlyPaths; null = no round scope */
    scopePaths: readonly string[] | null
    /** the round-scope location verdict (root / alias / recent / starred), computed by the caller */
    locationAllowed: boolean
    statFile: (path: string) => { isFile: () => boolean } | null
  },
): boolean {
  if (opts.scopePaths === null) return opts.inCorpus
  if (!opts.scopePaths.includes(path)) return false
  if (opts.inCorpus) return true
  return isReadableScopePath(path, {
    locationAllowed: opts.locationAllowed,
    statFile: opts.statFile,
  })
}

/**
 * Which requested scope paths the index does not cover — the answer's
 * `notIndexed` list and the chip's "not indexed" hint. A path without the
 * store is unanswerable, so the whole selection reports as not indexed.
 */
export async function notIndexedPaths(
  store: { has(path: string): Promise<boolean> } | null,
  paths: readonly string[],
): Promise<string[]> {
  if (!store) return [...paths]
  const indexed = await Promise.all(paths.map((p) => store.has(p).catch(() => false)))
  return paths.filter((_, i) => !indexed[i])
}

/** one corpus entry as the AI (and the panel) sees it */
export interface KbFile {
  /** absolute path — also the citation target (filenav:///abs/path) */
  path: string
  /** owning folder name ("contracts" in /Users/me/contracts/lease.docx) */
  folder: string
  /** absolute folder path */
  folderPath: string
  name: string
  ext: string
  sizeBytes: number
}

/** search result as the model sees it */
export interface KbSearchHit {
  path: string
  name: string
  ext: string
  folder: string
  snippet: string
}

/** text extraction cache, keyed by path, invalidated by mtime+size */
const kbTextCache = new Map<string, { stamp: string; text: string }>()
const KB_TEXT_CACHE_MAX = 8

/** shape one KbFile from a path and its known size (stats taken by the caller) */
function kbFileOf(path: string, sizeBytes: number): KbFile {
  const folderPath = dirname(path)
  return {
    path,
    folder: basename(folderPath),
    folderPath,
    name: basename(path),
    ext: basename(path)
      .replace(/^[^.]*./, '')
      .toLowerCase(),
    sizeBytes,
  }
}

/** the starred corpus: every starred file that still exists on disk */
export function listKnowledgeFiles(starredPaths: readonly string[]): KbFile[] {
  const out: KbFile[] = []
  for (const path of starredPaths) {
    try {
      const stat = statSync(path)
      if (!stat.isFile()) continue
      out.push(kbFileOf(path, stat.size))
    } catch {
      // a starred path can outlive its file (r158 keeps the dimmed row in the
      // home UI); for the corpus it simply drops out
    }
  }
  return out
}

/**
 * The prompt's listing window for the indexed scope: the starred files first
 * (the user's deliberate picks), then the most recently touched indexed files
 * that are not already listed, capped at KB_LIST_CAP. Discovery past the
 * window goes through search — the skill's prompt says exactly that.
 */
export function listIndexedKnowledgeFiles(
  starredPaths: readonly string[],
  recentIndexed: readonly IndexedFile[],
  cap: number = KB_LIST_CAP,
): KbFile[] {
  const out = listKnowledgeFiles(starredPaths)
  const listed = new Set(out.map((f) => f.path))
  for (const entry of recentIndexed) {
    if (out.length >= cap) break
    if (listed.has(entry.path)) continue
    out.push(kbFileOf(entry.path, entry.sizeBytes))
    listed.add(entry.path)
  }
  return out
}

/**
 * Fetch + trim one search page; pure so it is unit-testable. `scopePaths` null
 * means the indexed scope: no post-filter, so the store's own ranked page is
 * the answer (the SQL pushdown the starred scope cannot express). A real path
 * set keeps the starred behavior: a bounded pool, JS-filtered by membership.
 * Async because the search runs in the store worker.
 */
export async function searchKnowledgeBase(
  search: (
    q: string,
    limit: number,
  ) => Promise<ReadonlyArray<{ path: string; name: string; ext: string; snippetText?: string }>>,
  query: string,
  scopePaths: readonly string[] | null,
  limit: number,
): Promise<KbSearchHit[]> {
  const q = query.trim().slice(0, KB_QUERY_MAX_CHARS)
  if (!q) return []
  const scoped = scopePaths === null ? null : new Set(scopePaths)
  const pool = await search(q, scoped === null ? limit : KB_SEARCH_POOL)
  const seen = new Set<string>()
  const out: KbSearchHit[] = []
  for (const hit of pool) {
    if (scoped !== null && !scoped.has(hit.path)) continue
    if (seen.has(hit.path)) continue
    seen.add(hit.path)
    out.push({
      path: hit.path,
      name: hit.name,
      ext: hit.ext,
      folder: basename(dirname(hit.path)),
      snippet: hit.snippetText ?? '',
    })
    if (out.length >= limit) break
  }
  return out
}

/** parsed text of one corpus file; throws with a plain message the panel can
 *  show — shell main has no i18n channel of its own (unlike the docs
 *  attachment reader, whose errors carry locale strings and a larger budget) */
export async function readKnowledgeFileText(filePath: string): Promise<string> {
  const stat = statSync(filePath)
  const stamp = `${stat.mtimeMs}:${stat.size}`
  const cached = kbTextCache.get(filePath)
  if (cached && cached.stamp === stamp) return cached.text
  if (stat.size > KB_READ_MAX_BYTES) {
    throw new Error(
      `file too large to read (${Math.round(stat.size / 1024 / 1024)}MB, cap ${KB_READ_MAX_BYTES / 1024 / 1024}MB)`,
    )
  }
  const parsed = await parseFileToText(filePath)
  if (!parsed.ok || parsed.kind !== 'text' || parsed.text == null) {
    throw new Error(parsed.error ?? 'unreadable file')
  }
  kbTextCache.set(filePath, { stamp, text: parsed.text })
  if (kbTextCache.size > KB_TEXT_CACHE_MAX) {
    const oldest = kbTextCache.keys().next().value
    if (oldest) kbTextCache.delete(oldest)
  }
  return parsed.text
}

/** one page of a corpus file's text */
export interface KbReadPage {
  text: string
  /** total parsed characters, so the reader knows whether to page on */
  totalChars: number
  offset: number
}

export async function readKnowledgeFilePage(filePath: string, offset: number): Promise<KbReadPage> {
  const text = await readKnowledgeFileText(filePath)
  const start = Number.isFinite(offset) ? Math.max(0, Math.floor(offset)) : 0
  if (start >= text.length) return { text: '', totalChars: text.length, offset: start }
  return {
    text: text.slice(start, start + KB_READ_CHUNK_CHARS),
    totalChars: text.length,
    offset: start,
  }
}

// ---- create_document: the ask panel's one write path ----

export type KbCreateFormat = 'md' | 'docx'

export interface KbCreateRequest {
  title: string
  format: KbCreateFormat
  content: string
}

/** keep in step with CREATE_DOCUMENT_MAX_CHARS in agent-core */
export const KB_CREATE_MAX_CHARS = 2_000_000

/** validate the renderer's payload; throws a plain message for the panel */
export function parseKbCreateRequest(raw: unknown): KbCreateRequest {
  const req = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const format = req.format
  if (format !== 'md' && format !== 'docx') throw new Error('format must be "md" or "docx"')
  const content = typeof req.content === 'string' ? req.content : ''
  if (!content.trim()) throw new Error('content must not be empty')
  if (content.length > KB_CREATE_MAX_CHARS) throw new Error('content is too long')
  return { title: typeof req.title === 'string' ? req.title : '', format, content }
}

export interface KbCreateDeps {
  /** folder new files land in */
  dir: () => string
  /** `${sanitized title}${ext}` → first free absolute path in dir */
  uniquePath: (dir: string, fileName: string) => string
  sanitizeBase: (title: string) => string
  writeText: (path: string, text: string) => Promise<void>
  /** markdown → .docx at `out` (the bundled CLI) */
  markdownToDocx: (markdown: string, out: string) => Promise<void>
  recordRecent: (path: string) => void
}

export async function createKnowledgeDocument(
  deps: KbCreateDeps,
  req: KbCreateRequest,
): Promise<{ path: string }> {
  const path = deps.uniquePath(deps.dir(), `${deps.sanitizeBase(req.title)}.${req.format}`)
  if (req.format === 'md') await deps.writeText(path, req.content)
  else await deps.markdownToDocx(req.content, path)
  deps.recordRecent(path)
  return { path }
}
