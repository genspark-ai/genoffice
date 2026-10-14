/// Knowledge-base channel: the corpus as a set of sources — the reader's
/// starred files (the built-in source) plus the folder and file sources that
/// knowledge features register (mods-implementation-report.md §9: one concept,
/// not two; a mod contributes a source, never a second engine). AI can search
/// and read all of it, so an answer can cite the folder and file it came from.
/// Lives in the shell main process because that is the one place that already
/// holds both halves — the starred list (readStarredFiles, docs-main), the
/// full-text index (file-index) and the feature registry. The handlers are
/// app-agnostic IPC; the editors wire them into their AI panels.
///
/// A source's corpus is what the mod host's reader already serves
/// (listKnowledgeEntries → pathForEntry): top-level readable files, the entry
/// cap and the traversal guard stay where they were. Those paths ride the
/// index as extra paths, so retrieval and freshness are the index's own;
/// disabled sources and sources whose path is missing are never searched but
/// stay visible (§9.5.2, §9.5.9).
///
/// Why search filters in JS rather than in SQLite: the FTS index's path column
/// stores tokenized FOLDER names (see store.ts upsert), not real paths, so a
/// `path IN (...)` cannot express "within the corpus". The corpus is small (a
/// reader's hand-picked files plus ≤50 readable entries per source), and the
/// indexed whole is local, so fetching a bounded page and filtering by real
/// path is the cheap, correct filter.

import { existsSync, statSync } from 'node:fs'
import { basename, dirname } from 'node:path'
import { parseFileToText } from '@genoffice/file-parse'
import type { FeatureRecord } from '@genoffice/agent-core'

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

/** how many indexed hits to fetch before filtering to the corpus */
const KB_SEARCH_POOL = 200

/** the built-in source id: the reader's starred files */
export const KB_STARRED_SOURCE_ID = 'starred'

/** one source of the corpus, as the Ask panel's Sources section shows it */
export interface KbSource {
  /** 'starred' for the built-in source, else the feature id */
  id: string
  kind: 'starred' | 'folder' | 'file'
  label: string
  /** absolute path of a folder or file source; absent for starred */
  path?: string
  /** a disabled source is kept, never searched, and named in the coverage */
  enabled: boolean
  /** the path does not exist on this machine: flagged, never searched (§9.5.9) */
  missing: boolean
  /**
   * how the source is indexed: 'full-text' rides the search index, 'names'
   * contributes its files to the listing and reads only — their content is
   * never extracted into the index, so a search cannot hit them
   */
  index: 'full-text' | 'names'
  /** files the source contributes to the corpus right now */
  documents: number
  /** their total size, for the Sources section */
  bytes: number
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
  /** the source this file belongs to */
  sourceId: string
  /** the source's display label */
  sourceLabel: string
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

/** the starred source, as the corpus sees it (always enabled: it is the
 *  built-in; toggling it off belongs to unstar-ing, not to the base) */
function starredSource(): KbSource {
  return {
    id: KB_STARRED_SOURCE_ID,
    kind: 'starred',
    label: KB_STARRED_SOURCE_ID,
    enabled: true,
    missing: false,
    index: 'full-text',
    documents: 0,
    bytes: 0,
  }
}

/**
 * The sources the registry declares: one per knowledge feature, with the
 * per-source controls the knowledge base keeps (enabled) and the honest flag
 * for a path that does not exist here (a pack imported on another machine).
 */
export function sourcesFromFeatures(
  features: readonly FeatureRecord[],
  disabledIds: ReadonlySet<string>,
): KbSource[] {
  const sources: KbSource[] = [starredSource()]
  for (const feature of features) {
    const knowledge = feature.knowledge
    if (!knowledge) continue
    sources.push({
      id: feature.id,
      kind: knowledge.source,
      label: knowledge.label ?? feature.name,
      path: knowledge.path,
      enabled: !disabledIds.has(feature.id),
      missing: !existsSync(knowledge.path),
      index: knowledge.index ?? 'full-text',
      documents: 0,
      bytes: 0,
    })
  }
  return sources
}

/** the corpus files one source contributes: the starred paths, or what the mod
 *  host's reader offers for a folder/file source (bounded, traversal-guarded) */
function filesOfSource(
  source: KbSource,
  readEntries: (source: KbSource) => readonly { name: string; bytes: number }[],
  pathOfEntry: (source: KbSource, entry: string) => string,
): KbFile[] {
  if (source.id === KB_STARRED_SOURCE_ID) {
    // the caller passes starred paths through readEntries as paths already
    return []
  }
  const out: KbFile[] = []
  for (const entry of readEntries(source)) {
    let path: string
    try {
      path = pathOfEntry(source, entry.name)
    } catch {
      continue // an entry the guard refuses is not corpus; the source stays usable
    }
    const file = corpusFile(path, source)
    if (file) out.push(file)
  }
  return out
}

function corpusFile(path: string, source: KbSource): KbFile | null {
  try {
    const stat = statSync(path)
    if (!stat.isFile()) return null
    const folderPath = dirname(path)
    return {
      path,
      folder: basename(folderPath),
      folderPath,
      name: basename(path),
      ext: basename(path)
        .replace(/^[^.]*./, '')
        .toLowerCase(),
      sizeBytes: stat.size,
      sourceId: source.id,
      sourceLabel: source.label,
    }
  } catch {
    return null
  }
}

/** the starred files as corpus entries (the original shape, kept for the star list) */
export function listKnowledgeFiles(starredPaths: readonly string[]): KbFile[] {
  const source = starredSource()
  const out: KbFile[] = []
  for (const path of starredPaths) {
    const file = corpusFile(path, source)
    if (file) out.push(file)
  }
  return out
}

/**
 * The merged corpus: every enabled source's files, and per-source counts for
 * the Sources section and the coverage line. `readEntries`/`pathOfEntry` are
 * the mod host's reader (listKnowledgeEntries / pathForEntry); for the starred
 * source the caller supplies its paths in place of entries.
 */
export function assembleCorpus(
  starredPaths: readonly string[],
  sources: readonly KbSource[],
  readEntries: (source: KbSource) => readonly { name: string; bytes: number }[],
  pathOfEntry: (source: KbSource, entry: string) => string,
): { sources: KbSource[]; files: KbFile[] } {
  const files: KbFile[] = []
  const counted = sources.map((source) => {
    if (source.missing || !source.enabled) return source
    const sourceFiles =
      source.id === KB_STARRED_SOURCE_ID
        ? listKnowledgeFiles(starredPaths)
        : filesOfSource(source, readEntries, pathOfEntry)
    files.push(...sourceFiles)
    return {
      ...source,
      documents: sourceFiles.length,
      bytes: sourceFiles.reduce((sum, file) => sum + file.sizeBytes, 0),
    }
  })
  return { sources: counted, files }
}

/**
 * The corpus paths whose content rides the full-text index: a names-only
 * source keeps its files in the listing (readable, citable) but contributes
 * nothing to the index, so a search cannot hit it and the coverage line says
 * so. Pure so it is unit-testable.
 */
export function indexableCorpusPaths(corpus: {
  sources: readonly KbSource[]
  files: readonly KbFile[]
}): string[] {
  const fullText = new Set(
    corpus.sources
      .filter((source) => source.enabled && source.index === 'full-text')
      .map((source) => source.id),
  )
  return corpus.files.filter((file) => fullText.has(file.sourceId)).map((file) => file.path)
}

/** fetch + filter + trim one search page; pure so it is unit-testable */
export function searchKnowledgeBase(
  search: (
    q: string,
    limit: number,
  ) => ReadonlyArray<{ path: string; name: string; ext: string; snippetText?: string }>,
  query: string,
  corpusPaths: readonly string[],
  limit: number,
): KbSearchHit[] {
  const q = query.trim().slice(0, KB_QUERY_MAX_CHARS)
  if (!q) return []
  const corpus = new Set(corpusPaths)
  const pool = search(q, KB_SEARCH_POOL)
  const seen = new Set<string>()
  const out: KbSearchHit[] = []
  for (const hit of pool) {
    if (!corpus.has(hit.path) || seen.has(hit.path)) continue
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
