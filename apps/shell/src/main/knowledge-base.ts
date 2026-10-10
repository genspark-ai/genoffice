/// Knowledge-base channel: the reader's starred files (the star corpus) as
/// AI can search and read, so an answer can cite the folder and file it came
/// from. Lives in the shell main process because that is the one place that
/// already holds both halves — the starred list (readStarredFiles, docs-main)
/// and the full-text index (file-index). The handlers are app-agnostic IPC;
/// docs wires them into its AI panel first, the other apps can follow.
///
/// Why search filters in JS rather than in SQLite: the FTS index's path column
/// stores tokenized FOLDER names (see store.ts upsert), not real paths, so a
/// `path IN (...)` cannot express "within the starred set". Starred corpora
/// are small (a reader's hand-picked files), and the indexed whole is local,
/// so fetching a bounded page and filtering by real path is the cheap,
/// correct filter.

import { statSync } from 'node:fs'
import { basename, dirname } from 'node:path'
import { parseFileToText } from '@genoffice/file-parse'

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

/** the corpus: every starred file that still exists on disk */
export function listKnowledgeFiles(starredPaths: readonly string[]): KbFile[] {
  const out: KbFile[] = []
  for (const path of starredPaths) {
    try {
      const stat = statSync(path)
      if (!stat.isFile()) continue
      const folderPath = dirname(path)
      out.push({
        path,
        folder: basename(folderPath),
        folderPath,
        name: basename(path),
        ext: basename(path)
          .replace(/^[^.]*./, '')
          .toLowerCase(),
        sizeBytes: stat.size,
      })
    } catch {
      // a starred path can outlive its file (r158 keeps the dimmed row in the
      // home UI); for the corpus it simply drops out
    }
  }
  return out
}

/** fetch + filter + trim one search page; pure so it is unit-testable */
export function searchKnowledgeBase(
  search: (
    q: string,
    limit: number,
  ) => ReadonlyArray<{ path: string; name: string; ext: string; snippetText?: string }>,
  query: string,
  starredPaths: readonly string[],
  limit: number,
): KbSearchHit[] {
  const q = query.trim().slice(0, KB_QUERY_MAX_CHARS)
  if (!q) return []
  const starred = new Set(starredPaths)
  const pool = search(q, KB_SEARCH_POOL)
  const seen = new Set<string>()
  const out: KbSearchHit[] = []
  for (const hit of pool) {
    if (!starred.has(hit.path) || seen.has(hit.path)) continue
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
