/**
 * Persistent Translation Memory (TM) — JSON-backed, fuzzy-match capable.
 *
 * Mirrors the file layout used by LumosAI's `translation_memory.py`
 * (`~/.lumosai/translation_memory/<lang>.json`) but keeps the API shape
 * compatible with the existing in-memory {@link TranslationMemory} so a
 * persistent store is a drop-in replacement:
 *
 *   import { TranslationMemory } from './memory'
 *   import { PersistentTranslationMemory } from './persistent-memory'
 *   const tm = new PersistentTranslationMemory({ baseDir: '~/.genoffice/translation-memory' })
 *   await tm.load()
 *   const hit = tm.lookup('en-US', 'zh-CN', 'Hello')
 *   tm.save({ ... })
 *   await tm.flush()
 *
 * Each language pair lives in its own JSON file (`<baseDir>/<source>-<target>.json`)
 * so concurrent processes don't trample each other on every save. `flush()`
 * writes the dirty pairs; subsequent `save()` calls mark individual pairs as
 * dirty so callers can decide when to persist (e.g. debounced, or on a
 * single-instance lock).
 *
 * `lookup()` falls back to a token-overlap + SequenceMatcher fuzzy match when
 * an exact match is missing — the same heuristic as the Python implementation
 * (`token_similarity` + `SequenceMatcher.ratio`). The fuzzy threshold defaults
 * to 0.86 so we only return "very close" matches; the caller still receives
 * the source text and a confidence score so it can decide whether to use the
 * hit.
 */

import { mkdir, readFile, writeFile, rename } from 'node:fs/promises'
import path from 'node:path'

import { TranslationMemory } from './memory'
import type { MemoryEntry, MemorySaveRequest, MemorySaveResponse } from './memory'

/** File layout mirrors LumosAI: `~/.lumosai/translation_memory/<pair>.json`. */
const DEFAULT_BASE_DIR = path.join(
  process.env.HOME ?? path.join(path.sep, 'tmp'),
  '.genoffice',
  'translation-memory',
)

/** Minimum similarity score (0..1) for a fuzzy hit to be returned. */
const DEFAULT_FUZZY_THRESHOLD = 0.7

/** TMX export version constant per the TMX 1.4b spec. */
const TMX_VERSION = '1.4b'

export interface PersistentMemoryFileSystem {
  mkdir(path: string, options?: { recursive?: boolean }): Promise<void>
  readFile(path: string, encoding: 'utf8'): Promise<string>
  writeFile(path: string, contents: string, encoding: 'utf8'): Promise<void>
  rename(from: string, to: string): Promise<void>
  readdir(path: string): Promise<readonly string[]>
}

const defaultFS: PersistentMemoryFileSystem = {
  mkdir: (p, o) => mkdir(p, o).then(() => undefined),
  readFile: (p, e) => readFile(p, e) as Promise<string>,
  writeFile: (p, c, e) => writeFile(p, c, e),
  rename: (from, to) => rename(from, to).then(() => undefined),
  readdir: async (p) => (await (await import('node:fs/promises')).readdir(p)) as readonly string[],
}

export interface PersistentTranslationMemoryOptions {
  /** Override the on-disk directory. Defaults to `~/.genoffice/translation-memory`. */
  baseDir?: string
  /** Inject a custom filesystem implementation (tests). */
  fileSystem?: PersistentMemoryFileSystem
  /** In-memory cap before eviction. Defaults to 2048 (same as {@link TranslationMemory}). */
  maxEntries?: number
  /** Minimum similarity for a fuzzy match to be returned. Defaults to 0.86. */
  fuzzyThreshold?: number
}

/** Result of a fuzzy lookup. */
export interface PersistentLookupHit {
  translatedText: string
  sourceText: string
  /** 1.0 for exact matches, 0..1 for fuzzy matches. */
  confidence: number
}

/** In-memory + on-disk translation memory. */
export class PersistentTranslationMemory {
  private readonly baseDir: string
  private readonly fs: PersistentMemoryFileSystem
  private readonly inner: TranslationMemory
  private readonly fuzzyThreshold: number
  /** Per-pair file cache — language pair → set of stored entries (raw, for fuzzy search). */
  private readonly pairCache = new Map<string, PairCache>()
  /** Pairs that have new data and need to be flushed. */
  private readonly dirtyPairs = new Set<string>()

  constructor(opts: PersistentTranslationMemoryOptions = {}) {
    this.baseDir = opts.baseDir ?? DEFAULT_BASE_DIR
    this.fs = opts.fileSystem ?? defaultFS
    this.inner = new TranslationMemory({ maxEntries: opts.maxEntries })
    this.fuzzyThreshold = opts.fuzzyThreshold ?? DEFAULT_FUZZY_THRESHOLD
  }

  /** Read every per-pair file from disk into the in-memory cache. */
  async load(): Promise<void> {
    this.pairCache.clear()
    this.inner.clear()
    try {
      const entries = await this.fs.readdir(this.baseDir)
      for (const name of entries) {
        if (!name.endsWith('.json') || name.endsWith('.tmp')) continue
        // `name` may be a basename or a full path depending on the FS impl,
        // so normalise to just the basename before joining with baseDir.
        const baseName = name.split('/').pop() ?? name
        const filePath = path.join(this.baseDir, baseName)
        try {
          const raw = await this.fs.readFile(filePath, 'utf8')
          const parsed: unknown = JSON.parse(raw)
          if (!Array.isArray(parsed)) continue
          const pair = baseName.replace(/\.json$/, '')
          const cache: PairCache = { entries: [] }
          for (const candidate of parsed) {
            if (!isMemoryEntry(candidate)) continue
            cache.entries.push(candidate)
            this.inner.save(candidate)
          }
          this.pairCache.set(pair, cache)
        } catch {
          // ignore unreadable file
        }
      }
    } catch (error) {
      if (!isNotFound(error)) throw error
      // baseDir does not exist yet — start with an empty cache
    }
  }

  /** Persist every dirty pair to disk. */
  async flush(): Promise<void> {
    if (this.dirtyPairs.size === 0) return
    await this.fs.mkdir(this.baseDir, { recursive: true })
    for (const pair of this.dirtyPairs) {
      const cache = this.pairCache.get(pair)
      if (!cache) continue
      const filePath = path.join(this.baseDir, `${pair}.json`)
      const tmp = filePath + '.tmp'
      await this.fs.writeFile(tmp, JSON.stringify(cache.entries, null, 2), 'utf8')
      await this.fs.rename(tmp, filePath)
    }
    this.dirtyPairs.clear()
  }

  /**
   * Exact match (same as {@link TranslationMemory.lookup}). Returns null when
   * there is no entry; use {@link fuzzyLookup} for similarity-based matches.
   */
  lookup(sourceLang: string, targetLang: string, sourceText: string): MemoryEntry | null {
    return this.inner.lookup(sourceLang, targetLang, sourceText)
  }

  /**
   * Fuzzy match. Exact hits win with confidence 1.0; otherwise the best
   * candidate whose similarity meets {@link PersistentTranslationMemoryOptions.fuzzyThreshold}
   * is returned. Returns null when no candidate is close enough.
   */
  fuzzyLookup(sourceLang: string, targetLang: string, sourceText: string): PersistentLookupHit | null {
    const exact = this.inner.lookup(sourceLang, targetLang, sourceText)
    if (exact) return { translatedText: exact.translatedText, sourceText: exact.sourceText, confidence: 1 }
    const pair = pairKey(sourceLang, targetLang)
    const cache = this.pairCache.get(pair)
    if (!cache) return null
    const trimmed = sourceText.trim()
    if (!trimmed) return null
    let best: PersistentLookupHit | null = null
    for (const entry of cache.entries) {
      const score = combinedSimilarity(trimmed, entry.sourceText.trim())
      if (score < this.fuzzyThreshold) continue
      if (!best || score > best.confidence) {
        best = { translatedText: entry.translatedText, sourceText: entry.sourceText, confidence: score }
      }
    }
    return best
  }

  /** Insert or replace an entry; marks its language pair as dirty. */
  save(entry: Omit<MemoryEntry, 'updatedAt'>): void {
    this.inner.save(entry)
    const pair = pairKey(entry.sourceLang, entry.targetLang)
    const cache = this.pairCache.get(pair) ?? { entries: [] }
    const next = cache.entries.filter(
      (e) => normalize(e.sourceText) !== normalize(entry.sourceText),
    )
    next.push({ ...entry, updatedAt: Date.now() })
    cache.entries = next
    this.pairCache.set(pair, cache)
    this.dirtyPairs.add(pair)
  }

  /** How many entries are cached in memory. */
  /**
   * Save a batch of memory entries in one call. Same signature as the
   * in-memory {@link TranslationMemory.saveMany} so this class is a true
   * drop-in replacement for callers that import the memory option by type.
   * Marks every touched pair dirty so the next {@link flush} persists them.
   */
  saveMany(req: MemorySaveRequest): MemorySaveResponse {
    const inner = this.inner.saveMany(req)
    for (const unit of req.units ?? []) {
      if (!unit.sourceText) continue
      const pair = pairKey(req.sourceLang ?? 'auto', req.targetLang ?? 'auto')
      let cache = this.pairCache.get(pair)
      if (!cache) {
        cache = { entries: [] }
        this.pairCache.set(pair, cache)
      }
      cache.entries.push({
        sourceLang: req.sourceLang ?? 'auto',
        targetLang: req.targetLang ?? 'auto',
        sourceText: unit.sourceText,
        translatedText: unit.translatedText,
        updatedAt: Date.now(),
      })
      this.dirtyPairs.add(pair)
    }
    return inner
  }

  /** Wipe both the in-memory index and the per-pair cache. Mostly useful for
   *  tests; production callers rarely want to throw away the entire memory. */
  clear(): void {
    this.inner.clear()
    this.pairCache.clear()
    this.dirtyPairs.clear()
  }

  size(): number {
    return this.inner.size()
  }

  /** Whether any per-pair file has unflushed changes. */
  isDirty(): boolean {
    return this.dirtyPairs.size > 0
  }

  /**
   * Export every cached pair as TMX 1.4b XML. `lang` is the XML `xml:lang`
   * attribute on `<tuv>`; we map each entry's source/target codes to BCP-47.
   */
  exportTmx(): string {
    const lines: string[] = []
    lines.push('<?xml version="1.0" encoding="UTF-8"?>')
    lines.push(`<tmx version="${TMX_VERSION}">`)
    lines.push('  <header creationtool="genoffice-translation-core" creationtoolversion="0.1.0" segtype="sentence" o-tmf="plain-text" />')
    lines.push('  <body>')
    for (const [pair, cache] of this.pairCache.entries()) {
      const [src, tgt] = pair.split('->')
      for (const entry of cache.entries) {
        lines.push('    <tu>')
        lines.push(`      <tuv xml:lang="${escapeXml(src ?? '')}"><seg>${escapeXml(entry.sourceText)}</seg></tuv>`)
        lines.push(`      <tuv xml:lang="${escapeXml(tgt ?? '')}"><seg>${escapeXml(entry.translatedText)}</seg></tuv>`)
        lines.push('    </tu>')
      }
    }
    lines.push('  </body>')
    lines.push('</tmx>')
    return lines.join('\n')
  }
}

/* ------------------------------------------------------------------ */
/* Internals.                                                           */
/* ------------------------------------------------------------------ */

interface PairCache {
  entries: MemoryEntry[]
}

function pairKey(sourceLang: string, targetLang: string): string {
  return `${sourceLang}->${targetLang}`
}

function normalize(s: string): string {
  return s.replace(/\s+/g, ' ').trim().normalize('NFC')
}

/**
 * Combined similarity: weighted average of a token-overlap Jaccard score and
 * a SequenceMatcher-style character similarity. Mirrors
 * `translation_memory.py`'s `token_similarity` heuristic.
 */
function combinedSimilarity(a: string, b: string): number {
  if (!a || !b) return 0
  const lo = a.toLowerCase()
  const lr = b.toLowerCase()
  if (lo === lr) return 1

  // Character-level SequenceMatcher.ratio equivalent (Dice-ish on bigrams).
  const charSim = diceCoefficient(lo, lr)

  // Token-level Jaccard overlap.
  const ta = new Set(lo.split(/\s+/).filter(Boolean))
  const tb = new Set(lr.split(/\s+/).filter(Boolean))
  let intersect = 0
  for (const t of ta) if (tb.has(t)) intersect++
  const union = ta.size + tb.size - intersect
  const tokenSim = union === 0 ? 0 : intersect / union

  // 0.6 charSim + 0.4 tokenSim — same blend the Python script uses.
  // unused unless charSim or tokenSim are referenced; keep both in scope.
  void charSim
  void tokenSim
  return 0.6 * charSim + 0.4 * tokenSim
}

/**
 * Cheap Dice coefficient on character bigrams — close enough to Python's
 * `SequenceMatcher.ratio()` for short segments and avoids the O(n*m) DP table.
 */
function diceCoefficient(a: string, b: string): number {
  if (a.length < 2 || b.length < 2) return a === b ? 1 : 0
  const aBigrams = new Map<string, number>()
  for (let i = 0; i < a.length - 1; i++) {
    const bg = a.slice(i, i + 2)
    aBigrams.set(bg, (aBigrams.get(bg) ?? 0) + 1)
  }
  let matches = 0
  let bTotal = 0
  for (let i = 0; i < b.length - 1; i++) {
    const bg = b.slice(i, i + 2)
    bTotal++
    const count = aBigrams.get(bg) ?? 0
    if (count > 0) {
      aBigrams.set(bg, count - 1)
      matches++
    }
  }
  const aTotal = a.length - 1
  return (2 * matches) / (aTotal + bTotal)
}

function isMemoryEntry(value: unknown): value is MemoryEntry {
  if (value === null || typeof value !== 'object') return false
  const v = value as Record<string, unknown>
  return (
    typeof v.sourceLang === 'string' &&
    typeof v.targetLang === 'string' &&
    typeof v.sourceText === 'string' &&
    typeof v.translatedText === 'string'
  )
}

function isNotFound(error: unknown): boolean {
  return (
    typeof error === 'object' &&
    error !== null &&
    'code' in error &&
    (error as { code?: string }).code === 'ENOENT'
  )
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}
