import type { TranslationUnit } from './types'

/**
 * Tiny in-memory translation memory keyed by `${sourceLang}::${targetLang}::${normalizedSource}`.
 *
 * Holds up to `maxEntries` per call site; saves the user a model round-trip
 * when they retranslating the same sentence. The Dataflare bridge (when wired)
 * is the source of truth for cross-device memory; this is a desktop / web
 * fallback that always succeeds and never blocks on the network.
 */
export interface MemoryEntry {
  sourceLang: string
  targetLang: string
  sourceText: string
  translatedText: string
  /** unix millis */
  updatedAt: number
}

export interface MemorySaveRequest {
  scene: string
  sourceLang: string
  targetLang: string
  units: Array<{ unitId: string; sourceText: string; translatedText: string }>
}

export interface MemorySaveResponse {
  ok: boolean
  savedCount: number
  skippedCount: number
  error?: string
}

export class TranslationMemory {
  private readonly entries = new Map<string, MemoryEntry>()
  private readonly maxEntries: number

  constructor(opts: { maxEntries?: number } = {}) {
    // The caller picks the capacity (tests want 10, production wants 2048);
    // a tiny floor just keeps an accidentally-passed 0 from being a no-op.
    this.maxEntries = Math.max(1, opts.maxEntries ?? 2048)
  }

  private keyOf(sourceLang: string, targetLang: string, sourceText: string): string {
    return `${sourceLang}::${targetLang}::${normalize(sourceText)}`
  }

  /** Look up a unit by exact source text + language pair. */
  lookup(sourceLang: string, targetLang: string, sourceText: string): MemoryEntry | null {
    if (!sourceText || !sourceText.trim()) return null
    const entry = this.entries.get(this.keyOf(sourceLang, targetLang, sourceText))
    if (!entry) return null
    entry.updatedAt = Date.now() // LRU touch
    return entry
  }

  /** Insert or update a single translation. */
  save(entry: Omit<MemoryEntry, 'updatedAt'>): void {
    const key = this.keyOf(entry.sourceLang, entry.targetLang, entry.sourceText)
    this.entries.set(key, { ...entry, updatedAt: Date.now() })
    this.evictIfNeeded()
  }

  saveMany(req: MemorySaveRequest): MemorySaveResponse {
    let saved = 0
    let skipped = 0
    for (const unit of req.units) {
      if (!unit.translatedText || !unit.translatedText.trim()) {
        skipped++
        continue
      }
      this.save({
        sourceLang: req.sourceLang,
        targetLang: req.targetLang,
        sourceText: unit.sourceText,
        translatedText: unit.translatedText,
      })
      saved++
    }
    return { ok: true, savedCount: saved, skippedCount: skipped }
  }

  size(): number {
    return this.entries.size
  }

  clear(): void {
    this.entries.clear()
  }

  private evictIfNeeded(): void {
    if (this.entries.size <= this.maxEntries) return
    // drop the oldest 10% to amortize the cost
    const drop = Math.max(1, Math.ceil(this.maxEntries * 0.1))
    const sorted = [...this.entries.entries()].sort((a, b) => a[1].updatedAt - b[1].updatedAt)
    for (let i = 0; i < drop && i < sorted.length; i++) {
      const entry = sorted[i]
      if (entry) this.entries.delete(entry[0])
    }
  }
}

/** Trim surrounding whitespace, collapse runs, and NFC-normalize for stable keys. */
function normalize(s: string): string {
  return s.replace(/\s+/g, ' ').trim().normalize('NFC')
}

/** Build a memory key for a single unit. */
export function unitMemoryKey(
  unit: TranslationUnit,
  sourceLang: string,
  targetLang: string,
): string {
  return `${sourceLang}::${targetLang}::${normalize(unit.sourceText)}`
}
