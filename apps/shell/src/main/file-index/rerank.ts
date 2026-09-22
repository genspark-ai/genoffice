import { basename, dirname } from 'node:path'
import type { FileSearchRerank, FileSearchSettings, JevEndpoint } from '../../shared/home-api'
import { evaluate, MAX_DOCS, type JevTransport } from './jev'
import type { FileIndexStore } from './store'

const EXCERPT_CHARS = 1200
const CACHE_TTL_MS = 30 * 60 * 1000
const CACHE_MAX = 200

export const DEFAULT_FILE_SEARCH_SETTINGS: FileSearchSettings = {
  rerank: false,
  jevEndpoint: 'openrouter',
  jevKeys: { openrouter: '', direct: '' },
}

export function normalizeFileSearchSettings(raw: unknown): FileSearchSettings {
  const r = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const keys = (r.jevKeys && typeof r.jevKeys === 'object' ? r.jevKeys : {}) as Record<
    string,
    unknown
  >
  const key = (v: unknown) => (typeof v === 'string' ? v.trim().slice(0, 512) : '')
  return {
    rerank: r.rerank === true,
    jevEndpoint: r.jevEndpoint === 'direct' ? 'direct' : 'openrouter',
    jevKeys: { openrouter: key(keys.openrouter), direct: key(keys.direct) },
  }
}

/**
 * Reranks the hits the home screen is showing with Jev; the caller names the
 * candidates so the judged set is exactly the displayed one. Judgements are
 * cached per query and candidate set for half an hour so retyping the same
 * words does not bill twice. Any failure yields null and the caller keeps the
 * local order.
 */
export class SearchReranker {
  private readonly cache = new Map<string, { at: number; result: FileSearchRerank }>()
  /** a second request for the same key while the first is out joins it instead of billing again */
  private readonly inflight = new Map<string, Promise<FileSearchRerank | null>>()

  constructor(
    private readonly store: FileIndexStore,
    private readonly send?: JevTransport,
  ) {}

  async rerank(
    q: string,
    paths: readonly string[],
    settings: FileSearchSettings,
  ): Promise<FileSearchRerank | null> {
    const key = settings.jevKeys[settings.jevEndpoint]
    if (!settings.rerank || !key) return null
    const hits = this.store.excerptsFor(paths.slice(0, MAX_DOCS), q, EXCERPT_CHARS)
    if (hits.length < 2) return null
    const cacheKey = [settings.jevEndpoint, q, ...hits.map((h) => h.path)].join('\n')
    const cached = this.cache.get(cacheKey)
    if (cached && Date.now() - cached.at < CACHE_TTL_MS) return cached.result
    const pending = this.inflight.get(cacheKey)
    if (pending) return pending
    const call = this.judge(q, hits, settings.jevEndpoint, key, cacheKey).finally(() =>
      this.inflight.delete(cacheKey),
    )
    this.inflight.set(cacheKey, call)
    return call
  }

  private async judge(
    q: string,
    hits: ReturnType<FileIndexStore['excerptsFor']>,
    endpoint: JevEndpoint,
    key: string,
    cacheKey: string,
  ): Promise<FileSearchRerank | null> {
    const docs = hits.map((h) => ({
      title: h.name,
      heading: basename(dirname(h.path)),
      text: h.excerpt,
    }))
    let scores: number[]
    try {
      scores = (await evaluate(q, docs, endpoint, key, this.send)).scores
    } catch {
      return null
    }
    const judged = hits.slice(0, scores.length)
    const order = judged
      .map((h, i) => ({ path: h.path, score: scores[i]!, i }))
      .sort((a, b) => b.score - a.score || a.i - b.i)
    const result: FileSearchRerank = {
      order: order.map((o) => o.path),
      scores: Object.fromEntries(order.map((o) => [o.path, o.score])),
    }
    if (this.cache.size >= CACHE_MAX) this.cache.delete(this.cache.keys().next().value!)
    this.cache.set(cacheKey, { at: Date.now(), result })
    return result
  }
}

export function jevEndpointOf(v: unknown): JevEndpoint {
  return v === 'direct' ? 'direct' : 'openrouter'
}

const PROBE_DOCS = [
  { title: 'a.md', heading: '', text: 'The connection test document.' },
  { title: 'b.md', heading: '', text: 'An unrelated note.' },
]

/** the settings-UI connection test: one two-document judgement against the given key */
export async function probeJev(
  endpoint: JevEndpoint,
  key: string,
  send?: JevTransport,
): Promise<{ ok: boolean; error?: string }> {
  if (!key.trim()) return { ok: false, error: 'Enter an API key' }
  try {
    await evaluate('connection test', PROBE_DOCS, endpoint, key, send)
    return { ok: true }
  } catch (e) {
    const code = e instanceof Error ? e.message : String(e)
    const http = /^http-(\d+)$/.exec(code)
    return { ok: false, error: http ? `HTTP ${http[1]}` : code }
  }
}
