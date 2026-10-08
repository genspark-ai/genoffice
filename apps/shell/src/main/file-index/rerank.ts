import { basename, dirname } from 'node:path'
import type { DecisionEndpoint, FileSearchRerank, FileSearchSettings } from '../../shared/home-api'
import {
  evaluate,
  isLocalEndpoint,
  MAX_DOCS,
  type DecisionCallOptions,
  type JevTransport,
} from './decider'
import type { FileIndexStore } from './store'

const EXCERPT_CHARS = 1200
const CACHE_TTL_MS = 30 * 60 * 1000
const CACHE_MAX = 200

export const DEFAULT_FILE_SEARCH_SETTINGS: FileSearchSettings = {
  rerank: false,
  endpoint: 'openrouter',
  keys: {
    openrouter: '',
    direct: '',
    perplexity: '',
    cloudflare: '',
    kev: '',
    rizzo: '',
    custom: '',
  },
  customBaseUrl: '',
  customModel: '',
  cloudflareAccountId: '',
  cloudflareModel: '@cf/cloudflare/clef',
}

const ENDPOINT_IDS: readonly DecisionEndpoint[] = [
  'openrouter',
  'direct',
  'perplexity',
  'cloudflare',
  'kev',
  'rizzo',
  'custom',
]

/**
 * Accepts the current settings shape and the pre-decision-model one
 * (`jevEndpoint`/`jevKeys` with only the two hosted Jev routes) so existing
 * app-settings.json files keep working; anything unknown falls back to the
 * OpenRouter default.
 */
export function normalizeFileSearchSettings(raw: unknown): FileSearchSettings {
  const r = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const keysRaw = (r.keys ?? r.jevKeys ?? {}) as Record<string, unknown>
  const key = (v: unknown) => (typeof v === 'string' ? v.trim().slice(0, 512) : '')
  const text = (v: unknown, n: number) => (typeof v === 'string' ? v.trim().slice(0, n) : '')
  const endpointRaw = (r.endpoint ?? r.jevEndpoint) as unknown
  const endpoint: DecisionEndpoint = ENDPOINT_IDS.find((id) => id === endpointRaw) ?? 'openrouter'
  const keys = Object.fromEntries(
    ENDPOINT_IDS.map((id) => [id, key(keysRaw[id])]),
  ) as FileSearchSettings['keys']
  return {
    rerank: r.rerank === true,
    endpoint,
    keys,
    customBaseUrl: text(r.customBaseUrl, 512),
    customModel: text(r.customModel, 128),
    cloudflareAccountId: text(r.cloudflareAccountId, 128),
    cloudflareModel: text(r.cloudflareModel, 256) || DEFAULT_FILE_SEARCH_SETTINGS.cloudflareModel,
  }
}

/** Per-call view of the settings the decision client needs. */
function callOptions(s: FileSearchSettings): DecisionCallOptions {
  return {
    endpoint: s.endpoint,
    key: s.keys[s.endpoint] ?? '',
    customBaseUrl: s.customBaseUrl,
    customModel: s.customModel,
    cloudflareAccountId: s.cloudflareAccountId,
    cloudflareModel: s.cloudflareModel,
  }
}

/**
 * Reranks the hits the home screen is showing with a decision model; the caller
 * names the candidates so the judged set is exactly the displayed one.
 * Judgements are cached per query and candidate set for half an hour so
 * retyping the same words does not bill twice. Any failure yields null and the
 * caller keeps the local order.
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
    const opts = callOptions(settings)
    // local servers (Kev, Rizzo Flow, a custom URL) run without credentials
    if (!settings.rerank || (!isLocalEndpoint(settings.endpoint) && !opts.key)) return null
    if (settings.endpoint === 'custom' && !settings.customBaseUrl.trim()) return null
    if (settings.endpoint === 'cloudflare' && !settings.cloudflareAccountId.trim()) return null
    const hits = this.store.excerptsFor(paths.slice(0, MAX_DOCS), q, EXCERPT_CHARS)
    if (hits.length < 2) return null
    const cacheKey = [
      settings.endpoint,
      settings.customBaseUrl,
      settings.customModel,
      settings.cloudflareAccountId,
      settings.cloudflareModel,
      q,
      ...hits.map((h) => h.path),
    ].join('\n')
    const cached = this.cache.get(cacheKey)
    if (cached && Date.now() - cached.at < CACHE_TTL_MS) return cached.result
    const pending = this.inflight.get(cacheKey)
    if (pending) return pending
    const call = this.judge(q, hits, opts, cacheKey).finally(() => this.inflight.delete(cacheKey))
    this.inflight.set(cacheKey, call)
    return call
  }

  private async judge(
    q: string,
    hits: ReturnType<FileIndexStore['excerptsFor']>,
    opts: DecisionCallOptions,
    cacheKey: string,
  ): Promise<FileSearchRerank | null> {
    const docs = hits.map((h) => ({
      title: h.name,
      heading: basename(dirname(h.path)),
      text: h.excerpt,
    }))
    let scores: number[]
    try {
      scores = (await evaluate(q, docs, opts, this.send)).scores
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

const PROBE_DOCS = [
  { title: 'a.md', heading: '', text: 'The connection test document.' },
  { title: 'b.md', heading: '', text: 'An unrelated note.' },
]

/** what the settings UI shows for each failure the decision client can report */
const PROBE_ERRORS: Record<string, string> = {
  'missing-key': 'Enter an API key',
  'missing-url': 'Enter the server URL',
  'missing-account': 'Enter the Cloudflare account ID',
  'bad-url': 'Server URL is not a valid address',
  'insecure-url': 'Server URL must be https:// (http:// only for this machine)',
  'unsupported-model': 'Cloudflare only serves clef and clef-flash',
  'empty-request': 'Nothing to judge',
  'rate-limit': 'Rate limited — try again shortly',
  'response-too-large': 'The reply was too large',
  'invalid-response': 'The endpoint returned an unexpected reply',
  'model-mismatch': 'The endpoint answered with a different model',
  'provider-warning': 'The provider reported a warning',
  cancelled: 'The call timed out',
}

/** the settings-UI connection test: one two-document judgement against the given settings */
export async function probeDecision(
  settings: FileSearchSettings,
  send?: JevTransport,
): Promise<{ ok: boolean; error?: string }> {
  const opts = callOptions(settings)
  if (!isLocalEndpoint(settings.endpoint) && !opts.key.trim())
    return { ok: false, error: PROBE_ERRORS['missing-key'] }
  try {
    await evaluate('connection test', PROBE_DOCS, opts, send)
    return { ok: true }
  } catch (e) {
    const code = e instanceof Error ? e.message : String(e)
    const http = /^http-(\d+)$/.exec(code)
    return { ok: false, error: PROBE_ERRORS[code] ?? (http ? `HTTP ${http[1]}` : code) }
  }
}
