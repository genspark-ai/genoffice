/**
 * Search utilities (main process) — gsk (Genspark CLI) first, then Serper Google API,
 * then Tavily and Parallel, with DuckDuckGo as the keyless last resort. Runs in the main process
 * (Node fetch / child process) to avoid renderer CORS; the Serper key reuses SERPER_API_KEY,
 * the Tavily key reuses TAVILY_API_KEY and Parallel uses PARALLEL_API_KEY.
 * For gsk auth see ./gsk.ts (`gsk login` or GSK_API_KEY).
 */

import {
  asRecord,
  isCopyrightHost,
  safeHost,
  type ImageSearchResult,
  type WebSearchResult,
} from './shared'
import { gskImageSearch, gskWebSearch, hasGskAuth } from './gsk'
import { parallelMcpSearch } from './parallel-mcp'

export type { ImageSearchResult, WebSearchResult } from './shared'
export * from './gsk'
export * from './genoffice-auth'
export * from './media-tools'
export * from './search-tools'

const SERPER_KEY = () => process.env.SERPER_API_KEY ?? ''
const TAVILY_KEY = () => process.env.TAVILY_API_KEY ?? ''
const PARALLEL_KEY = () => process.env.PARALLEL_API_KEY ?? ''

/**
 * Backend selection for one search. Keys default to the SERPER_API_KEY /
 * TAVILY_API_KEY / PARALLEL_API_KEY env vars; settings-driven callers (search-tools.ts) pass the
 * user's key and turn gsk off so the chosen backend runs first.
 */
export interface SearchOptions {
  /** false = skip the Genspark backend (cloud tools off, or a BYOK search provider is active) */
  useGsk?: boolean
  serperKey?: string
  tavilyKey?: string
  parallelKey?: string
  /** which backend to try first (default serper); selecting parallel also enables its free MCP */
  prefer?: 'serper' | 'tavily' | 'parallel'
}

function normalizeOptions(opts: boolean | SearchOptions | undefined): Required<SearchOptions> {
  const o = typeof opts === 'boolean' ? { useGsk: opts } : (opts ?? {})
  return {
    useGsk: o.useGsk ?? true,
    serperKey: o.serperKey ?? SERPER_KEY(),
    tavilyKey: o.tavilyKey ?? TAVILY_KEY(),
    parallelKey: o.parallelKey ?? PARALLEL_KEY(),
    prefer: o.prefer ?? 'serper',
  }
}

type WebSearchResponse = {
  results: WebSearchResult[]
  answer?: string
  method: string
  error?: string
}

/** Serper Google web search; null when the key is empty, the call fails, or nothing comes back */
async function serperWebSearch(
  key: string,
  query: string,
  maxResults: number,
): Promise<WebSearchResponse | null> {
  if (!key) return null
  try {
    const resp = await fetchWithTimeout('https://google.serper.dev/search', {
      method: 'POST',
      headers: { 'X-API-KEY': key, 'Content-Type': 'application/json' },
      body: JSON.stringify({ q: query, num: maxResults, gl: 'us', hl: 'en' }),
    })
    if (!resp.ok) return null
    const data = asRecord(await resp.json())
    const organic: unknown[] = Array.isArray(data.organic) ? data.organic : []
    const results: WebSearchResult[] = organic.slice(0, maxResults).map((item) => {
      const o = asRecord(item)
      return {
        title: String(o.title ?? ''),
        url: String(o.link ?? ''),
        snippet: String(o.snippet ?? ''),
      }
    })
    const answerBox = asRecord(data.answerBox)
    const answerRaw =
      answerBox.answer || answerBox.snippet || asRecord(data.knowledgeGraph).description
    const answer = typeof answerRaw === 'string' && answerRaw ? answerRaw : undefined
    if (!results.length) return null
    return answer !== undefined
      ? { results, answer, method: 'serper' }
      : { results, method: 'serper' }
  } catch {
    return null
  }
}

async function tavilyWebSearch(
  key: string,
  query: string,
  maxResults: number,
): Promise<WebSearchResponse | null> {
  if (!key) return null
  try {
    const resp = await fetchWithTimeout('https://api.tavily.com/search', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        api_key: key,
        query,
        max_results: maxResults,
        include_answer: true,
      }),
    })
    if (!resp.ok) return null
    const data = asRecord(await resp.json())
    const raw: unknown[] = Array.isArray(data.results) ? data.results : []
    const results: WebSearchResult[] = raw.slice(0, maxResults).map((item) => {
      const o = asRecord(item)
      return {
        title: String(o.title ?? ''),
        url: String(o.url ?? ''),
        snippet: String(o.content ?? ''),
      }
    })
    const answerRaw = data.answer
    const answer = typeof answerRaw === 'string' && answerRaw ? answerRaw : undefined
    if (!results.length) return null
    return answer !== undefined
      ? { results, answer, method: 'tavily' }
      : { results, method: 'tavily' }
  } catch {
    return null
  }
}

/** Parallel's API and free Search MCP return source excerpts, not a synthesized answer. */
async function parallelWebSearch(
  key: string,
  query: string,
  maxResults: number,
  allowFreeMcp: boolean,
): Promise<WebSearchResponse | null> {
  key = key.trim()
  if (!key && !allowFreeMcp) return null
  try {
    let data: Record<string, unknown>
    if (key) {
      const resp = await fetchWithTimeout('https://api.parallel.ai/v1/search', {
        method: 'POST',
        headers: { 'x-api-key': key, 'Content-Type': 'application/json' },
        body: JSON.stringify({ search_queries: [query], mode: 'fast' }),
      })
      if (!resp.ok) return null
      data = asRecord(await resp.json())
    } else {
      data = asRecord(await parallelMcpSearch(query))
    }
    const raw: unknown[] = Array.isArray(data.results) ? data.results : []
    const results: WebSearchResult[] = []
    for (const item of raw) {
      const result = asRecord(item)
      if (typeof result.url !== 'string' || !/^https?:\/\//i.test(result.url)) continue
      const excerpts: unknown[] = Array.isArray(result.excerpts) ? result.excerpts : []
      results.push({
        title: typeof result.title === 'string' && result.title ? result.title : result.url,
        url: result.url,
        snippet: excerpts
          .filter((excerpt): excerpt is string => typeof excerpt === 'string')
          .join('\n'),
      })
    }
    // v1 search has no result-count parameter; apply GenOffice's limit locally.
    const limited = results.slice(0, maxResults)
    return limited.length ? { results: limited, method: 'parallel' } : null
  } catch {
    return null
  }
}

// ── Web search ──────────────────────────────────────────────────────

/**
 * Model-driven search args arrive unvalidated: clamp the result count to a
 * finite 1..20 and truncate the query so a wild maxResults cannot inflate
 * backend cost/loops and a megabyte query cannot flood request bodies.
 */
function normalizeSearchArgs(
  query: string,
  maxResults: number,
  fallback: number,
): { query: string; max: number } {
  const max = Number.isFinite(maxResults)
    ? Math.min(Math.max(1, Math.floor(maxResults)), 20)
    : fallback
  return { query: typeof query === 'string' ? query.slice(0, 500) : '', max }
}

export async function webSearch(
  query: string,
  maxResults = 6,
  options: boolean | SearchOptions = true,
): Promise<WebSearchResponse> {
  const o = normalizeOptions(options)
  const { query: q, max } = normalizeSearchArgs(query, maxResults, 6)
  // useGsk=false: the user turned Genspark cloud tools off or picked their own
  // search key — skip straight to the keyed/free backends
  if (o.useGsk && hasGskAuth()) {
    try {
      const r = await gskWebSearch(q, max)
      if (r.results.length) return { ...r, method: 'gsk' }
    } catch {
      /* fall back to Serper/Tavily/Parallel/DuckDuckGo */
    }
  }
  const keyed = {
    serper: () => serperWebSearch(o.serperKey, q, max),
    tavily: () => tavilyWebSearch(o.tavilyKey, q, max),
    parallel: () => parallelWebSearch(o.parallelKey, q, max, o.prefer === 'parallel'),
  }
  const order = [
    o.prefer,
    ...(['serper', 'tavily', 'parallel'] as const).filter((id) => id !== o.prefer),
  ]
  for (const id of order) {
    const r = await keyed[id]()
    if (r) return r
  }
  try {
    return { results: await duckWebSearch(q, max), method: 'duckduckgo' }
  } catch (err) {
    // an unreachable backend must not read as an empty result set
    return { results: [], method: 'error', error: `duckduckgo: ${String(err)}` }
  }
}

// ── Image search ────────────────────────────────────────────────────

export async function imageSearch(
  query: string,
  maxResults = 8,
  options: boolean | SearchOptions = true,
): Promise<{
  images: ImageSearchResult[]
  method: string
  error?: string
}> {
  const o = normalizeOptions(options)
  const { query: q, max } = normalizeSearchArgs(query, maxResults, 8)
  if (o.useGsk && hasGskAuth()) {
    try {
      const images = await gskImageSearch(q, max)
      if (images.length) return { images, method: 'gsk' }
    } catch {
      /* fall back to Serper/DuckDuckGo */
    }
  }
  // Tavily and Parallel have no image endpoint; Serper is the only keyed image backend
  const key = o.serperKey
  if (key) {
    try {
      const resp = await fetchWithTimeout('https://google.serper.dev/images', {
        method: 'POST',
        headers: { 'X-API-KEY': key, 'Content-Type': 'application/json' },
        body: JSON.stringify({ q, num: Math.min(max, 10), gl: 'us', hl: 'en' }),
      })
      if (resp.ok) {
        const data = asRecord(await resp.json())
        const raw: unknown[] = Array.isArray(data.images) ? data.images : []
        const images: ImageSearchResult[] = []
        for (const item of raw) {
          const img = asRecord(item)
          const imageUrl = String(img.imageUrl ?? img.original ?? '')
          if (!imageUrl) continue
          if (isCopyrightHost(imageUrl)) continue
          const entry: ImageSearchResult = {
            title: String(img.title ?? ''),
            imageUrl,
            sourceUrl: String(img.link ?? ''),
            source: String(img.source ?? safeHost(img.link)),
          }
          if (typeof img.imageWidth === 'number') entry.width = img.imageWidth
          if (typeof img.imageHeight === 'number') entry.height = img.imageHeight
          images.push(entry)
          if (images.length >= max) break
        }
        if (images.length) return { images, method: 'serper' }
      }
    } catch {
      /* fall back to DuckDuckGo */
    }
  }
  try {
    return { images: await duckImageSearch(q, max), method: 'duckduckgo' }
  } catch (err) {
    // an unreachable backend must not read as an empty gallery
    return { images: [], method: 'error', error: `duckduckgo: ${String(err)}` }
  }
}

// ── DuckDuckGo fallback (no key / quota exhausted) ──────────────────
// These throw on network/HTTP failure so the caller can distinguish
// "backend unreachable" from a genuinely empty result set.

// short timeout: an unreachable backend should fail fast so the next one gets its turn
const FALLBACK_TIMEOUT_MS = 5000

const BROWSER_HEADERS = {
  'User-Agent':
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
  'Accept-Language': 'en-US,en;q=0.9',
}

async function duckWebSearch(query: string, maxResults: number): Promise<WebSearchResult[]> {
  // DuckDuckGo HTML endpoint (lightweight, no key needed)
  const resp = await fetchWithTimeout(
    `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`,
    { headers: BROWSER_HEADERS, timeoutMs: FALLBACK_TIMEOUT_MS },
  )
  if (!resp.ok) throw new Error(`http ${resp.status}`)
  const html = await resp.text()
  const results: WebSearchResult[] = []
  const re = /<a[^>]*class="result__a"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(html)) !== null && results.length < maxResults) {
    const url = decodeDuckUrl(m[1]!)
    const title = stripTags(m[2]!)
    if (url && title) results.push({ title, url, snippet: '' })
  }
  return results
}

async function duckImageSearch(query: string, maxResults: number): Promise<ImageSearchResult[]> {
  // DuckDuckGo i.js needs a vqd token, so it takes two steps
  const tokenResp = await fetchWithTimeout(
    `https://duckduckgo.com/?q=${encodeURIComponent(query)}`,
    { headers: BROWSER_HEADERS, timeoutMs: FALLBACK_TIMEOUT_MS },
  )
  if (!tokenResp.ok) throw new Error(`http ${tokenResp.status}`)
  const tokenHtml = await tokenResp.text()
  const vqd = /vqd=["']?([\d-]+)["']?/.exec(tokenHtml)?.[1]
  if (!vqd) throw new Error('no vqd token')
  const resp = await fetchWithTimeout(
    `https://duckduckgo.com/i.js?l=us-en&o=json&q=${encodeURIComponent(query)}&vqd=${vqd}`,
    {
      headers: { ...BROWSER_HEADERS, Referer: 'https://duckduckgo.com/' },
      timeoutMs: FALLBACK_TIMEOUT_MS,
    },
  )
  if (!resp.ok) throw new Error(`http ${resp.status}`)
  const data = asRecord(await resp.json())
  const list: unknown[] = Array.isArray(data.results) ? data.results : []
  const out: ImageSearchResult[] = []
  for (const item of list.slice(0, maxResults)) {
    const img = asRecord(item)
    const imageUrl = String(img.image ?? '')
    if (!imageUrl || isCopyrightHost(imageUrl)) continue
    const entry: ImageSearchResult = {
      title: String(img.title ?? ''),
      imageUrl,
      sourceUrl: String(img.url ?? ''),
      source: safeHost(img.url),
    }
    if (typeof img.width === 'number') entry.width = img.width
    if (typeof img.height === 'number') entry.height = img.height
    out.push(entry)
  }
  return out
}

// ── utils ───────────────────────────────────────────────────────────

async function fetchWithTimeout(
  url: string,
  init: RequestInit & { timeoutMs?: number } = {},
): Promise<Response> {
  const controller = new AbortController()
  const t = setTimeout(() => controller.abort(), init.timeoutMs ?? 15000)
  try {
    return await fetch(url, { ...init, signal: controller.signal })
  } finally {
    clearTimeout(t)
  }
}

function stripTags(s: string): string {
  return s
    .replace(/<[^>]+>/g, '')
    .replace(/&quot;/g, '"')
    .replace(/&#x27;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .trim()
}

function decodeDuckUrl(href: string): string {
  // DuckDuckGo result links are often /l/?uddg=<encoded>
  const m = /[?&]uddg=([^&]+)/.exec(href)
  if (m) return decodeURIComponent(m[1]!)
  return href.startsWith('http') ? href : ''
}
