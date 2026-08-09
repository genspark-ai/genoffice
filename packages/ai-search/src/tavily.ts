/**
 * Tavily backend: web search, image search and page extraction.
 *
 * Tavily answers all three from two endpoints — `/search` returns results and,
 * with `include_images`, an image list alongside them; `/extract` returns a
 * page as markdown. It is preferred over Serper/DuckDuckGo when a key is set
 * because it returns cleaned article text rather than snippets, which is what
 * the agent actually needs when writing into a document.
 *
 * The key comes from the app settings (see `setTavilyApiKey`), falling back to
 * TAVILY_API_KEY so a shell-launched dev build works without touching the UI.
 */

import {
  COPYRIGHT_HOSTS,
  asRecord,
  safeHost,
  type ImageSearchResult,
  type WebSearchResult,
} from './shared'

const TAVILY_BASE = 'https://api.tavily.com'
const DEFAULT_TIMEOUT_MS = 20_000
/** /extract fetches and renders the page, so it needs a longer budget than /search */
const EXTRACT_TIMEOUT_MS = 45_000

let explicitKey = ''

/** Key from the settings file; takes priority over the environment. */
export function setTavilyApiKey(key: string): void {
  explicitKey = (key ?? '').trim()
}

export function tavilyApiKey(): string {
  return explicitKey || (process.env.TAVILY_API_KEY ?? '').trim()
}

export function hasTavilyKey(): boolean {
  return tavilyApiKey().length > 0
}

/** One extracted page. */
export interface PageExtract {
  url: string
  title: string
  /** page content as markdown */
  content: string
  images: string[]
}

async function tavilyPost(
  path: string,
  body: Record<string, unknown>,
  timeoutMs: number,
): Promise<Record<string, unknown>> {
  const key = tavilyApiKey()
  if (!key) throw new Error('No Tavily API key configured')
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    const resp = await fetch(`${TAVILY_BASE}${path}`, {
      method: 'POST',
      signal: controller.signal,
      headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    })
    if (!resp.ok) {
      // surface the status: 401 (bad key) and 432 (quota) are the ones users hit
      const detail = (await resp.text().catch(() => '')).slice(0, 200)
      throw new Error(`Tavily HTTP ${resp.status}${detail ? `: ${detail}` : ''}`)
    }
    return asRecord(await resp.json())
  } finally {
    clearTimeout(timer)
  }
}

/** Web search. `answer` is Tavily's synthesized reply when it produced one. */
export async function tavilySearch(
  query: string,
  maxResults = 6,
): Promise<{ results: WebSearchResult[]; answer?: string }> {
  const data = await tavilyPost(
    '/search',
    {
      query,
      max_results: Math.max(1, Math.min(maxResults, 20)),
      include_answer: true,
      search_depth: 'basic',
    },
    DEFAULT_TIMEOUT_MS,
  )
  const raw: unknown[] = Array.isArray(data.results) ? data.results : []
  const results: WebSearchResult[] = raw.slice(0, maxResults).map((item) => {
    const r = asRecord(item)
    return {
      title: String(r.title ?? ''),
      url: String(r.url ?? ''),
      // Tavily's `content` is an extracted passage, longer than a search
      // snippet; cap it so a multi-result turn stays a sane prompt size
      snippet: String(r.content ?? '').slice(0, 500),
    }
  })
  const answer = typeof data.answer === 'string' && data.answer ? data.answer : undefined
  return answer !== undefined ? { results, answer } : { results }
}

/**
 * Image search. Tavily has no dedicated image endpoint — images ride along
 * with a normal search, described when `include_image_descriptions` is set.
 */
export async function tavilyImageSearch(
  query: string,
  maxResults = 8,
): Promise<ImageSearchResult[]> {
  const data = await tavilyPost(
    '/search',
    {
      query,
      // images are attached to the search, so ask for few text results
      max_results: 5,
      include_images: true,
      include_image_descriptions: true,
      search_depth: 'basic',
    },
    DEFAULT_TIMEOUT_MS,
  )
  const raw: unknown[] = Array.isArray(data.images) ? data.images : []
  const out: ImageSearchResult[] = []
  for (const item of raw) {
    // entries are either a bare URL string or { url, description }
    const img = typeof item === 'string' ? { url: item } : asRecord(item)
    const imageUrl = String(img.url ?? '')
    if (!imageUrl) continue
    if (COPYRIGHT_HOSTS.some((d) => imageUrl.toLowerCase().includes(d))) continue
    out.push({
      title: String(img.description ?? img.title ?? ''),
      imageUrl,
      sourceUrl: imageUrl,
      source: safeHost(imageUrl),
    })
    if (out.length >= maxResults) break
  }
  return out
}

/**
 * Fetch one or more pages as markdown. Tavily does the fetching, so it works
 * on pages that block a plain request and needs no browser.
 */
export async function tavilyExtract(
  urls: string[],
  opts: { advanced?: boolean } = {},
): Promise<{ pages: PageExtract[]; failed: string[] }> {
  const list = urls.filter(Boolean).slice(0, 20)
  if (!list.length) return { pages: [], failed: [] }
  const data = await tavilyPost(
    '/extract',
    {
      urls: list,
      extract_depth: opts.advanced ? 'advanced' : 'basic',
      include_images: true,
    },
    EXTRACT_TIMEOUT_MS,
  )
  const raw: unknown[] = Array.isArray(data.results) ? data.results : []
  const pages: PageExtract[] = raw.map((item) => {
    const r = asRecord(item)
    const images: unknown[] = Array.isArray(r.images) ? r.images : []
    return {
      url: String(r.url ?? ''),
      title: String(r.title ?? ''),
      content: String(r.raw_content ?? ''),
      images: images.map((i) => String(i)).filter(Boolean),
    }
  })
  const failedRaw: unknown[] = Array.isArray(data.failed_results) ? data.failed_results : []
  const failed = failedRaw.map((item) => String(asRecord(item).url ?? '')).filter(Boolean)
  return { pages, failed }
}
