/**
 * Zotero's local API (Zotero 7+, Settings ▸ Advanced ▸ "Allow other applications
 * on this computer to communicate with Zotero"): the read-only Zotero Web API
 * served by the desktop app on 127.0.0.1:23119/api, over the user's own
 * library. Docs reads it for Quick Cite and the library pane; nothing leaves
 * the machine, and Zotero itself still formats and refreshes the citations
 * through its word-processor integration.
 */
import type {
  ZoteroLocalCiteData,
  ZoteroLocalCollection,
  ZoteroLocalItem,
  ZoteroLocalResult,
} from '../shared/ipc'

const BASE = 'http://127.0.0.1:23119/api/users/0'
const TIMEOUT_MS = 8_000
const PAGE = 100
/** a library this size is listed whole; past it the pane pages through search */
const MAX_COLLECTIONS = 2_000
/** item keys are 8 characters from Zotero's own alphabet */
const ITEM_KEY_RE = /^[23456789ABCDEFGHIJKLMNPQRSTUVWXYZ]{8}$/
/** a CSL style id as Zotero names it (apa, ieee, chicago-author-date, …) */
const STYLE_RE = /^[a-z0-9][a-z0-9-]{0,120}$/

type Fetch = typeof fetch

interface ApiItem {
  key?: unknown
  library?: { id?: unknown }
  meta?: { creatorSummary?: unknown; parsedDate?: unknown }
  data?: { itemType?: unknown; title?: unknown; parentItem?: unknown }
  citation?: unknown
  csljson?: unknown
}

interface ApiCollection {
  key?: unknown
  data?: { name?: unknown; parentCollection?: unknown }
}

class LocalApiError extends Error {
  constructor(
    readonly code: 'off' | 'unreachable' | 'bad-response',
    message: string,
  ) {
    super(message)
  }
}

async function getJson(path: string, fetchImpl: Fetch): Promise<{ body: unknown; total: number }> {
  let response: Response
  try {
    response = await fetchImpl(`${BASE}${path}`, {
      headers: { 'Zotero-API-Version': '3' },
      signal: AbortSignal.timeout(TIMEOUT_MS),
    })
  } catch {
    throw new LocalApiError('unreachable', 'Zotero is not running')
  }
  // with the setting off, Zotero answers the endpoint itself with 403 (older builds 404)
  if (response.status === 403 || response.status === 404) {
    throw new LocalApiError('off', 'Zotero local API is off')
  }
  if (!response.ok) throw new LocalApiError('bad-response', `Zotero answered ${response.status}`)
  const total = Number(response.headers.get('Total-Results'))
  return { body: await response.json(), total: Number.isFinite(total) ? total : 0 }
}

function asItems(body: unknown): ApiItem[] {
  return Array.isArray(body) ? (body as ApiItem[]) : []
}

/** the regular items a citation can point at: no attachments, notes or annotations */
function citable(item: ApiItem): boolean {
  const type = item.data?.itemType
  return (
    typeof item.key === 'string' &&
    type !== 'attachment' &&
    type !== 'note' &&
    type !== 'annotation' &&
    !item.data?.parentItem
  )
}

function toItem(item: ApiItem): ZoteroLocalItem {
  const year = /\d{4}/.exec(String(item.meta?.parsedDate ?? ''))?.[0] ?? ''
  return {
    key: String(item.key),
    libraryId: Number(item.library?.id) || 0,
    title: typeof item.data?.title === 'string' ? item.data.title : '',
    creators: typeof item.meta?.creatorSummary === 'string' ? item.meta.creatorSummary : '',
    year,
    itemType: typeof item.data?.itemType === 'string' ? item.data.itemType : '',
  }
}

function failure<T>(error: unknown): ZoteroLocalResult<T> {
  if (error instanceof LocalApiError)
    return { ok: false, errorCode: error.code, error: error.message }
  return { ok: false, errorCode: 'bad-response', error: String(error) }
}

/** items whose title, creators or year match `query`; the most recently modified first */
export async function searchZoteroItems(
  query: string,
  limit = 25,
  fetchImpl: Fetch = fetch,
): Promise<ZoteroLocalResult<ZoteroLocalItem[]>> {
  const q = query.trim().slice(0, 200)
  if (!q) return { ok: true, value: [] }
  try {
    const size = Math.max(1, Math.min(PAGE, Math.trunc(limit)))
    const { body } = await getJson(
      `/items/top?q=${encodeURIComponent(q)}&qmode=titleCreatorYear&sort=dateModified&direction=desc&limit=${size}`,
      fetchImpl,
    )
    return { ok: true, value: asItems(body).filter(citable).map(toItem) }
  } catch (error) {
    return failure(error)
  }
}

/** every collection of the library, flat, each with its parent's key */
export async function listZoteroCollections(
  fetchImpl: Fetch = fetch,
): Promise<ZoteroLocalResult<ZoteroLocalCollection[]>> {
  try {
    const collections: ZoteroLocalCollection[] = []
    for (let start = 0; start < MAX_COLLECTIONS; start += PAGE) {
      const { body, total } = await getJson(`/collections?limit=${PAGE}&start=${start}`, fetchImpl)
      const page = Array.isArray(body) ? (body as ApiCollection[]) : []
      for (const entry of page) {
        if (typeof entry.key !== 'string') continue
        const parent = entry.data?.parentCollection
        collections.push({
          key: entry.key,
          name: typeof entry.data?.name === 'string' ? entry.data.name : '',
          parent: typeof parent === 'string' ? parent : null,
        })
      }
      if (page.length < PAGE || start + PAGE >= total) break
    }
    collections.sort((a, b) => a.name.localeCompare(b.name))
    return { ok: true, value: collections }
  } catch (error) {
    return failure(error)
  }
}

/** the items of one collection (or the whole library for null), optionally filtered */
export async function listZoteroCollectionItems(
  collectionKey: string | null,
  query = '',
  fetchImpl: Fetch = fetch,
): Promise<ZoteroLocalResult<ZoteroLocalItem[]>> {
  if (collectionKey !== null && !ITEM_KEY_RE.test(collectionKey)) {
    return { ok: false, errorCode: 'bad-request', error: 'Invalid collection key' }
  }
  try {
    const scope = collectionKey ? `/collections/${collectionKey}/items/top` : '/items/top'
    const q = query.trim().slice(0, 200)
    const filter = q ? `&q=${encodeURIComponent(q)}&qmode=titleCreatorYear` : ''
    const { body } = await getJson(`${scope}?limit=${PAGE}&sort=title${filter}`, fetchImpl)
    return { ok: true, value: asItems(body).filter(citable).map(toItem) }
  } catch (error) {
    return failure(error)
  }
}

/**
 * What a citation field needs for each item: its CSL-JSON and the citation
 * Zotero formats for it in `style`, keyed as the request listed them.
 */
export async function zoteroCiteData(
  keys: string[],
  style: string,
  fetchImpl: Fetch = fetch,
): Promise<ZoteroLocalResult<ZoteroLocalCiteData[]>> {
  const wanted = [...new Set(keys)].filter((key) => ITEM_KEY_RE.test(key)).slice(0, 50)
  if (wanted.length === 0) return { ok: false, errorCode: 'bad-request', error: 'No items' }
  const styleId = STYLE_RE.test(style) ? style : 'apa'
  try {
    const { body } = await getJson(
      `/items?itemKey=${wanted.join(',')}&include=csljson,citation&style=${styleId}`,
      fetchImpl,
    )
    const byKey = new Map<string, ZoteroLocalCiteData>()
    for (const item of asItems(body)) {
      if (typeof item.key !== 'string') continue
      let csl: unknown = item.csljson
      if (typeof csl === 'string') {
        try {
          csl = JSON.parse(csl)
        } catch {
          csl = null
        }
      }
      const itemData = Array.isArray(csl) ? csl[0] : csl
      if (!itemData || typeof itemData !== 'object') continue
      byKey.set(item.key, {
        key: item.key,
        libraryId: Number(item.library?.id) || 0,
        itemData: itemData as Record<string, unknown>,
        citation: typeof item.citation === 'string' ? item.citation : '',
      })
    }
    const missing = wanted.filter((key) => !byKey.has(key))
    if (missing.length > 0) {
      return {
        ok: false,
        errorCode: 'bad-response',
        error: `Items not found: ${missing.join(', ')}`,
      }
    }
    return { ok: true, value: wanted.map((key) => byKey.get(key)!) }
  } catch (error) {
    return failure(error)
  }
}
