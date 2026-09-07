/**
 * Full-text search — minimal in-memory index built from `search:index`
 * calls. Phase 1.x keeps this simple; Phase 4 (LUM-551 plan) will move
 * to SQLite / FTS if needed.
 */
import { registerHandle, SEARCH_INDEX } from '../common/index.js'

export function registerSearchHandlers(): void {
  registerHandle('search:index', (_event: unknown, args: unknown) => {
    const { id, type, title, content, tags } = args as {
      id: string
      type: string
      title: string
      content: string
      tags?: string[]
    }

    SEARCH_INDEX.set(id, {
      id,
      type,
      title,
      content,
      tags: tags || [],
      createdAt: Date.now(),
    })

    return { ok: true, indexed: SEARCH_INDEX.size }
  })

  registerHandle('search:query', (_event: unknown, args: unknown) => {
    const { query, type, limit, offset } = args as {
      query: string
      type?: string
      limit?: number
      offset?: number
    }

    const maxResults = limit || 20
    const startOffset = offset || 0
    const queryLower = query.toLowerCase()

    let results = [...SEARCH_INDEX.values()].filter(item => {
      if (type && item.type !== type) return false
      return (
        item.title.toLowerCase().includes(queryLower) ||
        item.content.toLowerCase().includes(queryLower) ||
        item.tags.some(tag => tag.toLowerCase().includes(queryLower))
      )
    })

    const total = results.length
    results = results.slice(startOffset, startOffset + maxResults)

    return {
      results: results.map(r => ({
        id: r.id,
        type: r.type,
        title: r.title,
        snippet: r.content.slice(0, 200) + (r.content.length > 200 ? '...' : ''),
        score: 1.0,
      })),
      total,
      limit: maxResults,
      offset: startOffset,
    }
  })

  registerHandle('search:delete', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    if (!SEARCH_INDEX.has(id)) return { ok: false, error: 'Not found' }
    SEARCH_INDEX.delete(id)
    return { ok: true }
  })
}
