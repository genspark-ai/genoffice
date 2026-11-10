/**
 * The deck outline an agent writes between the style sheet and the page specs:
 * the same shape the app's planning stage produces (core hook + one entry per
 * page). The check mirrors the rules that stage is prompted with, so a CLI
 * agent gets them as findings instead of relying on having read them.
 */
import { MAX_DECK_PAGES } from './deck-spec'

export const PAGE_TYPES = ['cover', 'content', 'data', 'closing'] as const
export type PageType = (typeof PAGE_TYPES)[number]

export const LAYOUT_VARIANTS: Readonly<Record<PageType, readonly string[]>> = {
  cover: [
    'cover_typography_hero',
    'cover_dark_minimal',
    'cover_split_color',
    'cover_full_image_overlay',
    'cover_magazine',
    'cover_split_image',
  ],
  content: [
    'left_text_right_image',
    'three_column_cards',
    'hero_big_number',
    'two_column_comparison',
    'timeline_horizontal',
    'full_image_text_overlay',
  ],
  data: ['kpi_cards_row', 'chart_with_insight', 'two_by_two_grid'],
  closing: ['closing_cta', 'closing_thank_you'],
}

export interface OutlinePage {
  title: string
  type: PageType
  layout: string
  brief: string
  image_queries: string[]
}

export interface DeckOutline {
  topic?: string
  core_hook: string
  pages: OutlinePage[]
}

export interface OutlineIssue {
  /** 0-based page index; absent for deck-level findings */
  page?: number
  level: 'error' | 'warning'
  message: string
}

export const PLACEHOLDER = /\bXX+(?:%|\b)|\blorem\b|\bTBD\b|\bTODO\b|\{\{|\[insert\b|\bN\/A\b/i
const HAS_CJK = new RegExp('[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af]')
const MIN_BRIEF = 40
/** Raw LLM outline budget; larger payloads are rejected before JSON.parse. */
const MAX_OUTLINE_RAW_CHARS = 512_000
/** Per-page image query budget; extras are dropped with a warning. */
const MAX_IMAGE_QUERIES_PER_PAGE = 8

export function parseOutline(
  raw: string,
): { ok: true; outline: DeckOutline; issues: OutlineIssue[] } | { ok: false; error: string } {
  // LLM output is untrusted: refuse a megabyte dump before JSON.parse builds
  // a giant object graph from it.
  if (typeof raw !== 'string' || raw.length > MAX_OUTLINE_RAW_CHARS) {
    return { ok: false, error: `outline too large (limit ${MAX_OUTLINE_RAW_CHARS} chars)` }
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch (e) {
    return { ok: false, error: `invalid JSON: ${e instanceof Error ? e.message : String(e)}` }
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    return {
      ok: false,
      error:
        'expected { "core_hook", "pages": [ { "title", "type", "layout", "brief", "image_queries" } ] }',
    }
  }
  const root = parsed as Record<string, unknown>
  if (!Array.isArray(root.pages) || root.pages.length === 0) {
    return { ok: false, error: 'the "pages" array is missing or empty' }
  }
  if (root.pages.length > MAX_DECK_PAGES) {
    return {
      ok: false,
      error: `too many pages (${root.pages.length}; the limit is ${MAX_DECK_PAGES})`,
    }
  }
  const issues: OutlineIssue[] = []
  const pages: OutlinePage[] = root.pages.map((p, i) => normalizePage(p, i, issues))
  const coreHook = typeof root.core_hook === 'string' ? root.core_hook.trim() : ''
  if (!coreHook)
    issues.push({
      level: 'error',
      message:
        '"core_hook" is missing: one sentence with tension, a number or a counter-intuitive contrast',
    })
  const topic = typeof root.topic === 'string' ? root.topic.trim() : undefined
  checkDeck(pages, issues)
  return { ok: true, outline: { ...(topic ? { topic } : {}), core_hook: coreHook, pages }, issues }
}

function normalizePage(raw: unknown, i: number, issues: OutlineIssue[]): OutlinePage {
  const p = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const title = typeof p.title === 'string' ? p.title.trim() : ''
  const type = typeof p.type === 'string' ? (p.type.trim() as PageType) : ('' as PageType)
  const layout = typeof p.layout === 'string' ? p.layout.trim() : ''
  const brief = typeof p.brief === 'string' ? p.brief.trim() : ''
  const queries = Array.isArray(p.image_queries) ? p.image_queries : []
  const err = (message: string) => issues.push({ page: i, level: 'error', message })
  const warn = (message: string) => issues.push({ page: i, level: 'warning', message })

  if (!title) err('"title" is missing')
  if (!(PAGE_TYPES as readonly string[]).includes(type)) {
    err(`"type" must be one of ${PAGE_TYPES.join(', ')}`)
  } else if (!layout) {
    err(`"layout" is missing; pick one of ${LAYOUT_VARIANTS[type].join(', ')}`)
  } else if (!LAYOUT_VARIANTS[type].includes(layout)) {
    err(
      `"layout" ${layout} is not a ${type} variant; pick one of ${LAYOUT_VARIANTS[type].join(', ')}`,
    )
  }
  if (!brief) err('"brief" is missing: what goes in each region, with the real facts and figures')
  else if (brief.length < MIN_BRIEF)
    warn(`"brief" is thin (${brief.length} chars); name the content of every region`)
  const placeholder = PLACEHOLDER.exec(brief) ?? PLACEHOLDER.exec(title)
  if (placeholder)
    err(
      `placeholder "${placeholder[0]}" in the copy; every figure and name comes from the material`,
    )
  if (!Array.isArray(p.image_queries))
    warn('"image_queries" is missing; use [] for a page without photos')
  const image_queries: string[] = []
  for (const q of queries) {
    if (image_queries.length >= MAX_IMAGE_QUERIES_PER_PAGE) {
      warn(`"image_queries" capped at ${MAX_IMAGE_QUERIES_PER_PAGE} per page; extras dropped`)
      break
    }
    if (typeof q !== 'string' || !q.trim()) {
      err(
        '"image_queries" entries must be non-empty strings (an English scene query or an http(s) URL)',
      )
      continue
    }
    const s = q.trim()
    if (!/^https?:\/\//i.test(s) && HAS_CJK.test(s)) {
      warn(`image query "${s}" should be an English phrase naming a concrete scene`)
    }
    image_queries.push(s)
  }
  return { title, type, layout, brief, image_queries }
}

function checkDeck(pages: OutlinePage[], issues: OutlineIssue[]): void {
  if (pages[0] && pages[0].type !== 'cover') {
    issues.push({ page: 0, level: 'warning', message: 'the first page is not a cover' })
  }
  const last = pages.length - 1
  if (pages.length > 2 && pages[last] && pages[last].type !== 'closing') {
    issues.push({ page: last, level: 'warning', message: 'the last page is not a closing page' })
  }
  const body = pages
    .map((p, i) => ({ p, i }))
    .filter(({ p }) => p.type === 'content' || p.type === 'data')
  for (let k = 1; k < body.length; k++) {
    const prev = body[k - 1]!
    const cur = body[k]!
    if (cur.p.layout && cur.p.layout === prev.p.layout && cur.i === prev.i + 1) {
      issues.push({
        page: cur.i,
        level: 'error',
        message: `layout ${cur.p.layout} repeats the previous page's; vary the composition`,
      })
    }
  }
  const distinct = new Set(body.map(({ p }) => p.layout).filter(Boolean))
  if (body.length >= 4 && distinct.size < 3) {
    issues.push({
      level: 'warning',
      message: `only ${distinct.size} layout variant(s) across ${body.length} body pages; use at least three`,
    })
  }
}
